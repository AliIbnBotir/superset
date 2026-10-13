/**
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */
import { cloneDeep, omit } from 'lodash';
import { DataRecord, GenericDataType } from '@superset-ui/core';
import { CellData, CellValue, WorkbookData } from '../runtimeTypes';

/** Query result as a grid: the header row followed by the data rows. */
export type Grid = (CellValue | null)[][];

/** What the chart stores in `form_data.univer_workbook` (as a JSON string). */
export interface UniverSavedState {
  version: 1;
  /** Size of the query result area (header row included) when it was saved. */
  dataRows: number;
  dataCols: number;
  /** true when query values were left out of `workbook` to save space */
  stripped: boolean;
  /**
   * The query's column labels when it was saved. A header cell in a stripped
   * workbook that still has a value is one the user renamed, it is kept while
   * the query's label for that column stays the same.
   */
  headers?: (CellValue | null)[];
  workbook: WorkbookData;
}

export const HEADER_STYLE_ID = 'superset_header';
const WORKBOOK_ID = 'superset-univer';
const SHEET_ID = 'data';
// cell fields that hold the value, everything else (like style) is kept
const DATA_KEYS = ['v', 't', 'f', 'si', 'p'];

// pixel widths used to fit columns to their content in a fresh workbook
export const MIN_COL_WIDTH = 80;
const MAX_COL_WIDTH = 320;
const CHAR_WIDTH = 8;
const COL_PADDING = 16;
const WIDTH_SAMPLE_ROWS = 100;
// Univer's defaults for the width of the 1, 2, 3… row header and of a column
const ROW_HEADER_WIDTH = 46;
const DEFAULT_COL_WIDTH = 88;
// room for Univer's vertical scrollbar, which is drawn over the cells
const SCROLLBAR_WIDTH = 12;

export function parseSavedState(raw?: string | null): UniverSavedState | null {
  if (!raw) return null;
  try {
    const state = JSON.parse(raw);
    if (state?.version === 1 && state.workbook?.sheetOrder?.length) {
      return state;
    }
  } catch {
    // fall through: a broken value is treated as "nothing saved"
  }
  return null;
}

function formatTemporal(value: number | string): string {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  const iso = date.toISOString();
  // drop the time part for whole days, Superset sends timestamps in UTC
  return iso.endsWith('T00:00:00.000Z')
    ? iso.slice(0, 10)
    : iso.slice(0, 19).replace('T', ' ');
}

export function toCellValue(
  value: unknown,
  coltype?: GenericDataType,
): CellValue | null {
  if (value === null || value === undefined) return null;
  if (
    coltype === GenericDataType.Temporal &&
    (typeof value === 'number' || typeof value === 'string')
  ) {
    return formatTemporal(value);
  }
  if (
    typeof value === 'number' &&
    Number.isFinite(value) &&
    !Number.isInteger(value)
  ) {
    // drop floating point noise such as 524.999999999985, which the
    // spreadsheet would otherwise show with all its digits
    return Number(value.toPrecision(12));
  }
  if (
    typeof value === 'number' ||
    typeof value === 'string' ||
    typeof value === 'boolean'
  ) {
    return value;
  }
  return JSON.stringify(value);
}

export function buildGrid(
  data: DataRecord[],
  colnames: string[],
  coltypes: GenericDataType[],
  labels: Record<string, string> = {},
): Grid {
  return [
    colnames.map(col => labels[col] ?? col),
    ...data.map(row =>
      colnames.map((col, i) => toCellValue(row[col], coltypes[i])),
    ),
  ];
}

function gridSize(grid: Grid) {
  return { rows: grid.length, cols: grid[0]?.length ?? 0 };
}

/**
 * Writes the query result into the top-left corner of the first sheet.
 * Styles of existing cells are kept, so formatting survives a data refresh.
 * Cells that belonged to the previous (bigger) result are cleared.
 * Returns a new object, the input is not modified.
 */
export function applyGrid(
  workbook: WorkbookData,
  grid: Grid,
  previous: { rows: number; cols: number } | null,
): WorkbookData {
  const result = cloneDeep(workbook);
  const sheetId = result.sheetOrder?.[0];
  const sheet = sheetId ? result.sheets?.[sheetId] : undefined;
  if (!sheet) return result;

  const { rows, cols } = gridSize(grid);
  const cellData = (sheet.cellData ??= {});
  const maxRows = Math.max(rows, previous?.rows ?? 0);
  const maxCols = Math.max(cols, previous?.cols ?? 0);
  const hasHeaderStyle = !!result.styles?.[HEADER_STYLE_ID];

  for (let r = 0; r < maxRows; r += 1) {
    const row = (cellData[r] ??= {});
    for (let c = 0; c < maxCols; c += 1) {
      // a data cell can't keep a formula or rich text from an older version
      const cell: CellData = omit(row[c] ?? {}, DATA_KEYS);
      const value = r < rows && c < cols ? grid[r][c] : null;
      if (value !== null) cell.v = value;
      if (r === 0 && value !== null && !cell.s && hasHeaderStyle) {
        cell.s = HEADER_STYLE_ID;
      }
      if (Object.keys(cell).length) row[c] = cell;
      else delete row[c];
    }
    if (!Object.keys(row).length) delete cellData[r];
  }

  return result;
}

/**
 * Sizes the first sheet to its content (query result, formulas, styled
 * cells) instead of an empty grid around it. Users add rows and columns
 * themselves, empty ones they added are dropped the next time the chart loads.
 */
export function fitSheetSize(workbook: WorkbookData, grid: Grid): WorkbookData {
  const sheetId = workbook.sheetOrder?.[0];
  const sheet = sheetId ? workbook.sheets?.[sheetId] : undefined;
  if (!sheetId || !sheet) return workbook;

  let { rows, cols } = gridSize(grid);
  Object.entries(sheet.cellData ?? {}).forEach(([r, row]) => {
    Object.entries(row).forEach(([c, cell]) => {
      if (!cell || !Object.keys(cell).length) return;
      rows = Math.max(rows, Number(r) + 1);
      cols = Math.max(cols, Number(c) + 1);
    });
  });
  return {
    ...workbook,
    sheets: {
      ...workbook.sheets,
      [sheetId]: {
        ...sheet,
        rowCount: Math.max(rows, 1),
        columnCount: Math.max(cols, 1),
      },
    },
  };
}

/** Width of a column on the first sheet, when it has one set. */
export function columnWidth(
  workbook: WorkbookData,
  column: number,
): number | undefined {
  const sheetId = workbook.sheetOrder?.[0];
  return sheetId
    ? workbook.sheets?.[sheetId]?.columnData?.[column]?.w
    : undefined;
}

/** Column index to the width it had before stretching and the width given. */
export type StretchedColumns = Record<number, { base: number; w: number }>;

/**
 * Widens the columns of the first sheet so together they fill `width`
 * pixels, keeping their proportions. Columns that already overflow are left
 * as they are, the sheet scrolls then.
 */
export function stretchColumns(
  workbook: WorkbookData,
  width: number,
): { workbook: WorkbookData; stretched: StretchedColumns } {
  const sheetId = workbook.sheetOrder?.[0];
  const sheet = sheetId ? workbook.sheets?.[sheetId] : undefined;
  const stretched: StretchedColumns = {};
  if (!sheetId || !sheet) return { workbook, stretched };

  const defaultWidth = sheet.defaultColumnWidth ?? DEFAULT_COL_WIDTH;
  const columnData = { ...sheet.columnData };
  const visible: number[] = [];
  let total = 0;
  for (let c = 0; c < (sheet.columnCount ?? 0); c += 1) {
    if (!columnData[c]?.hd) {
      visible.push(c);
      total += columnData[c]?.w ?? defaultWidth;
    }
  }
  const headerWidth = sheet.rowHeader?.hidden
    ? 0
    : (sheet.rowHeader?.width ?? ROW_HEADER_WIDTH);
  const scale = (width - headerWidth - SCROLLBAR_WIDTH) / total;
  if (!visible.length || !(scale > 1)) return { workbook, stretched };

  visible.forEach(c => {
    const base = columnData[c]?.w ?? defaultWidth;
    // whole pixels, so the columns never add up to more than the view
    const w = Math.floor(base * scale);
    stretched[c] = { base, w };
    columnData[c] = { ...columnData[c], w };
  });
  return {
    workbook: {
      ...workbook,
      sheets: { ...workbook.sheets, [sheetId]: { ...sheet, columnData } },
    },
    stretched,
  };
}

/**
 * Undoes `stretchColumns` before saving, so the saved widths don't depend on
 * the size of the view. A column resized by the user keeps its new width.
 */
export function unstretchColumns(
  workbook: WorkbookData,
  stretched: StretchedColumns,
): WorkbookData {
  const sheetId = workbook.sheetOrder?.[0];
  const sheet = sheetId ? workbook.sheets?.[sheetId] : undefined;
  if (!sheetId || !sheet?.columnData) return workbook;

  const columnData = { ...sheet.columnData };
  Object.entries(stretched).forEach(([c, { base, w }]) => {
    const column = columnData[Number(c)];
    if (column?.w === w) columnData[Number(c)] = { ...column, w: base };
  });
  return {
    ...workbook,
    sheets: { ...workbook.sheets, [sheetId]: { ...sheet, columnData } },
  };
}

function fitColumnWidths(grid: Grid): Record<number, { w: number }> {
  const widths: Record<number, { w: number }> = {};
  const sample = grid.slice(0, WIDTH_SAMPLE_ROWS);
  for (let c = 0; c < gridSize(grid).cols; c += 1) {
    const chars = Math.max(...sample.map(row => String(row[c] ?? '').length));
    widths[c] = {
      w: Math.min(
        MAX_COL_WIDTH,
        Math.max(MIN_COL_WIDTH, chars * CHAR_WIDTH + COL_PADDING),
      ),
    };
  }
  return widths;
}

export function createWorkbookData(grid: Grid): WorkbookData {
  const { rows, cols } = gridSize(grid);
  return applyGrid(
    {
      id: WORKBOOK_ID,
      name: 'Superset',
      sheetOrder: [SHEET_ID],
      styles: {
        // spreadsheet content, not app UI, so it doesn't follow the theme
        // eslint-disable-next-line theme-colors/no-literal-colors
        [HEADER_STYLE_ID]: { bl: 1, bg: { rgb: '#F2F2F2' } },
      },
      sheets: {
        [SHEET_ID]: {
          id: SHEET_ID,
          name: 'Data',
          rowCount: rows,
          columnCount: cols,
          defaultColumnWidth: MIN_COL_WIDTH,
          // keep the header row visible while scrolling
          freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 },
          columnData: fitColumnWidths(grid),
          cellData: {},
        },
      },
    },
    grid,
    null,
  );
}

/**
 * Removes query values from the data area before saving. They are filled in
 * again from fresh query results when the chart loads, which keeps the saved
 * chart small. Formatting of those cells is kept, and so are headers the user
 * renamed.
 */
export function stripGrid(workbook: WorkbookData, grid: Grid): WorkbookData {
  const { rows, cols } = gridSize(grid);
  const result = applyGrid(workbook, [], { rows, cols });
  const sheetId = workbook.sheetOrder?.[0];
  const header = sheetId ? workbook.sheets?.[sheetId]?.cellData?.[0] : null;
  const cellData = sheetId ? result.sheets?.[sheetId]?.cellData : null;
  if (!header || !cellData) return result;

  for (let c = 0; c < cols; c += 1) {
    const value = header[c]?.v;
    if (value !== undefined && value !== null && value !== grid[0][c]) {
      cellData[0] = { ...cellData[0], [c]: { ...cellData[0]?.[c], v: value } };
    }
  }
  return result;
}

/** Hides the A, B, C… and 1, 2, 3… headers on every sheet. */
export function hideHeaders(workbook: WorkbookData): WorkbookData {
  const sheets = Object.fromEntries(
    Object.entries(workbook.sheets ?? {}).map(([id, sheet]) => [
      id,
      {
        ...sheet,
        rowHeader: { width: 0, hidden: 1 },
        columnHeader: { height: 0, hidden: 1 },
      },
    ]),
  );
  return { ...workbook, sheets };
}

export function buildWorkbook(
  grid: Grid,
  saved: UniverSavedState | null,
  liveData: boolean,
): WorkbookData {
  if (!saved) return createWorkbookData(grid);
  // a stripped sheet has no values in the data area, they must come from the query
  if (!liveData && !saved.stripped) {
    return fitSheetSize(cloneDeep(saved.workbook), []);
  }
  // the query's labels go in as is, except where the user renamed a column
  const renamed =
    saved.workbook.sheets?.[saved.workbook.sheetOrder?.[0] ?? '']
      ?.cellData?.[0];
  const headers = (grid[0] ?? []).map((label, c) => {
    const value = renamed?.[c]?.v;
    return value !== undefined && value !== null && saved.headers?.[c] === label
      ? value
      : label;
  });
  return fitSheetSize(
    applyGrid(saved.workbook, [headers, ...grid.slice(1)], {
      rows: saved.dataRows,
      cols: saved.dataCols,
    }),
    grid,
  );
}

export function serializeState(
  workbook: WorkbookData,
  grid: Grid,
  liveData: boolean,
): string {
  const { rows, cols } = gridSize(grid);
  const state: UniverSavedState = {
    version: 1,
    dataRows: rows,
    dataCols: cols,
    stripped: liveData,
    headers: grid[0] ?? [],
    workbook: liveData ? stripGrid(workbook, grid) : workbook,
  };
  return JSON.stringify(state);
}
