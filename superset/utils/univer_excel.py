# Licensed to the Apache Software Foundation (ASF) under one
# or more contributor license agreements.  See the NOTICE file
# distributed with this work for additional information
# regarding copyright ownership.  The ASF licenses this file
# to you under the Apache License, Version 2.0 (the
# "License"); you may not use this file except in compliance
# with the License.  You may obtain a copy of the License at
#
#   http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing,
# software distributed under the License is distributed on an
# "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
# KIND, either express or implied.  See the License for the
# specific language governing permissions and limitations
# under the License.
"""
Excel export of the spreadsheet chart (viz type ``univer_sheet``).

The chart shows the query result in a Univer spreadsheet that users can edit:
rename headers, add formulas, format cells, merge cells, add sheets. Those
edits are saved in ``form_data["univer_workbook"]`` as Univer workbook data.
This module rebuilds the sheet the chart shows from the query result and that
saved state, the same way the frontend does (see
``superset-frontend/plugins/plugin-chart-univer/src/utils/workbook.ts``), and
writes it as an .xlsx file with XlsxWriter, like Superset's other Excel
exports.
"""

from __future__ import annotations

import copy
import io
import math
import re
from datetime import date, datetime
from typing import Any, Optional, Union

import numpy as np
import pandas as pd
import xlsxwriter
from openpyxl.formula.translate import Translator
from xlsxwriter.format import Format
from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.worksheet import Worksheet

from superset.utils import json
from superset.utils.core import GenericDataType

VIZ_TYPE = "univer_sheet"
WORKBOOK_CONTROL = "univer_workbook"
LIVE_DATA_CONTROL = "univer_live_data"

CellValue = Union[str, int, float, bool]
Grid = list[list[Optional[CellValue]]]
UniverWorkbook = dict[str, Any]

HEADER_STYLE_ID = "superset_header"
# cell fields that hold the value, everything else (like style) is kept
DATA_KEYS = ("v", "t", "f", "si", "p")

# pixel widths used to fit columns to their content in a fresh workbook,
# as in the frontend
MIN_COL_WIDTH = 80
MAX_COL_WIDTH = 320
CHAR_WIDTH = 8
COL_PADDING = 16
WIDTH_SAMPLE_ROWS = 100
DEFAULT_COL_WIDTH = 88
DEFAULT_ROW_HEIGHT = 24

# Univer sizes are in pixels, Excel's row heights in points
PT_PER_PX = 0.75

# Univer BorderStyleTypes (from 0) as XlsxWriter border indexes
BORDER_STYLES = [0, 1, 7, 4, 3, 9, 11, 6, 2, 8, 10, 12, 13, 5]
# Univer HorizontalAlign and VerticalAlign, in order from 0
H_ALIGN: list[Optional[str]] = [
    None,
    "left",
    "center",
    "right",
    "justify",
    "justify",
    "distributed",
]
V_ALIGN: list[Optional[str]] = [None, "top", "vcenter", "bottom"]
WRAP = 3
CELL_TYPE_BOOLEAN = 3


def is_univer_chart(form_data: Optional[dict[str, Any]]) -> bool:
    """Whether the form data belongs to a spreadsheet chart."""
    return bool(form_data) and form_data.get("viz_type") == VIZ_TYPE  # type: ignore


def parse_saved_state(raw: Any) -> Optional[dict[str, Any]]:
    """The chart's saved state, or None when nothing (valid) is saved."""
    if not raw:
        return None
    try:
        state = json.loads(raw) if isinstance(raw, str) else raw
    except ValueError:
        return None
    if (
        isinstance(state, dict)
        and state.get("version") == 1
        and (state.get("workbook") or {}).get("sheetOrder")
    ):
        return state
    return None


def _format_temporal(value: Any) -> str:
    timestamp = pd.Timestamp(value)
    if timestamp.tzinfo is not None:
        timestamp = timestamp.tz_convert("UTC").tz_localize(None)
    # drop the time part for whole days
    if timestamp == timestamp.normalize():
        return timestamp.strftime("%Y-%m-%d")
    return timestamp.strftime("%Y-%m-%d %H:%M:%S")


def _temporal_value(value: Any) -> str:
    try:
        if isinstance(value, (int, float)):
            return _format_temporal(pd.Timestamp(value, unit="ms"))
        return _format_temporal(value)
    except (ValueError, TypeError):
        return str(value)


def _float_value(value: float) -> CellValue:
    if not math.isfinite(value):
        return str(value)
    if not value.is_integer():
        # drop floating point noise such as 524.999999999985
        return float(f"{value:.12g}")
    return value


def to_cell_value(
    value: Any, coltype: Optional[GenericDataType] = None
) -> Optional[CellValue]:
    """A query value as the chart shows it in a cell."""
    if isinstance(value, np.generic):
        value = value.item()
    if value is None or value is pd.NaT:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if coltype == GenericDataType.TEMPORAL and isinstance(
        value, (int, float, str, datetime, date)
    ):
        return _temporal_value(value)
    if isinstance(value, float):
        return _float_value(value)
    if isinstance(value, (bool, int, str)):
        return value
    if isinstance(value, (datetime, date)):
        return _format_temporal(value)
    return json.dumps(value, default=str)


def build_grid(df: pd.DataFrame, coltypes: list[GenericDataType]) -> Grid:
    """The query result as a grid: the header row followed by the data rows."""
    columns = list(df.columns)
    types: list[Optional[GenericDataType]] = list(coltypes) + [None] * (
        len(columns) - len(coltypes)
    )
    grid: Grid = [[str(column) for column in columns]]
    for row in df.itertuples(index=False, name=None):
        grid.append([to_cell_value(value, types[i]) for i, value in enumerate(row)])
    return grid


def _grid_size(grid: Grid) -> tuple[int, int]:
    return len(grid), len(grid[0]) if grid else 0


def _first_sheet(workbook: UniverWorkbook) -> Optional[dict[str, Any]]:
    order = workbook.get("sheetOrder") or []
    return (workbook.get("sheets") or {}).get(order[0]) if order else None


def apply_grid(
    workbook: UniverWorkbook,
    grid: Grid,
    previous: Optional[tuple[int, int]],
) -> UniverWorkbook:
    """
    Writes the query result into the top-left corner of the first sheet.
    Styles of existing cells are kept, cells that belonged to the previous
    (bigger) result are cleared. Returns a new workbook.
    """
    result = copy.deepcopy(workbook)
    sheet = _first_sheet(result)
    if sheet is None:
        return result

    rows, cols = _grid_size(grid)
    cell_data: dict[str, Any] = sheet.setdefault("cellData", {})
    max_rows = max(rows, previous[0] if previous else 0)
    max_cols = max(cols, previous[1] if previous else 0)
    has_header_style = HEADER_STYLE_ID in (result.get("styles") or {})

    for r in range(max_rows):
        row: dict[str, Any] = cell_data.setdefault(str(r), {})
        for c in range(max_cols):
            existing = row.get(str(c)) or {}
            cell = {k: v for k, v in existing.items() if k not in DATA_KEYS}
            value = grid[r][c] if r < rows and c < cols else None
            if value is not None:
                cell["v"] = value
            if r == 0 and value is not None and not cell.get("s") and has_header_style:
                cell["s"] = HEADER_STYLE_ID
            if cell:
                row[str(c)] = cell
            else:
                row.pop(str(c), None)
        if not row:
            cell_data.pop(str(r), None)
    return result


def _fit_column_widths(grid: Grid) -> dict[str, dict[str, int]]:
    sample = grid[:WIDTH_SAMPLE_ROWS]
    widths: dict[str, dict[str, int]] = {}
    for c in range(_grid_size(grid)[1]):
        chars = max(len(str(row[c] if row[c] is not None else "")) for row in sample)
        width = min(MAX_COL_WIDTH, max(MIN_COL_WIDTH, chars * CHAR_WIDTH + COL_PADDING))
        widths[str(c)] = {"w": width}
    return widths


def create_workbook_data(grid: Grid) -> UniverWorkbook:
    """The workbook the chart shows when the user hasn't saved any edits."""
    rows, cols = _grid_size(grid)
    workbook = {
        "id": "superset-univer",
        "name": "Superset",
        "sheetOrder": ["data"],
        "styles": {HEADER_STYLE_ID: {"bl": 1, "bg": {"rgb": "#F2F2F2"}}},
        "sheets": {
            "data": {
                "id": "data",
                "name": "Data",
                "rowCount": rows,
                "columnCount": cols,
                "defaultColumnWidth": MIN_COL_WIDTH,
                "freeze": {"xSplit": 0, "ySplit": 1, "startRow": 1, "startColumn": 0},
                "columnData": _fit_column_widths(grid),
                "cellData": {},
            }
        },
    }
    return apply_grid(workbook, grid, None)


def build_workbook(
    grid: Grid, saved: Optional[dict[str, Any]], live_data: bool
) -> UniverWorkbook:
    """The workbook the chart shows, the query result with the user's edits."""
    if saved is None:
        return create_workbook_data(grid)
    # a stripped sheet has no values in the data area, they come from the query
    if not live_data and not saved.get("stripped"):
        return copy.deepcopy(saved["workbook"])

    # the query's labels go in as is, except where the user renamed a column
    sheet = _first_sheet(saved["workbook"]) or {}
    renamed = (sheet.get("cellData") or {}).get("0") or {}
    saved_headers = saved.get("headers") or []
    headers: list[Optional[CellValue]] = []
    for c, label in enumerate(grid[0] if grid else []):
        value = (renamed.get(str(c)) or {}).get("v")
        same_column = c < len(saved_headers) and saved_headers[c] == label
        headers.append(value if value is not None and same_column else label)
    return apply_grid(
        saved["workbook"],
        [headers, *grid[1:]],
        (saved.get("dataRows", 0), saved.get("dataCols", 0)),
    )


def _to_hex(color: Any) -> Optional[str]:
    """'#F2F2F2', '#FFF' or 'rgb(242, 242, 242)' as '#F2F2F2'."""
    rgb = str((color or {}).get("rgb") or "").strip()
    if not rgb:
        return None
    if match := re.match(r"^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)", rgb, re.I):
        return "#" + "".join(f"{min(int(n), 255):02X}" for n in match.groups())
    hex_value = rgb.lstrip("#")
    if len(hex_value) == 3:
        hex_value = "".join(ch * 2 for ch in hex_value)
    if re.fullmatch(r"[0-9a-fA-F]{6}", hex_value):
        return "#" + hex_value.upper()
    return None


def _pick(values: list[Any], index: Any) -> Any:
    if isinstance(index, int) and 0 <= index < len(values):
        return values[index]
    return None


def _font_props(style: dict[str, Any]) -> dict[str, Any]:
    return {
        "font_name": style.get("ff") or None,
        "font_size": style.get("fs") or None,
        "bold": bool(style.get("bl")) or None,
        "italic": bool(style.get("it")) or None,
        "underline": 1 if (style.get("ul") or {}).get("s") else None,
        "font_strikeout": bool((style.get("st") or {}).get("s")) or None,
        "font_color": _to_hex(style.get("cl")),
    }


def _border_props(style: dict[str, Any]) -> dict[str, Any]:
    borders = style.get("bd")
    props: dict[str, Any] = {}
    if not isinstance(borders, dict):
        return props
    for name, key in (("top", "t"), ("right", "r"), ("bottom", "b"), ("left", "l")):
        border = borders.get(key)
        if isinstance(border, dict) and _pick(BORDER_STYLES, border.get("s")):
            props[name] = BORDER_STYLES[border["s"]]
            props[f"{name}_color"] = _to_hex(border.get("cl"))
    return props


def format_props(style: dict[str, Any]) -> dict[str, Any]:
    """A Univer cell style as XlsxWriter format properties."""
    fill = _to_hex(style.get("bg"))
    props = {
        **_font_props(style),
        "bg_color": fill,
        "pattern": 1 if fill else None,
        "align": _pick(H_ALIGN, style.get("ht")),
        "valign": _pick(V_ALIGN, style.get("vt")),
        "text_wrap": True if style.get("tb") == WRAP else None,
        "num_format": (style.get("n") or {}).get("pattern") or None,
        **_border_props(style),
    }
    return {key: value for key, value in props.items() if value is not None}


class _Writer:
    """Writes Univer workbook data with XlsxWriter, sharing equal formats."""

    def __init__(self, book: xlsxwriter.Workbook, styles: dict[str, Any]) -> None:
        self.book = book
        self.styles = styles
        self.formats: dict[str, Optional[Format]] = {}

    def cell_format(self, style: Any) -> Optional[Format]:
        if isinstance(style, str):
            style = self.styles.get(style)
        if not isinstance(style, dict):
            return None
        props = format_props(style)
        key = json.dumps(props, sort_keys=True)
        if key not in self.formats:
            self.formats[key] = self.book.add_format(props) if props else None
        return self.formats[key]

    def write_cell(
        self,
        ws: Worksheet,
        position: tuple[int, int],
        data: dict[str, Any],
        masters: dict[str, tuple[str, str]],
    ) -> None:
        row, col = position
        fmt = self.cell_format(data.get("s"))
        value = _cell_value(data)
        formula = _cell_formula(data, xl_rowcol_to_cell(row, col), masters)
        if formula:
            ws.write_formula(row, col, formula, fmt, "" if value is None else value)
        elif isinstance(value, bool):
            ws.write_boolean(row, col, value, fmt)
        elif isinstance(value, (int, float)):
            ws.write_number(row, col, value, fmt)
        elif value is not None:
            # written as text even when it looks like a formula
            ws.write_string(row, col, str(value), fmt)
        elif fmt is not None:
            ws.write_blank(row, col, None, fmt)

    def add_sheet(self, sheet: dict[str, Any], name: str) -> None:
        ws = self.book.add_worksheet(name)
        _set_layout(ws, sheet)
        _set_dimensions(ws, sheet)

        cell_data = sheet.get("cellData") or {}
        masters = _shared_formulas(cell_data)
        merges = sheet.get("mergeData") or []
        for merge in merges:
            ws.merge_range(
                merge["startRow"],
                merge["startColumn"],
                merge["endRow"],
                merge["endColumn"],
                "",
                self.cell_format(
                    (
                        (cell_data.get(str(merge["startRow"])) or {}).get(
                            str(merge["startColumn"])
                        )
                        or {}
                    ).get("s")
                ),
            )
        hidden = {
            (r, c)
            for merge in merges
            for r in range(merge["startRow"], merge["endRow"] + 1)
            for c in range(merge["startColumn"], merge["endColumn"] + 1)
            if (r, c) != (merge["startRow"], merge["startColumn"])
        }
        for r, row in cell_data.items():
            for c, data in (row or {}).items():
                if isinstance(data, dict) and data and (int(r), int(c)) not in hidden:
                    self.write_cell(ws, (int(r), int(c)), data, masters)


def _rich_text(cell: dict[str, Any]) -> Optional[str]:
    text = ((cell.get("p") or {}).get("body") or {}).get("dataStream")
    if not isinstance(text, str):
        return None
    return re.sub(r"\r?\n$", "", text)


def _cell_value(data: dict[str, Any]) -> Optional[CellValue]:
    value = data.get("v")
    if value is None:
        value = _rich_text(data)
    if data.get("t") == CELL_TYPE_BOOLEAN and value is not None:
        return value is True or value == 1 or str(value).upper() == "TRUE"
    return value


def _shared_formulas(cell_data: dict[str, Any]) -> dict[str, tuple[str, str]]:
    """Shared formula id to its formula and the cell it is written for."""
    masters: dict[str, tuple[str, str]] = {}
    for r, row in cell_data.items():
        for c, cell in (row or {}).items():
            if isinstance(cell, dict) and cell.get("si") and cell.get("f"):
                address = xl_rowcol_to_cell(int(r), int(c))
                masters.setdefault(cell["si"], (cell["f"], address))
    return masters


def _cell_formula(
    data: dict[str, Any], address: str, masters: dict[str, tuple[str, str]]
) -> Optional[str]:
    formula = data.get("f")
    if not formula and data.get("si") in masters:
        master_formula, origin = masters[data["si"]]
        formula = Translator(master_formula, origin=origin).translate_formula(address)
    if not formula:
        return None
    return formula if formula.startswith("=") else f"={formula}"


def _set_layout(ws: Worksheet, sheet: dict[str, Any]) -> None:
    """Default row height, gridlines and frozen panes."""
    ws.set_default_row(
        (sheet.get("defaultRowHeight") or DEFAULT_ROW_HEIGHT) * PT_PER_PX
    )
    if sheet.get("showGridlines") == 0:
        ws.hide_gridlines(2)
    freeze = sheet.get("freeze") or {}
    x_split = max(int(freeze.get("xSplit") or 0), 0)
    y_split = max(int(freeze.get("ySplit") or 0), 0)
    if x_split or y_split:
        ws.freeze_panes(y_split, x_split)


def _set_dimensions(ws: Worksheet, sheet: dict[str, Any]) -> None:
    """Column widths, row heights and hidden rows and columns."""
    default_width = sheet.get("defaultColumnWidth") or DEFAULT_COL_WIDTH
    columns = sheet.get("columnData") or {}
    count = max([sheet.get("columnCount") or 0, *[int(c) + 1 for c in columns]])
    for c in range(count):
        column = columns.get(str(c)) or {}
        ws.set_column_pixels(
            c,
            c,
            column.get("w") or default_width,
            None,
            {"hidden": True} if column.get("hd") else None,
        )
    for r, row in (sheet.get("rowData") or {}).items():
        if row and (row.get("h") or row.get("hd")):
            ws.set_row_pixels(
                int(r),
                row.get("h") or sheet.get("defaultRowHeight") or DEFAULT_ROW_HEIGHT,
                None,
                {"hidden": True} if row.get("hd") else None,
            )


def _sheet_names(sheets: list[dict[str, Any]]) -> list[str]:
    """Sheet names Excel accepts: no []:*?/\\, at most 31 characters, unique."""
    names: list[str] = []
    for i, sheet in enumerate(sheets):
        base = re.sub(r"[\[\]:*?/\\]", "_", str(sheet.get("name") or ""))
        base = base.strip("'")[:31] or f"Sheet{i + 1}"
        name, n = base, 1
        while name.lower() in (existing.lower() for existing in names):
            n += 1
            name = f"{base[: 31 - len(str(n)) - 1]}_{n}"
        names.append(name)
    return names


def workbook_to_excel(workbook: UniverWorkbook) -> bytes:
    """Writes Univer workbook data as an .xlsx file."""
    output = io.BytesIO()
    book = xlsxwriter.Workbook(
        output,
        {
            "in_memory": True,
            # text in cells stays text, formulas come only from formula cells
            "strings_to_formulas": False,
            "strings_to_numbers": False,
            "strings_to_urls": False,
        },
    )
    writer = _Writer(book, workbook.get("styles") or {})
    all_sheets = workbook.get("sheets") or {}
    sheets = [
        all_sheets[sheet_id]
        for sheet_id in workbook.get("sheetOrder") or []
        if all_sheets.get(sheet_id)
    ]
    for sheet, name in zip(sheets, _sheet_names(sheets), strict=True):
        writer.add_sheet(sheet, name)
    if not sheets:
        book.add_worksheet("Data")
    _hide_sheets(book, sheets)
    book.close()
    return output.getvalue()


def _hide_sheets(book: xlsxwriter.Workbook, sheets: list[dict[str, Any]]) -> None:
    """Hides the sheets hidden in the chart, Excel needs one to stay visible."""
    visible = [i for i, sheet in enumerate(sheets) if not sheet.get("hidden")]
    if not visible:
        return
    worksheets = book.worksheets()
    worksheets[visible[0]].activate()
    for i, sheet in enumerate(sheets):
        if sheet.get("hidden"):
            worksheets[i].hide()


def df_to_univer_excel(
    df: pd.DataFrame,
    coltypes: list[GenericDataType],
    form_data: dict[str, Any],
) -> bytes:
    """
    The spreadsheet chart as an .xlsx file: the query result in ``df`` (with
    column labels as the chart shows them) and the user's saved edits.
    """
    saved = parse_saved_state(form_data.get(WORKBOOK_CONTROL))
    live_data = form_data.get(LIVE_DATA_CONTROL)
    workbook = build_workbook(
        build_grid(df, coltypes),
        saved,
        True if live_data is None else bool(live_data),
    )
    return workbook_to_excel(workbook)
