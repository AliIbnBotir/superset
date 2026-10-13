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

/**
 * Everything that touches Univer lives here. Univer needs React 18 while
 * Superset runs React 17, so build.mjs bundles this file
 * together with its own React 18 copy into runtime-dist/. Univer renders into
 * its own React root, so the two React versions never share a tree.
 */
import {
  CommandType,
  createUniver,
  ICommandInfo,
  IWorkbookData,
  LocaleType,
  mergeLocales,
} from '@univerjs/presets';
import {
  FWorkbook,
  UniverSheetsCorePreset,
} from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import '@univerjs/preset-sheets-core/lib/index.css';
import type { MountUniver, WorkbookData } from '../src/runtimeTypes';

export const mountUniver: MountUniver = ({
  container,
  workbook: workbookData,
  editable,
  onChange,
}) => {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.EN_US,
    locales: {
      [LocaleType.EN_US]: mergeLocales(UniverPresetSheetsCoreEnUS),
    },
    presets: [
      UniverSheetsCorePreset({
        container,
        header: editable,
        toolbar: editable,
        formulaBar: editable,
        footer: editable ? undefined : false,
        contextMenu: editable,
        // the editor needs focus handling for typing into cells, but a
        // read-only dashboard chart shouldn't take focus from the page
        disableAutoFocus: editable ? undefined : true,
      }),
    ],
  });
  // the chart's WorkbookData only describes the fields it uses
  const workbook: FWorkbook = univerAPI.createWorkbook(
    workbookData as Partial<IWorkbookData>,
  );
  if (!editable) workbook.setEditable(false);

  const listener =
    editable && onChange
      ? workbook.onCommandExecuted((command: ICommandInfo) => {
          // mutations are what changes the document; formula mutations only
          // carry recalculated values and would otherwise fire on every load
          if (
            command.type === CommandType.MUTATION &&
            !command.id.startsWith('formula.')
          ) {
            onChange();
          }
        })
      : null;

  return {
    save: () => workbook.save() as unknown as WorkbookData,
    dispose: () => {
      listener?.dispose();
      univer.dispose();
    },
  };
};
