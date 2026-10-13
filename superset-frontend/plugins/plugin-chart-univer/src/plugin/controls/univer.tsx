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
import {
  ControlSetItem,
  CustomControlConfig,
} from '@superset-ui/chart-controls';
import { t } from '@superset-ui/core';
import { Button } from '@superset-ui/core/components';
import { WORKBOOK_CONTROL } from '../../types';

/**
 * Holds the spreadsheet saved from the chart itself. The chart writes it with
 * `setControlValue`, so the control only shows a way to start over.
 */
const WorkbookControl = ({
  value,
  onChange,
}: CustomControlConfig<{ value?: string | null }>) => (
  <Button
    buttonSize="small"
    buttonStyle="secondary"
    disabled={!value}
    onClick={() => onChange?.(null)}
  >
    {t('Reset spreadsheet')}
  </Button>
);

export const workbookControlSetItem: ControlSetItem = {
  name: WORKBOOK_CONTROL,
  config: {
    type: WorkbookControl,
    label: t('Spreadsheet'),
    description: t(
      'Edit the spreadsheet in the chart area, then save the chart to keep the changes. Reset removes all edits.',
    ),
    default: null,
    renderTrigger: true,
  },
};

export const liveDataControlSetItem: ControlSetItem = {
  name: 'univer_live_data',
  config: {
    type: 'CheckboxControl',
    label: t('Refresh data on load'),
    description: t(
      'Fill the query result area with fresh data every time the chart loads. Formatting and cells outside that area are kept, but edits to values inside it are replaced. Turn off to keep the sheet exactly as saved.',
    ),
    default: true,
    renderTrigger: true,
  },
};

export const showHeadersControlSetItem: ControlSetItem = {
  name: 'univer_show_headers',
  config: {
    type: 'CheckboxControl',
    label: t('Show row and column headers on dashboards'),
    description: t('Show the A, B, C… and 1, 2, 3… headers outside Explore'),
    default: false,
    renderTrigger: true,
  },
};
