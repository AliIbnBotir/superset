<!--
Licensed to the Apache Software Foundation (ASF) under one
or more contributor license agreements.  See the NOTICE file
distributed with this work for additional information
regarding copyright ownership.  The ASF licenses this file
to you under the Apache License, Version 2.0 (the
"License"); you may not use this file except in compliance
with the License.  You may obtain a copy of the License at

  http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing,
software distributed under the License is distributed on an
"AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
KIND, either express or implied.  See the License for the
specific language governing permissions and limitations
under the License.
-->

## @superset-ui/plugin-chart-univer

An Excel-like spreadsheet chart built on [Univer](https://github.com/dream-num/univer) (open source core only).

- **Explore:** the query result is loaded into a full spreadsheet editor (toolbar, formula bar,
  sheet tabs, context menu). Edits are written to the `univer_workbook` control and saved with the chart.
- **Dashboards:** the saved sheet is shown read-only, without toolbars or editors.

### Data refresh

With **Refresh data on load** on (the default), the query result is written into the top-left corner
of the first sheet every time the chart loads. Formatting, formulas and anything outside that area are
kept, and the query values are not stored in the chart. Put your own formulas to the right of the data
or on another sheet, because the data area grows when the query returns more rows or columns.

With it off, the sheet is shown exactly as it was saved.

### Usage

```js
import { UniverChartPlugin } from '@superset-ui/plugin-chart-univer';

new UniverChartPlugin().configure({ key: 'univer_sheet' }).register();
```

### How Univer is built

Univer needs React 18 and Superset runs React 17, so Univer is not a normal dependency of this plugin.
`runtime/` is a separate npm project (its own `package.json` and lockfile, not a Superset workspace)
with Univer and React 18. `runtime/build.mjs` installs it if needed and bundles `runtime/index.ts`
into `runtime-dist/` (git-ignored). Univer renders into its own React root, so the two React
versions never share a component tree.

Superset's `build`, `build-dev`, `dev` and `dev-server` scripts run this first. To run it by hand:

```bash
npm run build-univer-runtime
```

To upgrade Univer, change the versions in `runtime/package.json`, run `npm install` in `runtime/`,
and rebuild.
