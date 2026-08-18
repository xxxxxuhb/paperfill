# Changelog

本仓库按 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 记录公开 API 变化。版本遵循 semver；`0.x` 仍可能有破坏性调整。

## 0.2.0 — 2026-08-18

全部包对齐 `0.2.0`：`paperfill`、`@paperfill/kernel`、`docx`、`xlsx`、`pdf`、`ui`、`vue`、`react`。

### Added

- `createKernel({ markers: { start, end } })`：自定义标记起止符，默认仍是 `{{` / `}}`。`load` 写入 `definition.markers`（默认省略）；`hydrate` 以 definition 为准。`DocxAdapter` / `XlsxAdapter` / `mountDocxPreview` 接受同一份 `markers`；`kernel.getMarkers()` 给预览用。
- 可选 `plugins` 钩子，**不替代** `validators` / `formatters` / `subscribe` / `mountField`。同一对象可当 `PaperfillPlugin` 四处传入，未实现的方法跳过。
  - kernel：`afterDiscover` → `afterHydrate` → `validators`；`formatters` → `beforeExport` → `bind` → `afterExport` → `exported`
  - `mountDocxPreview`：`afterHtml` / `afterExpand` / `afterSlots`（内置铺满仍先跑；可选 `docxFitHostPlugin`）
  - `mountXlsxPreview`：`afterSheets` / `afterTable`
  - `exportFilledDocument`：`afterPdfHtml`

## 0.1.5 — 2026-08-17

伞包 `paperfill@0.1.5`，`@paperfill/xlsx@0.1.1`，`@paperfill/pdf@0.1.1`。

### Changed

- Excel 预览默认裁掉末尾幽灵行/列（ExcelJS `rowCount` 空行），中间空行保留；导出 `.xlsx` 不改。PDF 的 xlsx 渲染同样裁。

## 0.1.4 — 2026-08-17

伞包 `paperfill@0.1.4`，`@paperfill/docx@0.1.2`。

### Changed

- `mountDocxPreview` 默认 `ignoreWidth` / `ignoreHeight: true`，渲染后把 `section.docx` 拉成 `width: 100%`，避免 A4 `595.3pt` 留白。要页框时传 `render: { inWrapper: true, ignoreWidth: false, ignoreHeight: false }`。

## 0.1.3 — 2026-08-17

伞包 `paperfill@0.1.3`，`@paperfill/docx@0.1.1`。

### Changed

- `mountDocxPreview` 默认 `inWrapper: false`：正文渲进传入的 `container`，不再套 `docx-preview` 灰底页框。需要页框时传 `render: { inWrapper: true }`。可另传 `styleContainer`。

## 0.1.2 — 2026-08-16

伞包 `paperfill@0.1.2`，内核 `@paperfill/kernel@0.1.1`。

### Added

- `Field.outputFormat` / `FieldColumn.outputFormat`：导出时格式化日期、数字、选项文案；`data` 仍存规范值。
- `createKernel({ formatters })`：按名字注册自定义格式（如金额大写）。
- `getExportData()` / `formatData` / `formatFieldValue`：查看或复用导出用数据。
- npm 包内 README（安装、场景、前后端配合）。

### Changed

- 文档标记约定改为只写锚点：`{{name}}` / `{{items.qty}}`。`type`、`label`、`options`、`outputFormat` 都在 definition。
- `{{name:type}}` 仍能解析，仅作第一次 `load` 的类型提示；`hydrate` 以 definition 为准。

### Notes

- `outputFormat` 无默认值；不写则导出与 `data` 相同。
- 内置格式：`YYYY年MM月DD日` / `DD/MM/YYYY` / `DD日MM月YYYY年`、`#,##0.00`、`label`。

## 0.1.0

### Added

- Headless `createKernel`：`load` / `hydrate` / `setValue` / `export`，以及字段增删改。
- `@paperfill/docx`：`DocxAdapter`、`mountDocxPreview`、锚点写回、image 嵌图、明细表扩行。
- `@paperfill/xlsx`：`XlsxAdapter`、`mountXlsxPreview`、单元格颜色预览、锚点写回、image 嵌图。
- `@paperfill/ui`：原生 `createField` / `mountField` / `nativeFieldMounter`。
- `@paperfill/vue` / `@paperfill/react`：`useContractKit` 订阅快照。
- 可扩展校验：`Field.rules` + `createKernel({ validators })`。
- 表格行：`insertRow` / `removeRow`；空表预览与导出保留一行占位。
- Persist 约定：`toPersistBundle` / `hydrateFromBundle` / `snapshotKernel`。
- 伞包 `paperfill` 子路径：`/docx` `/xlsx` `/ui` `/vue` `/react` `/pdf`。

### Notes

- `export` bind 时未写入 `data` 的扁平标记会被清空（与预览空槽一致）；`replaceMarkers` 默认仍保留缺失标记。
- Word / Excel 预览版式受 `docx-preview` / ExcelJS 限制，见 README。
