# paperfill 接口文档

面向接入方的公共 API。实现细节与时序见 [architecture.md](./architecture.md)。

## 安装与入口

```bash
npm i paperfill
```

| 导入路径 | 内容 |
|----------|------|
| `paperfill` | **全部 API**：`createKernel`、Docx / Xlsx / Html adapter 与 preview、`nativeFieldMounter`、`exportFilledDocument`、类型与 persist |
| `paperfill/ui/style.css` | 原生字段样式 |
| `paperfill/xlsx/style.css` | Excel 预览表格骨架样式 |
| `paperfill/vue` | `useContractKit`（需 Vue 3 peer） |
| `paperfill/react` | `useContractKit`（需 React ≥18 peer） |

接入方只装 `paperfill`，业务代码只从 `paperfill` 取 API。仓库内 `@paperfill/*` 是实现拆包，不必再单独安装。子路径仅用于 **CSS** 和 **Vue/React 封装**。

---

## 标记语法

模板正文中的占位符：

```text
{{name}}
{{items.col}}
{{items.$index}}
```

| 项 | 说明 |
|----|------|
| `name` | 字段名（data / definition 的 key） |
| `items.col` | **循环明细表**列标记：表字段名 + `.` + 列名 |
| `items.$index` | 可选；导出/预览时替换为 1-based 行号 |

文档里只放锚点名字。`type` / `label` / `required` / `options` / `outputFormat` / 表 `columns` 都写在 `TemplateDefinition`。

默认起止符是 `{{` / `}}`。可换成一对字符串（如 `[[` / `]]`、`${` / `}`），不要用 `{` / `}` 或 `[` / `]` 这种易和正文撞车的单字符。创建 kernel 时传入；`load` 会写入 `definition.markers`（默认双括号省略该字段）。`hydrate` 以 definition 为准。`mountDocxPreview` 传同一份 `markers: kernel.getMarkers()`。

```ts
const markers = { start: '[[', end: ']]' }
const kernel = createKernel({
  adapter: new DocxAdapter({ markers }),
  markers,
})
```

`{{name:type}}` 仍能解析，仅作第一次 `load` 的类型提示；`hydrate` 之后以 definition 为准。

> `image`：data 为 data URL（`data:image/png;base64,...`）；`bind` / `export` 会嵌入 Word drawing / Excel 单元格图。  
> `display`：纯展示，data 仍可 `setValue` / 导出替换，预览不渲染输入框（适合联动结果、系统填入的日期等）。

### 循环明细表

模板中保留**一行**带 `{{items.*}}` 的行作为行模板；表头/合计用普通扁平标记。

```ts
// data
{ items: [{ name: '苹果', qty: 10 }, { name: '橙', qty: 5 }], amount: 100 }

// definition 字段（discover 自动聚合成 type: 'table'）
{
  name: 'items',
  type: 'table',
  columns: [
    { name: 'name', type: 'text', label: '货物名称', required: true },
    { name: 'qty', type: 'number', label: '数量' },
  ],
  anchor: { kind: 'marker', name: 'items' }
}
```

`setValue` 支持：

- `items` — 整表数组  
- `items.0.qty` — 单元格  

也可用一等公民命令：

```ts
await kernel.dispatch({ type: 'insertRow', table: 'items', row: { name: '苹果' } })
await kernel.dispatch({ type: 'insertRow', table: 'items', index: 0 })
await kernel.dispatch({ type: 'removeRow', table: 'items', index: 0 })
```

导出 / 预览按数组长度克隆模板行；**空表保留一行空白占位**（不删行模板），便于继续填。校验：`table.required` → 非空数组；列 `required` → `items.i.col`。

辅助函数（`paperfill`）：

| 函数 | 说明 |
|------|------|
| `parseMarkers(text, markers?)` | 解析去重后的标记列表 |
| `aggregateMarkerFields(markers)` | 将 `items.col` 聚成一个 `table` 字段 + columns |
| `splitByMarkers(text, markers?)` | 拆成 `text` / `field` 段（预览挂控件用） |
| `replaceMarkers(text, data, options?, markers?)` | 用 data 替换扁平标记；`options.missing` 为 `blank` 时清空未写入的标记 |
| `replaceRowMarkers(text, table, row, index, markers?)` | 替换一行内的 `table.col` / `$index` |
| `createMarkerSyntax(markers?)` | 按起止符生成 wrap / regex / contains |
| `rowsForExpand(value)` | 空表 → `[{}]` 占位行 |
| `toPersistBundle(kernel)` | `{ file, definition, data }` |
| `hydrateFromBundle(bundle)` | 生成 `hydrate` 命令 |
| `snapshotKernel(kernel)` | UI 只读快照 |
| `stringifyFieldValue(value)` | 把值转成字符串（无 outputFormat） |
| `formatFieldValue(value, field, ctx?)` | 按 `outputFormat` / 自定义 formatter 转展示值 |
| `formatData(definition, data, formatters?)` | 生成导出用 data（`export` / `getExportData` 内部会调；之后再跑 `beforeExport`） |

---

## 核心类型

### `TemplateDefinition`

```ts
interface TemplateDefinition {
  version: 1
  source: { kind: 'docx' | 'xlsx' | 'html' | 'custom'; hash: string }
  fields: Field[]
  markers?: { start: string; end: string } // 省略 = {{ }}
}
```

### `Field`

```ts
interface Field {
  id: string
  name: string
  type: FieldType
  label?: string
  required?: boolean
  options?: { value: string; label: string }[]
  columns?: FieldColumn[] // type === 'table'
  rules?: FieldRules
  outputFormat?: string
  anchor: Anchor
}

interface FieldRules {
  min?: number
  max?: number
  minLength?: number
  maxLength?: number
  pattern?: string
  dateFormat?: boolean // YYYY-MM-DD
}

interface FieldColumn {
  name: string
  type: FieldType
  label?: string
  required?: boolean
  options?: FieldOption[]
  rules?: FieldRules
  outputFormat?: string
}
```

常见 `anchor`：

- Word：`{ kind: 'marker', name: string }`
- Excel：`{ kind: 'cell', sheet: string, address: string }`

### `Source`

```ts
type Source =
  | { kind: 'docx' | 'xlsx'; buffer: Uint8Array; hash?: string }
  | { kind: 'html'; html: string; hash?: string }
  | { kind: 'html'; buffer: Uint8Array; hash?: string }
  | { kind: 'custom'; data: unknown; hash?: string }
```

HTML `buffer` 按 UTF-8 解码。`custom.data` 由自定义 adapter 解释；若它不能稳定
`JSON.stringify`，调用方必须传 `hash`。

### `ValidationResult`

```ts
interface ValidationResult {
  ok: boolean
  issues: { path: string; message: string }[]
}
```

当前规则：

1. `required` 且值为 `undefined` / `null` / `''` →「必填」
2. `select` 有值但不在 `options` →「不在选项中」
3. `table.required` 且数组为空 →「必填」；列 `required` → `items.0.name` 等路径
4. `date` 有值且不是 `YYYY-MM-DD`
5. `rules`：`min`/`max`、`minLength`/`maxLength`、`pattern`、`dateFormat`
6. `createKernel({ validators })`：自定义 / 跨字段，返回 `{ path, message }` 或数组

### `outputFormat`

`data` 始终存规范值（日期 `YYYY-MM-DD`、选项 `value`、数字为 number）。编辑器用控件自己的显示格式；`export` / `getExportData()` 按 `outputFormat` 写成文档字符串。

| `outputFormat` | 适用 | 例 |
|----------------|------|----|
| `YYYY年MM月DD日` / `DD/MM/YYYY` / `DD日MM月YYYY年` | `date` | `2026-08-16` → `16日08月2026年` |
| `#,##0.00` / `0` | `number` | `15998` → `15,998.00` |
| `label` | `select` / `multiselect` | `wire` → `电汇` |
| 自定义名 | 任意 | `createKernel({ formatters: { amountCn } })` |

表格列用 `columns[].outputFormat`。图片字段不会被格式化。

### `FormSchema`

```ts
interface FormSchema {
  fields: FormSchemaField[] // 含 value、required、label、options…
}
```

用于渲染表单控件。

---

## Kernel

```ts
import { createKernel, DocxAdapter } from 'paperfill'

const kernel = createKernel({
  adapter: new DocxAdapter(),
  formatters: {
    amountCn: ({ value }) => String(value), // outputFormat: 'amountCn'
  },
  // markers: { start: '[[', end: ']]' },
  // plugins: [stampPlugin],
})
```

### 方法一览

| 方法 | 返回 | 说明 |
|------|------|------|
| `dispatch(command)` | `Promise<DispatchResult>` | **唯一写入口** |
| `getState()` | `KernelState` | definition + data + source + validation 快照 |
| `getDefinition()` | `TemplateDefinition \| null` | 字段定义 |
| `getData()` | `Record<string, unknown>` | 填写值（规范值，未套 `outputFormat`） |
| `getExportData()` | `Record<string, unknown>` | 套过 `outputFormat` / `formatters` / `beforeExport` 的导出值 |
| `getFormSchema()` | `FormSchema` | 表单字段（含当前 value） |
| `getView()` | `ViewModel` | `{ id, label, value }[]` |
| `getPreview()` | `PreviewModel \| null` | adapter 预览结构 |
| `getSource()` | `Source \| null` | 原始文件 |
| `getMarkers()` | `MarkerDelimiters` | 当前起止符（默认 `{{` / `}}`） |
| `validate()` | `ValidationResult` | 立即重算校验 |
| `can(command)` | `boolean` | 当前是否可执行该命令 |
| `subscribe(listener)` | `() => void` | 订阅事件；返回取消函数 |
| `setViewport(port)` | `void` | 可选：对接预览选区 / 高亮 |

`DispatchResult`：

```ts
type DispatchResult =
  | { type: 'ok' }
  | {
      type: 'exported'
      buffer: Uint8Array
      format: 'docx' | 'xlsx' | 'html' | 'custom'
    }
```

### `dispatch` 命令

| Command | 作用 | 典型返回 |
|---------|------|----------|
| `{ type: 'load', source }` | 加载文件并扫标记生成 definition | `ok` |
| `{ type: 'hydrate', source, definition, data? }` | 用已发布 definition + data 恢复 | `ok` |
| `{ type: 'insertField', field }` | 新增字段 | `ok` |
| `{ type: 'updateField', id, patch }` | 改 label / required / options 等 | `ok` |
| `{ type: 'removeField', id }` | 删除字段 | `ok` |
| `{ type: 'setValue', path, value }` | 写单个字段 | `ok` |
| `{ type: 'setData', data }` | 整表替换 data | `ok` |
| `{ type: 'resetData' }` | 清空 data | `ok` |
| `{ type: 'insertRow', table, index?, row? }` | 明细表增行（`index` 缺省为末尾） | `ok` |
| `{ type: 'removeRow', table, index }` | 明细表删行 | `ok` |
| `{ type: 'export', format? }` | bind 后导出新文件（不改原 buffer） | `exported` |

说明：

- `load`：适合「只有文件」；业务 options / 必填需再 hydrate 或 `updateField`。
- `hydrate`：已发布模板主路径；`source.kind` 须与 adapter 一致。见下方「hydrate 三件套」。
- `export`：先按 `outputFormat` 生成导出值，再跑 `beforeExport`，再在**副本**上替换标记并扩行/嵌图；`getData()` 仍是规范值；未写入 `data` 的扁平标记导出时清空；`getSource().buffer` 不变。`afterExport` 可改返回的 buffer，然后才发 `exported`。

### 事件 `KernelEvent`

`subscribe` 收到的事件（多数写操作后还会再发一次 `state-changed`）：

| `type` | 含义 |
|--------|------|
| `state-changed` | 状态有变更 |
| `data-changed` | data 变了 |
| `validated` | 带 `result: ValidationResult` |
| `field-inserted` / `field-updated` / `field-removed` | definition 字段变更 |
| `exported` | 导出完成（含 format / bytes） |

### `can(command)` 规则摘要

| 命令 | 条件 |
|------|------|
| `load` / `hydrate` | `source.kind === adapter.kind` |
| 其余写命令 / `export` | 已有 `definition` 且已有 `source` |

### `plugins`

可选钩子数组，**不替代** `validators` / `formatters` / `subscribe`。缺方法就跳过，按数组顺序同步执行。同一对象可当 `PaperfillPlugin` 传给 kernel、预览、PDF，用不到的钩子忽略。

| 钩子 | 时机 | 返回 |
|------|------|------|
| `afterDiscover(fields, { source })` | `load` 扫完标记、写入 definition / `insertAnchor` 之前 | `DiscoveredField[]` 替换列表 |
| `afterHydrate({ definition, data, source })` | `load` / `hydrate` 状态已写入之后；随后跑 `validators` | `{ data }` 替换填写值 |
| `beforeExport(data, ctx)` | `formatData` / `formatters` 之后、`bind` 之前；`getExportData()` 同样走这里 | 替换导出 data |
| `afterExport({ buffer, format }, ctx)` | `adapter.export()` 之后、`exported` 事件之前 | `Uint8Array` 或 `{ buffer }` |

不提供 bind/OOXML、`insertAnchor`、通用 `dispatch` 中间件。

```ts
const kernel = createKernel({
  adapter: new DocxAdapter(),
  validators: [crossField],
  formatters: { amountCn },
  plugins: [
    {
      afterHydrate({ data }) {
        return { data: { ...data, signDate: data.signDate ?? today } }
      },
      beforeExport(data) {
        return { ...data, stamp: 'CONFIDENTIAL' }
      },
    },
  ],
})
```

顺序：`afterHydrate` → `validators`；`formatters` → `beforeExport` → `bind` → `afterExport` → `subscribe('exported')`。

---

## Adapters

实现 `DocumentAdapter`，由 kernel 调用；接入方一般只 `new` 后交给 `createKernel`。

```ts
interface DocumentAdapter {
  kind: 'docx' | 'xlsx' | 'html' | 'custom'
  load(source: Source): Promise<void>
  discoverFields(): Promise<DiscoveredField[]>
  getPreview(): PreviewModel
  insertAnchor(field: Field): Promise<void>
  updateAnchor?(field: Field): Promise<void>
  removeAnchor(fieldId: string): Promise<void>
  bind(data: Record<string, unknown>): Promise<void>
  export(): Promise<Uint8Array>
}
```

| 类 / API | 导入 | 说明 |
|----------|------|------|
| `DocxAdapter` | `paperfill` | OOXML / JSZip；按标记 bind。`new DocxAdapter({ markers })` |
| `mountDocxPreview` | `paperfill` | 把 `.docx` 渲进传入的 `container` 并铺满宽度；默认 `inWrapper: false`、`ignoreWidth: true`；自定义起止符传 `markers`；`plugins`：`afterHtml` / `afterExpand` / `afterSlots` |
| `XlsxAdapter` | `paperfill` | ExcelJS；单元格文本中的标记；`getPreview` 含 `style`（背景/字体色等）。`new XlsxAdapter({ markers })` |
| `mountXlsxPreview` | `paperfill` | 把 `getPreview()` 的 sheets 渲成表格 DOM，业务只注入 `mountField`；`plugins`：`afterSheets` / `afterTable` |
| `HtmlAdapter` | `paperfill` | 接受 HTML 字符串或 UTF-8 字节；扫描、bind 并导出 HTML |
| `mountHtmlPreview` | `paperfill` | 渲染 HTML 并在 marker 位置挂字段；默认移除脚本、事件属性及危险 URL |

`custom` 不提供内置 adapter。业务实现 `DocumentAdapter`，`load` 中读取
`CustomSource.data`，并返回 `{ kind: 'custom', ... }` 预览模型。

```ts
import { mountDocxPreview } from 'paperfill'

const handle = mountDocxPreview(container, {
  buffer: source.buffer,
  fields: kernel.getFormSchema().fields,
  validation: kernel.validate(),
  mountField: myMounter,
  onChange: (path, value) => kernel.dispatch({ type: 'setValue', path, value }),
  markers: kernel.getMarkers(),
  // plugins: [{ afterHtml(root) { ... } }],
  // 默认铺满 container。要 A4 页框：
  // render: { inWrapper: true, ignoreWidth: false, ignoreHeight: false },
})
```

```ts
import { XlsxAdapter, mountXlsxPreview } from 'paperfill'
import 'paperfill/xlsx/style.css'

const preview = kernel.getPreview()
if (preview?.kind === 'xlsx') {
  const handle = mountXlsxPreview(container, {
    sheets: preview.sheets,
    fields: kernel.getFormSchema().fields,
    validation: kernel.validate(),
    mountField: myMounter, // 或 nativeFieldMounter
    onChange: (path, value) => kernel.dispatch({ type: 'setValue', path, value }),
    // plugins: [{ afterTable(root) { ... } }],
  })
  // handle.update({ sheets, fields, validation }); handle.destroy()
}
```

`XlsxPreviewCell.style`（可选）：`background` / `color`（`#rrggbb`）、`fontWeight`。`argb` 直出；仅有 `theme` 索引时用默认 Office 主题近似色。不覆盖渐变/条件格式。

`getPreview()` 默认裁掉末尾没有内容/样式的幽灵行和幽灵列（ExcelJS `rowCount` 常被空行撑大）；中间空行保留。导出文件不改。

Word 预览内置仍会先 `fitDocxToContainer`（把 `section.docx` 拉满宿主），再跑你的 `afterHtml`。可选再传 `docxFitHostPlugin`（幂等）。不要用 kernel 的 `afterHydrate` 当预览钩子——预览槽位挂完后是 `afterSlots`。

---

## 原生字段 UI（可选）

框架无关原生控件。

```ts
import { createField, mountField } from 'paperfill'
import 'paperfill/ui/style.css'

const handle = mountField(slotEl, {
  name: 'partyA',
  field: schemaField, // Partial<FormSchemaField>
  value: schemaField?.value,
  error: issueMessage,
  onChange: (value) => kernel.dispatch({ type: 'setValue', path: 'partyA', value }),
})

handle.update({ value, error, field })
handle.destroy()
```

| API | 说明 |
|-----|------|
| `createField(options)` | 创建控件，返回 `FieldHandle`（含 `el`） |
| `mountField(parent, options)` | 创建并挂到 `parent` |
| `nativeFieldMounter(container, ctx)` | 直接当作 `mountDocxPreview` / `mountXlsxPreview` 的 `mountField` |

`FieldHandle`：`el` / `update` / `destroy`。

不引入本包时，在槽位自挂组件即可（见 `examples/nuxt-demo` 的 `/custom`）。

---

## PDF（可选，浏览器）

对 **已 `export` 的文件** 再转 PDF（不是截填写页表单 DOM）。

```ts
import { exportFilledDocument, supportsCanvasDrawElement } from 'paperfill'

const filled = await kernel.dispatch({ type: 'export' })
if (filled.type !== 'exported') throw new Error('export failed')

// 注意：这里用 kind；kernel 导出结果字段名是 format（同义）
const pdf = await exportFilledDocument({
  kind: filled.format,
  buffer: filled.buffer,
  options: { mode: 'html2canvas' },
  // plugins: [{ afterPdfHtml(host) { ... } }],
})
```

### 输入

```ts
type FilledExportInput = {
  kind: 'docx' | 'xlsx'
  buffer: Uint8Array
  options?: ExportPdfOptions
  plugins?: PdfPlugin[] // afterPdfHtml(host, { kind })：填好的 HTML 之后、截图/打印之前
}

interface ExportPdfOptions {
  mode?: 'html2canvas' | 'canvas-draw-element' | 'print' // 默认 html2canvas
  scale?: number       // 默认 2
  format?: 'a4' | 'letter'
  marginMm?: number    // 默认 8
  pageSelector?: string
  background?: string
}
```

### API

| 函数 | 返回 | 说明 |
|------|------|------|
| `exportFilledDocument(input)` | `Uint8Array \| void` | 统一入口；`print` 时为 `void` |
| `exportFilledToPdf(input)` | `Uint8Array` | 仅截图像素模式 |
| `printFilledDocument(input)` | `Promise<void>` | 系统打印 → 另存为 PDF |
| `supportsCanvasDrawElement()` | `boolean` | 探测 Chromium html-in-canvas |
| `exportElementToPdf(el, options?)` | `Uint8Array` | 对已有 DOM 根节点截图导出 |

| `mode` | 行为 |
|--------|------|
| `html2canvas` | DOM → canvas → jsPDF 字节 |
| `canvas-draw-element` | 需开启 `chrome://flags/#canvas-draw-element` |
| `print` | `window.print()`，无文件字节 |

---

## hydrate 三件套

接入方只持久化这三样，恢复用 `hydrate`（不要把填写结果写回原 Word/Excel 当唯一真相）：

| 产物 | 来源 |
|------|------|
| `source` | 完整原始输入；docx/xlsx 为 buffer，HTML 可为字符串或 buffer，custom 为业务 data |
| `file` | 仅 docx/xlsx/HTML buffer 的兼容二进制别名 |
| `definition` | `kernel.getDefinition()` |
| `data` | `kernel.getData()` |

```ts
import { hydrateFromBundle, toPersistBundle } from 'paperfill'

const bundle = toPersistBundle(kernel)
if (bundle) await save(bundle)

await kernel.dispatch(hydrateFromBundle(bundle))
```

`snapshotKernel(kernel)` 给 UI 一层只读快照（schema / data / validation / preview）。

---

## Vue / React

薄封装，只订阅 kernel；文档布局仍走 `mountDocxPreview` / `mountXlsxPreview`。

```ts
import { useContractKit } from 'paperfill/vue'
// import { useContractKit } from 'paperfill/react'

const { schema, data, validation, preview } = useContractKit(kernel)
```

等价子包：`@paperfill/vue`、`@paperfill/react`（peer：Vue 3 / React ≥18）。

### SSR 边界

| 可在服务端 | 仅浏览器（`ClientOnly` / `useEffect` 后） |
|------------|------------------------------------------|
| `createKernel` / `dispatch` / `hydrate` / `validate` / `export` | `mountDocxPreview` / `mountXlsxPreview` |
| `toPersistBundle` / `snapshotKernel` | `@paperfill/ui` 的 `mountField` |
| 读 `definition` / `data` | `@paperfill/pdf` |

Nuxt 示例用 `<ClientOnly>` 包文档宿主。

---

## 典型调用顺序

```ts
// 1. 打开已发布模板
await kernel.dispatch({
  type: 'hydrate',
  source: { kind: 'docx', buffer },
  definition,
  data: savedData ?? {},
})

// 2. 渲染：getSource + getFormSchema；槽位 onChange → setValue
kernel.subscribe(() => {
  schema = kernel.getFormSchema()
  validation = kernel.validate()
})

await kernel.dispatch({ type: 'setValue', path: 'partyA', value: 'Acme' })

// 3. 导出 Office
const file = await kernel.dispatch({ type: 'export' })
// file.buffer / file.format

// 4. 可选 PDF
await exportFilledDocument({
  kind: file.format,
  buffer: file.buffer,
  options: { mode: 'html2canvas' },
})

// 5. 持久化
save({
  file: kernel.getSource()!.buffer,
  definition: kernel.getDefinition(),
  data: kernel.getData(),
})
```

---

## 示例（Nuxt · `examples/nuxt-demo`）

本地：`npm run example` → http://localhost:5210

| 路径 | 字段 UI | 校验 |
|------|---------|------|
| `/` | `@paperfill/ui` | 「校验」→ `kernel.validate()` |
| `/custom` | Element Plus 自绘 | 「校验」→ `el-form` rules（页面实现） |

API（Nitro）：

| 接口 | 说明 |
|------|------|
| `GET /api/contracts` | 合同列表 |
| `GET /api/contracts/:id` | `definition`（已合并 options）+ 草稿 `data` |
| `GET /api/contracts/:id/options` | 选项字典 |
| `GET /api/contracts/:id/file` | 模板二进制 |

---

## 明确不在本库范围

- 用户 / 权限 / 审批 / OSS
- 在线 Word/Excel 编辑器
- 服务端「正宗 Office → PDF」引擎
- `.doc` / `.xls` 老格式
