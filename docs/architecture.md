# paperfill 架构

面向接入方的公共 API 清单见 [api.md](./api.md)。

## 一句话

**headless kernel + document adapter + 页面自绘预览/字段 UI（可选原生 ui / pdf）。**

库不做业务系统。页面只通过 `dispatch` / query / `subscribe` 驱动；真相是三样持久化产物，文档文件只是渲染与导出载体。

| 产物 | 谁存 | 说明 |
|------|------|------|
| 原始 `.docx` / `.xlsx` | 接入方 | 不可变模板 |
| `definition`（字段 JSON） | 接入方 | 类型、label、required、options、锚点 |
| `data`（填写 JSON） | 接入方 | 用户填的值 |

预览按原文排版嵌输入框；导出 = `原始文件 ⊕ definition ⊕ data`（在**副本**上 bind，不改原 buffer）。

---

## 包结构

```text
paperfill/
  packages/
    paperfill   # 伞包：默认 re-export kernel（子路径见各自 package）
    kernel         # 状态机：Command / State / Schema / View / Preview 缓存
    docx           # Word adapter：扫 {{marker}}、bind、export、getPreview、mountDocxPreview
    xlsx           # Excel adapter：单元格标记、合并预览、bind、export、mountXlsxPreview
    html           # HTML adapter：字符串/UTF-8 文件、bind、export、mountHtmlPreview
    ui             # 可选：框架无关原生 createField / mountField / nativeFieldMounter
    vue / react    # 可选：useContractKit 订阅快照（peer 框架）
    pdf            # 可选（浏览器）：已填写 docx/xlsx → PDF
  examples/
    nuxt-demo      # Nuxt：原生 / + 自定义 /custom；Nitro 模拟 API（:5210）
    mock           # 共用模拟后端数据
    templates      # 采购合同.docx / .xlsx + definition
```

| 包 | 做 | 不做 |
|----|----|------|
| `kernel` | 状态、命令、校验、schema/view、缓存 preview | DOM、框架、解析 OOXML |
| `docx` | load / discoverFields / getPreview(blocks) / bind / export / **`mountDocxPreview` 文档布局** | 存业务 data、字段控件 |
| `xlsx` | load / discoverFields / getPreview(含单元格样式) / bind / export / **`mountXlsxPreview` 表格布局** | 存业务 data、字段控件 |
| `html` | HTML 字符串或 UTF-8 文件的 discover / bind / export / **`mountHtmlPreview`** | 执行业务脚本、解释 custom 数据 |
| `ui` | 原生控件（含 select / multiselect / display / image 选文件）+ `nativeFieldMounter` | 绑死某一框架、文档排版 |
| `vue` / `react` | `useContractKit` 订阅 schema/data/validation | 自绘字段、改 OOXML |
| `pdf` | 浏览器里把**已导出文件**渲成 PDF | 服务端 Office 矢量转换、改 kernel 状态 |
| `custom` | 由接入方实现 `DocumentAdapter`，kernel 仍提供状态、命令、校验和插件 | 内置布局或导出格式 |
| 页面 | 上传、OSS、权限、注入 FieldMounter | 直接改 zip/OOXML；自绘 Excel 表格 |

---

## 分层总览

```mermaid
flowchart TB
  subgraph Page["接入方页面 / examples"]
    Shell["壳：打开 / hydrate / 导出"]
    DocxHost["Word：mountDocxPreview + FieldMounter"]
    FieldUI["字段 UI：@paperfill/ui<br/>或自绘 Element/自有组件"]
  end

  subgraph Core["运行时"]
    Kernel["@paperfill/kernel<br/>dispatch / get* / subscribe"]
    Docx["DocxAdapter + mountDocxPreview"]
    Xlsx["XlsxAdapter + mountXlsxPreview"]
  end

  subgraph Optional["可选"]
    Pdf["@paperfill/pdf<br/>exportFilledDocument"]
  end

  Shell -->|Command| Kernel
  Docx -->|"文档布局 + 槽位"| FieldUI
  Xlsx -->|"表格布局 + 槽位"| FieldUI
  FieldUI -->|setValue| Kernel
  Kernel --> Docx
  Kernel --> Xlsx
  Shell -->|export 得到 buffer| Pdf
  Docx --> OOXML[".docx / JSZip"]
  Xlsx --> XLSX[".xlsx / ExcelJS"]
```

**字段 UI 与布局解耦：** Word 用 `mountDocxPreview` 渲原文件并挂槽位；Excel 用 `mountXlsxPreview` 渲表格（含背景/字体色）。接入方只向槽位 `mount` 控件。以上 API 均从 `paperfill` 导入。

**字段 UI 与 kernel 解耦：** kernel 只产出 `FormSchema`；谁往槽位里 `mount` 控件由接入方决定。`ui` 是默认实现，不是必选。

**PDF 与填写解耦：** 先 `dispatch({ type: 'export' })` 得到填好的 Office 文件，再 `exportFilledDocument({ kind: filled.format, buffer })`。注意 API 用 `kind`，kernel 导出结果字段名是 `format`（二者同义）。

---

## 核心对象

```mermaid
classDiagram
  class Kernel {
    +dispatch(Command)
    +getState()
    +getDefinition()
    +getData()
    +getFormSchema()
    +getView()
    +getPreview()
    +getSource()
    +validate()
    +subscribe()
  }

  class KernelState {
    definition: TemplateDefinition
    data: Record
    source: Source
    validation: ValidationResult
  }

  class TemplateDefinition {
    version: 1
    source: kind + hash
    fields: Field[]
  }

  class Field {
    id
    name
    type
    label?
    required?
    options?
    anchor
  }

  class DocumentAdapter {
    <<interface>>
    load()
    discoverFields()
    getPreview()
    insertAnchor()
    bind()
    export()
  }

  Kernel --> KernelState
  Kernel --> DocumentAdapter
  KernelState --> TemplateDefinition
  TemplateDefinition --> Field
  DocxAdapter ..|> DocumentAdapter
  XlsxAdapter ..|> DocumentAdapter
```

文档只写锚点：`{{partyA}}`、`{{items.qty}}`。`type` / `options` / `outputFormat` 写在随模板发布的 `*.definition.json`，填写页用 `hydrate` 恢复。  
`{{name:type}}` 仅作 `load` 时的可选提示，不以文档为准。

---

## 命令与只读 API

| Command | 作用 |
|---------|------|
| `load` | 加载文件，扫标记生成 fields |
| `hydrate` | 用已存 definition + data 恢复（已发布模板主路径） |
| `insertField` / `updateField` / `removeField` | 改 definition |
| `setValue` / `setData` / `resetData` | 改填写数据 |
| `export` | `formatData`（`outputFormat`）+ bind + 导出新文件 → `{ type:'exported', buffer, format }` |

只读：`getFormSchema`、`getView`、`getPreview`、`getSource`、`validate`、`can`。  
事件：`subscribe` → `state-changed` / `data-changed` / `validated` / …

---

## 时序：打开已发布模板并填写

```mermaid
sequenceDiagram
  actor User as 用户
  participant Page as 页面
  participant K as Kernel
  participant A as DocxAdapter
  participant Prev as docx-preview

  User->>Page: 打开已发布合同
  Page->>Page: 拉取 .docx + .definition.json
  Page->>K: createKernel({ adapter })
  Page->>K: dispatch({ type: 'hydrate', source, definition, data })

  K->>A: load(source)
  A-->>K: 解析 zip，保留原 buffer
  loop definition.fields
    K->>A: insertAnchor(field)
  end
  K->>A: getPreview()
  K-->>Page: data-changed / state-changed

  Page->>K: getSource() / getFormSchema()
  Page->>Prev: renderAsync(原 buffer)
  Prev-->>Page: Word 页面 DOM
  Page->>Page: 扫 {{marker}} → FieldMounter 挂控件
  User->>Page: 在原文位置输入
  Page->>K: dispatch({ type: 'setValue', path, value })
  K->>K: 更新 data + validate
  K-->>Page: data-changed / validated
  Page->>Page: 刷新槽位显示
```

仅有文件、尚无 definition 时：先 `load`（discoverFields），再上传/合并 definition 后 `hydrate`。

---

## 时序：导出 Office / PDF

```mermaid
sequenceDiagram
  actor User as 用户
  participant Page as 页面
  participant K as Kernel
  participant A as Adapter
  participant Pdf as @paperfill/pdf

  User->>Page: 导出文件
  Page->>K: dispatch({ type: 'export' })
  K->>A: bind(formatData)（副本上替换 marker）
  K->>A: export()
  A-->>K: Uint8Array
  K-->>Page: { type: 'exported', buffer, format }
  Page->>User: 下载 .docx / .xlsx

  User->>Page: 导出 PDF
  Page->>K: dispatch({ type: 'export' })
  K-->>Page: filled.buffer + filled.format
  Page->>Pdf: exportFilledDocument({ kind: format, buffer, options.mode })
  Note over Pdf: html2canvas → 字节<br/>canvas-draw-element → 字节（需 Chromium flag）<br/>print → 系统打印另存为 PDF
  Pdf-->>Page: Uint8Array | void
```

PDF 在浏览器侧再次渲染**已 bind 的文件**（docx-preview / 简易表格），不是截填写页上的表单 DOM。

| `mode` | 行为 |
|--------|------|
| `html2canvas`（默认） | DOM 截图 + jsPDF → `Uint8Array` |
| `canvas-draw-element` | WICG html-in-canvas；需开启 flag |
| `print` | `window.print()` →「另存为 PDF」，无字节返回 |

---

## 时序：持久化与恢复

```mermaid
sequenceDiagram
  participant Page as 接入方
  participant Store as 接入方存储
  participant K as Kernel

  Note over Page,K: 保存
  Page->>K: getDefinition() / getData()
  Page->>Store: 存 { file, definition, data }

  Note over Page,K: 恢复
  Page->>Store: 读回三样
  Page->>K: dispatch({ type: 'hydrate', source, definition, data })
  Page->>Page: 再按原文件排版 + 挂字段
```

---

## 示例怎么接

| 目录 / 路径 | 说明 |
|-------------|------|
| `examples/nuxt-demo` `/` | 只依赖 `paperfill`；原生 UI；「校验」→ `kernel.validate()` |
| `examples/nuxt-demo` `/custom` | Element Plus；「校验」→ `el-form` rules |
| `examples/templates` | Word / Excel 全类型模板 |
| `examples/mock` | Nitro API 共用数据 |

本地：`npm run example` → http://localhost:5210

---

## 明确不做

- 在线 Word/Excel 编辑器  
- 用户 / 租户 / 审批 / OSS  
- 把填写结果只写进文档当唯一真相  
- 老格式 `.doc` / `.xls`  
- 服务端「正宗 Office → PDF」引擎（`pdf` 包是浏览器方案）

## 待办

库自身缺口（不含 ExcelJS / docx-preview / html2canvas 等第三方上限）：

- [x] **Word 布局 API**：`mountDocxPreview`（docx-preview 版式 + 槽位水合）；接入方只注入 FieldMounter。
- [x] **image bind / 导出嵌图**：data URL → Word `w:drawing` / Excel 单元格图。
- [x] **可扩展校验**：`Field.rules`（min/max/长度/正则/日期）+ `createKernel({ validators })` 跨字段。
- [x] **导出格式**：`Field.outputFormat` + `createKernel({ formatters })`；`data` 保持规范值，`export` / `getExportData()` 套格式。文档标记默认 `{{name}}`，可配 `markers` 起止符。
- [x] **插件钩子**：可选 `plugins`（kernel / 预览 / PDF），不替代 `validators` / `formatters` / `subscribe` / `mountField`。
- [x] **锚点写回文档**：`insertField` 在 Word 文末 / Excel 新行写入 `{{name}}`；`updateField` 改标记；`removeField` 删标记。
- [x] **预览与导出一致**：填写页（原文件+槽位）与 `export`（文本替换+扩行+嵌图）对齐；未填扁平标记导出清空；空表两边各留一行占位。版式仍受 docx-preview / ExcelJS 第三方上限约束。
- [x] **表格行 API**：`insertRow` / `removeRow`；空表预览/导出保留一行空白占位（UI 仍可由业务做）。
- [x] **接入面**：`@paperfill/vue` / `react` 的 `useContractKit`；`toPersistBundle` / `hydrateFromBundle` / `snapshotKernel`；SSR 边界见 api.md；umbrella 子路径 + [CHANGELOG](../CHANGELOG.md)。

已完成：

- [x] **循环明细表**：`{{items.col}}` 行模板 + `data.items[]`；docx/xlsx bind 扩行；预览按 data 行数渲染（空表一行占位）；行数由 `insertRow` / `removeRow` / `setValue` 控制。
- [x] **Nuxt 示例**：`npm run example`（`/` 原生 UI + `/custom` Element Plus）。
- [x] **Excel 布局进 xlsx**：`mountXlsxPreview` + `getPreview` 单元格颜色；预览裁末尾幽灵行/列。
- [x] **Word 布局进 docx**：`mountDocxPreview`（demo `DocxLayout` 仅为薄宿主）。
