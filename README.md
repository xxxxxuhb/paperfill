# paperfill

**中文** | [English](./README.en.md)

**在 Word / Excel 原文排版上填写合同，而不是另开一张表再生成文档。**

`paperfill` 是一套 **headless** 合同模板运行时：内核只管字段定义、填写数据与导出；页面自己决定预览和输入框长什么样。库**不做**账号、OSS、审批、合同系统。

## 使用场景

| 场景 | 谁做什么 |
|------|----------|
| 采购 / 销售合同在**原文**填写 | 法务出 Word/Excel 排版，正文写 `{{partyA}}`、`{{items.qty}}`；业务员在浏览器里对着原文填，而不是另开一张表 |
| 模板发布 | 运营上传 `.docx` / `.xlsx`；后端（或发布脚本）`load` 扫标记，补上 type / label / required / options / outputFormat，存成 **definition** |
| 草稿与续填 | 打开已发布合同：拉回原文件 + definition + 上次 `data`，`hydrate` 后接着填 |
| 导出给对方 / 归档 | 浏览器 `export` 出已填写 Word/Excel；需要 PDF 时再对**已导出文件**做浏览器转 PDF |
| 批量套打（无预览） | 服务端只跑 kernel：`hydrate` + `setData` + `export`，不挂 DOM |

不适合：在线 Word 编辑器、把填写结果只写回文档当唯一真相、服务端「正宗 Office → PDF」。

## 前后端怎么配合

**真相是三样东西**，都由你们的后端存；文档文件只是模板和导出载体。

| 产物 | 谁写 | 说明 |
|------|------|------|
| 原始 `.docx` / `.xlsx` | 后端对象存储 | 发布后不要改；`export` 只在副本上 bind |
| `definition` JSON | 发布时生成，之后可改字段元数据 | 类型、label、必填、options、锚点 |
| `data` JSON | 填写页随改随存 | 用户填的值；明细表是数组 |

```text
法务/运营                 你们的后端                      填写页（浏览器）
   │                         │                                │
   │  上传 采购合同.docx      │                                │
   │  （正文含 {{marker}}）   │                                │
   ├────────────────────────►│  load → 扫标记                  │
   │                         │  合并业务 meta                  │
   │                         │  存 file + definition           │
   │                         │                                │
   │                         │◄──── GET /contracts/:id ───────┤
   │                         │     file + definition + data   │
   │                         │                                │
   │                         │                    hydrate + 原文预览
   │                         │                    setValue → 草稿 PUT data
   │                         │                    export → 本地下载 Word/Excel/PDF
```

### 后端（存储 + 发布，不画 UI）

- 存三件套；提供「列表 / 详情 / 原文件下载 / 保存草稿」。
- **发布模板**（Node 即可，无需浏览器）：

```ts
const kernel = createKernel({ adapter: new DocxAdapter() })
await kernel.dispatch({ type: 'load', source: { kind: 'docx', buffer } })
// 再 updateField 补 label / options，或直接改 getDefinition()
await db.saveTemplate({
  file: buffer,
  definition: kernel.getDefinition(),
})
```

- **批量出件**（同样在 Node）：`hydrate` + `setData` + `export`，把 `exported.buffer` 写进 OSS。不要用导出文件覆盖原模板。
- 权限、审批流、水印策略属于业务，不在本库。

示例里的模拟后端：`GET /api/contracts`、`/:id`（definition + data）、`/:id/file`（二进制）。

### 前端（预览 + 填写 + 导出）

```ts
import { createKernel, DocxAdapter, XlsxAdapter, mountDocxPreview, nativeFieldMounter } from 'paperfill'

const kernel = createKernel({
  adapter: detail.kind === 'docx' ? new DocxAdapter() : new XlsxAdapter(),
})

await kernel.dispatch({
  type: 'hydrate',
  source: { kind: detail.kind, buffer: fileBuffer },
  definition: detail.definition,
  data: detail.data ?? {},
})

// 仅浏览器：原文排版 + 槽位挂控件
mountDocxPreview(el, {
  buffer: fileBuffer,
  fields: kernel.getFormSchema().fields,
  mountField: nativeFieldMounter, // 或自绘 Element Plus
  onChange: (path, value) => kernel.dispatch({ type: 'setValue', path, value }),
})

// 存草稿：只提交 data（模板 file / definition 一般不变）
await api.saveDraft(id, kernel.getData())

const office = await kernel.dispatch({ type: 'export' })
// PDF：对 office.buffer 再 exportFilledDocument（勿在 SSR 里调）
```

`mountDocxPreview` / `mountXlsxPreview` / `nativeFieldMounter` / `exportFilledDocument` 必须在 **ClientOnly / useEffect 之后**；`createKernel` / `hydrate` / `export` 前后端都能跑。

更细的时序见 [docs/architecture.md](./docs/architecture.md)，接口见 [docs/api.md](./docs/api.md)。

## 示例（Nuxt）

```bash
npm install
npm run templates
npm run example          # http://localhost:5210
```

| 路径 | 说明 |
|------|------|
| `/` | 原生 UI（`nativeFieldMounter`），校验 = `kernel.validate()` |
| `/custom` | Element Plus 自绘字段，校验 = 页面 `el-form` rules |

模拟 API：`/api/contracts`、`/:id`、`/:id/options`、`/:id/file`（Nitro）。

模板：`examples/templates` 下 Word + Excel，覆盖全部字段类型。

## 安装

```bash
npm i paperfill
```

只装这一个包，**API 都从 `paperfill` 取**（Word / Excel / 原生控件 / PDF 都在同一入口）：

```ts
import {
  createKernel,
  DocxAdapter,
  mountDocxPreview,
  XlsxAdapter,
  mountXlsxPreview,
  HtmlAdapter,
  mountHtmlPreview,
  nativeFieldMounter,
  exportFilledDocument,
} from 'paperfill'
import 'paperfill/ui/style.css'
import 'paperfill/xlsx/style.css'
```

HTML 模板可直接传字符串，也可传 UTF-8 文件字节：

```ts
const kernel = createKernel({ adapter: new HtmlAdapter() })
await kernel.dispatch({
  type: 'load',
  source: { kind: 'html', html: '<p>甲方：{{partyA}}</p>' },
  // 文件：source: { kind: 'html', buffer: new Uint8Array(await file.arrayBuffer()) }
})
```

`kind: 'custom'` 不带内置实现：业务实现 `DocumentAdapter`，通过
`source: { kind: 'custom', data, hash? }` 接收任意模板数据。不能稳定 JSON 序列化时必须提供 `hash`。

样式仍走子路径。Vue / React 订阅因框架 peer 不同，用 `paperfill/vue` 或 `paperfill/react`。

文档：[docs/architecture.md](./docs/architecture.md) · [docs/api.md](./docs/api.md)

## 限制与环境

| 项 | 要求 / 说明 |
|----|-------------|
| **Node** | `>= 20`（构建、示例、脚本） |
| **浏览器** | 现代 evergreen：Chrome / Edge **≥ 94**，Firefox **≥ 93**，Safari **≥ 15.4**（编译目标 ES2022；需 `Uint8Array`、基础 DOM） |
| **运行时** | `kernel` / `docx` / `xlsx` / `html` adapter 可在 Node 与浏览器使用；`ui`、各 `mount*Preview`、`pdf` 依赖 **浏览器 DOM**，勿在纯 SSR 路径直接调用 |
| **PDF** | `exportFilledDocument` **仅浏览器**；默认 `html2canvas`，不是服务端 Office 矢量转 PDF；`canvas-draw-element` 需 Chromium 实验 flag |
| **Word 预览** | `mountDocxPreview` 版式靠 `docx-preview`，与桌面 Word 不完全一致；自定义标记起止符时传 `markers: kernel.getMarkers()` |
| **Excel 预览** | `getPreview` / `mountXlsxPreview` 支持单元格填充色、字体色（`argb`；主题色为近似）；默认去掉末尾幽灵行/列；不覆盖渐变、条件格式、图表等 |
| **图片字段** | data URL；`bind`/`export` 嵌入 Word drawing / Excel 单元格图（固定约 120px，无裁剪/锚点微调） |
| **非目标** | 完整 Office 编辑器、服务端「正宗」排版引擎；`ui` 仍是框架无关 DOM，Vue/React 只提供订阅薄封装 |

## License

MIT
