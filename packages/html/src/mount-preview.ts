import {
  createMarkerSyntax,
  parseTableColumnRef,
  rowsForExpand,
  splitByMarkers,
  type FormSchemaField,
  type MarkerDelimiters,
  type ValidationResult,
} from '@paperfill/kernel'

export type HtmlFieldMountContext = {
  name: string
  field?: FormSchemaField | Partial<FormSchemaField>
  value?: unknown
  error?: string
  onChange: (value: unknown) => void
}

export type HtmlFieldHandle = {
  update: (patch: {
    field?: FormSchemaField | Partial<FormSchemaField>
    value?: unknown
    error?: string
  }) => void
  destroy: () => void
}

export type HtmlFieldMounter = (
  container: HTMLElement,
  ctx: HtmlFieldMountContext,
) => HtmlFieldHandle

export type MountHtmlPreviewOptions = {
  html: string | Uint8Array
  fields: FormSchemaField[]
  validation?: ValidationResult
  mountField: HtmlFieldMounter
  onChange: (path: string, value: unknown) => void
  markers?: MarkerDelimiters
  /** Disable only for already-sanitized, trusted templates. Default: true. */
  sanitize?: boolean
}

export type HtmlPreviewHandle = {
  update: (patch: Partial<MountHtmlPreviewOptions>) => Promise<void>
  destroy: () => void
}

function fieldAt(
  name: string,
  fields: FormSchemaField[],
  validation?: ValidationResult,
): { field: Partial<FormSchemaField>; value: unknown; error?: string } {
  const error = validation?.issues.find((issue) => issue.path === name)?.message
  const direct = fields.find((field) => field.name === name)
  if (direct) return { field: direct, value: direct.value, error }
  const ref = /^([^.]+)\.(\d+)\.([^.]+)$/.exec(name)
  if (!ref) return { field: { name, type: 'text' }, value: undefined, error }
  const [, tableName, rowText, columnName] = ref
  const table = fields.find((field) => field.name === tableName)
  const column = table?.columns?.find((item) => item.name === columnName)
  const rows = Array.isArray(table?.value) ? (table.value as Record<string, unknown>[]) : []
  const value = rows[Number(rowText)]?.[columnName]
  return {
    field: {
      name,
      type: column?.type ?? 'text',
      label: column?.label ?? columnName,
      required: Boolean(column?.required),
      options: column?.options,
      rules: column?.rules,
      outputFormat: column?.outputFormat,
      value,
    },
    value,
    error,
  }
}

function sanitize(root: HTMLElement) {
  for (const blocked of root.querySelectorAll('script, iframe, object, embed, base')) blocked.remove()
  for (const element of root.querySelectorAll('*')) {
    for (const attribute of [...element.attributes]) {
      const name = attribute.name.toLowerCase()
      const value = attribute.value.trim().toLowerCase()
      if (name.startsWith('on') || ((name === 'href' || name === 'src') && value.startsWith('javascript:'))) {
        element.removeAttribute(attribute.name)
      }
    }
  }
}

function rewriteRow(row: HTMLElement, table: string, index: number, markers?: MarkerDelimiters) {
  const syntax = createMarkerSyntax(markers)
  const walker = row.ownerDocument.createTreeWalker(row, 4)
  let node = walker.nextNode()
  while (node) {
    const text = node.textContent ?? ''
    node.textContent = text.replace(syntax.regex(), (raw, name: string) => {
      const ref = parseTableColumnRef(name)
      if (!ref || ref.table !== table) return raw
      if (ref.column === '$index') return String(index + 1)
      return syntax.wrap(`${table}.${index}.${ref.column}`)
    })
    node = walker.nextNode()
  }
}

function expandRows(root: HTMLElement, fields: FormSchemaField[], markers?: MarkerDelimiters) {
  const syntax = createMarkerSyntax(markers)
  for (const field of fields) {
    if (field.type !== 'table') continue
    const templates = [...root.querySelectorAll('tr')].filter((row) =>
      (row.textContent ?? '').includes(`${syntax.start}${field.name}.`),
    )
    for (const template of templates) {
      const parent = template.parentElement
      if (!parent) continue
      const fragment = root.ownerDocument.createDocumentFragment()
      const rows = rowsForExpand(field.value)
      rows.forEach((_row, index) => {
        const clone = template.cloneNode(true) as HTMLElement
        rewriteRow(clone, field.name, index, markers)
        fragment.appendChild(clone)
      })
      parent.insertBefore(fragment, template)
      template.remove()
    }
  }
}

export function mountHtmlPreview(
  container: HTMLElement,
  initial: MountHtmlPreviewOptions,
): HtmlPreviewHandle {
  let options = initial
  let handles: HtmlFieldHandle[] = []

  function destroyFields() {
    handles.forEach((handle) => handle.destroy())
    handles = []
  }

  function render() {
    destroyFields()
    container.replaceChildren()
    const html = typeof options.html === 'string' ? options.html : new TextDecoder().decode(options.html)
    const parser = container.ownerDocument.defaultView?.DOMParser
    if (parser) {
      const parsed = new parser().parseFromString(html, 'text/html')
      for (const style of [...parsed.head.querySelectorAll('style')]) {
        container.appendChild(container.ownerDocument.importNode(style, true))
      }
      for (const child of [...parsed.body.childNodes]) {
        container.appendChild(container.ownerDocument.importNode(child, true))
      }
    } else {
      container.innerHTML = html
    }
    if (options.sanitize !== false) sanitize(container)
    expandRows(container, options.fields, options.markers)

    const walker = container.ownerDocument.createTreeWalker(container, 4)
    const textNodes: Text[] = []
    let node = walker.nextNode()
    while (node) {
      if (createMarkerSyntax(options.markers).contains(node.textContent ?? '')) textNodes.push(node as Text)
      node = walker.nextNode()
    }
    for (const textNode of textNodes) {
      const segments = splitByMarkers(textNode.textContent ?? '', options.markers)
      const fragment = container.ownerDocument.createDocumentFragment()
      for (const segment of segments) {
        if (segment.kind === 'text') {
          fragment.appendChild(container.ownerDocument.createTextNode(segment.text))
          continue
        }
        const holder = container.ownerDocument.createElement('span')
        holder.className = 'ck-field-slot'
        holder.dataset.field = segment.name
        const resolved = fieldAt(segment.name, options.fields, options.validation)
        handles.push(
          options.mountField(holder, {
            name: segment.name,
            field: resolved.field,
            value: resolved.value,
            error: resolved.error,
            onChange: (value) => options.onChange(segment.name, value),
          }),
        )
        fragment.appendChild(holder)
      }
      textNode.replaceWith(fragment)
    }
  }

  render()
  return {
    async update(patch) {
      options = { ...options, ...patch }
      render()
    },
    destroy() {
      destroyFields()
      container.replaceChildren()
    },
  }
}
