import {
  aggregateMarkerFields,
  createMarkerSyntax,
  normalizeMarkers,
  parseMarkers,
  rowsForExpand,
  stringifyFieldValue,
  tableNamesInText,
  type DiscoveredField,
  type DocumentAdapter,
  type Field,
  type HtmlPreview,
  type MarkerDelimiters,
  type Source,
} from '@paperfill/kernel'

export {
  mountHtmlPreview,
  type HtmlFieldHandle,
  type HtmlFieldMountContext,
  type HtmlFieldMounter,
  type HtmlPreviewHandle,
  type MountHtmlPreviewOptions,
} from './mount-preview'

export type HtmlAdapterOptions = {
  markers?: MarkerDelimiters
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function replaceFlatText(
  text: string,
  data: Record<string, unknown>,
  markers: MarkerDelimiters,
): string {
  return text.replace(createMarkerSyntax(markers).regex(), (_raw, name: string) => {
    if (!(name in data)) return ''
    return escapeHtml(stringifyFieldValue(data[name]))
  })
}

function replaceRowText(
  text: string,
  table: string,
  row: Record<string, unknown>,
  index: number,
  markers: MarkerDelimiters,
): string {
  const prefix = `${table}.`
  return text.replace(createMarkerSyntax(markers).regex(), (raw, name: string) => {
    if (!name.startsWith(prefix)) return raw
    const column = name.slice(prefix.length)
    if (column === '$index') return String(index + 1)
    return escapeHtml(stringifyFieldValue(row[column]))
  })
}

function replaceTextNodes(
  html: string,
  replace: (text: string) => string,
): string {
  return html
    .split(/(<[^>]+>)/g)
    .map((part) => (part.startsWith('<') ? part : replace(part)))
    .join('')
}

function expandTableRows(
  html: string,
  data: Record<string, unknown>,
  markers: MarkerDelimiters,
): string {
  let next = html
  for (const table of tableNamesInText(html, markers)) {
    const rows = rowsForExpand(data[table])
    next = next.replace(/<tr\b[^>]*>[\s\S]*?<\/tr\s*>/gi, (rowHtml) => {
      if (!rowHtml.includes(`${markers.start}${table}.`)) return rowHtml
      return rows
        .map((row, index) =>
          replaceTextNodes(rowHtml, (text) => replaceRowText(text, table, row, index, markers)),
        )
        .join('')
    })
  }
  return next
}

export function bindHtmlTemplate(
  html: string,
  data: Record<string, unknown>,
  markers?: MarkerDelimiters | null,
): string {
  const normalized = normalizeMarkers(markers)
  return replaceTextNodes(expandTableRows(html, data, normalized), (text) =>
    replaceFlatText(text, data, normalized),
  )
}

function visibleTemplateText(html: string): string {
  return html
    .replace(/<(script|style)\b[\s\S]*?<\/\1\s*>/gi, '')
    .replace(/<[^>]+>/g, ' ')
}

export class HtmlAdapter implements DocumentAdapter {
  readonly kind = 'html' as const
  private html = ''
  private loaded = false
  private bound: string | null = null
  private fields = new Map<string, Field>()
  private markers: MarkerDelimiters

  constructor(options?: HtmlAdapterOptions) {
    this.markers = normalizeMarkers(options?.markers)
  }

  setMarkers(markers: MarkerDelimiters): void {
    this.markers = normalizeMarkers(markers)
  }

  async load(source: Source): Promise<void> {
    if (source.kind !== 'html') throw new Error('HtmlAdapter only accepts html')
    this.html =
      source.html !== undefined ? source.html : new TextDecoder().decode(source.buffer)
    this.loaded = true
    this.bound = null
    this.fields.clear()
  }

  async discoverFields(): Promise<DiscoveredField[]> {
    return aggregateMarkerFields(parseMarkers(visibleTemplateText(this.html), this.markers)).map(
      (field) => ({
        name: field.name,
        type: field.type,
        label: field.name,
        columns: field.columns,
        anchor: { kind: 'marker' as const, name: field.name },
      }),
    )
  }

  getPreview(): HtmlPreview {
    return { kind: 'html', html: this.html }
  }

  async insertAnchor(field: Field): Promise<void> {
    this.fields.set(field.id, field)
    const syntax = createMarkerSyntax(this.markers)
    if (syntax.namedRegex(field.name).test(this.html)) return
    const marker = `<p>${escapeHtml(syntax.wrap(field.name, field.type))}</p>`
    this.html = /<\/body\s*>/i.test(this.html)
      ? this.html.replace(/<\/body\s*>/i, `${marker}</body>`)
      : `${this.html}${marker}`
    this.bound = null
  }

  async updateAnchor(field: Field): Promise<void> {
    const previous = this.fields.get(field.id)
    this.fields.set(field.id, field)
    const syntax = createMarkerSyntax(this.markers)
    const next = syntax.wrap(field.name, field.type)
    const updated = this.html.replace(syntax.namedRegex(previous?.name ?? field.name), next)
    if (updated !== this.html) {
      this.html = updated
      this.bound = null
      return
    }
    await this.insertAnchor(field)
  }

  async removeAnchor(fieldId: string): Promise<void> {
    const field = this.fields.get(fieldId)
    this.fields.delete(fieldId)
    if (!field) return
    this.html = this.html.replace(createMarkerSyntax(this.markers).namedRegex(field.name), '')
    this.bound = null
  }

  async bind(data: Record<string, unknown>): Promise<void> {
    this.bound = bindHtmlTemplate(this.html, data, this.markers)
  }

  async export(): Promise<Uint8Array> {
    if (!this.loaded) throw new Error('no html loaded')
    return new TextEncoder().encode(this.bound ?? this.html)
  }
}
