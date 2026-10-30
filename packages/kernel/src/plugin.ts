import type {
  DiscoveredField,
  ExportFormat,
  KernelPlugin,
  Source,
  TemplateDefinition,
} from './types'

export function applyAfterDiscover(
  plugins: KernelPlugin[],
  fields: DiscoveredField[],
  source: Source,
): DiscoveredField[] {
  let next = fields
  for (const plugin of plugins) {
    const patched = plugin.afterDiscover?.(next, { source })
    if (patched) next = patched
  }
  return next
}

export function applyAfterHydrate(
  plugins: KernelPlugin[],
  ctx: { definition: TemplateDefinition; data: Record<string, unknown>; source: Source },
): Record<string, unknown> {
  let data = ctx.data
  for (const plugin of plugins) {
    const patched = plugin.afterHydrate?.({ definition: ctx.definition, data, source: ctx.source })
    if (patched?.data) data = patched.data
  }
  return data
}

export function applyBeforeExport(
  plugins: KernelPlugin[],
  data: Record<string, unknown>,
  ctx: { definition: TemplateDefinition; source: Source },
): Record<string, unknown> {
  let next = data
  for (const plugin of plugins) {
    const patched = plugin.beforeExport?.(next, ctx)
    if (patched) next = patched
  }
  return next
}

export function applyAfterExport(
  plugins: KernelPlugin[],
  result: { buffer: Uint8Array; format: ExportFormat },
  ctx: { definition: TemplateDefinition; source: Source },
): Uint8Array {
  let buffer = result.buffer
  for (const plugin of plugins) {
    const patched = plugin.afterExport?.({ buffer, format: result.format }, ctx)
    if (!patched) continue
    buffer = patched instanceof Uint8Array ? patched : patched.buffer
  }
  return buffer
}
