import type { FormSchema, Kernel, PreviewModel, Source, TemplateDefinition, ValidationResult } from './types'

/** Persist the three artifacts the page owns. */
export type PersistBundle = {
  /** Complete source, including HTML strings and custom adapter data. */
  source: Source
  /** Backward-compatible binary alias for docx/xlsx bundles. */
  file?: Uint8Array
  definition: TemplateDefinition
  data: Record<string, unknown>
}

type LegacyPersistBundle = {
  file: Uint8Array
  source?: never
  definition: TemplateDefinition
  data: Record<string, unknown>
}

export type KernelSnapshot = {
  schema: FormSchema
  data: Record<string, unknown>
  validation: ValidationResult
  preview: PreviewModel | null
  definition: TemplateDefinition | null
  source: Source | null
}

export function snapshotKernel(kernel: Kernel): KernelSnapshot {
  return {
    schema: kernel.getFormSchema(),
    data: kernel.getData(),
    validation: kernel.validate(),
    preview: kernel.getPreview(),
    definition: kernel.getDefinition(),
    source: kernel.getSource(),
  }
}

export function toPersistBundle(kernel: Kernel): PersistBundle | null {
  const source = kernel.getSource()
  const definition = kernel.getDefinition()
  if (!source || !definition) return null
  const bundle: PersistBundle = {
    source,
    definition,
    data: kernel.getData(),
  }
  if ('buffer' in source && source.buffer instanceof Uint8Array) bundle.file = source.buffer
  return bundle
}

export function hydrateFromBundle(bundle: PersistBundle | LegacyPersistBundle) {
  const source =
    bundle.source ??
    ({
      kind: bundle.definition.source.kind,
      buffer: bundle.file,
    } as Source)
  return {
    type: 'hydrate' as const,
    source,
    definition: bundle.definition,
    data: bundle.data,
  }
}
