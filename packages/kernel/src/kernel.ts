import { formatData } from './format'
import { hashBytes, hashText, hashValue } from './hash'
import { createId } from './id'
import { isDefaultMarkers, normalizeMarkers, DEFAULT_MARKERS } from './markers'
import { cloneData, insertTableRow, removeTableRow, setDataPath } from './path'
import {
  applyAfterDiscover,
  applyAfterExport,
  applyAfterHydrate,
  applyBeforeExport,
} from './plugin'
import { buildFormSchema, buildView, emptyValidation, validateState } from './schema'
import type {
  Command,
  CreateKernelOptions,
  DispatchResult,
  DocumentAdapter,
  Field,
  Kernel,
  KernelEvent,
  KernelPlugin,
  KernelState,
  MarkerDelimiters,
  PreviewModel,
  Source,
  TemplateDefinition,
  ViewportPort,
} from './types'

export function createKernel(options: CreateKernelOptions): Kernel {
  const adapter: DocumentAdapter = options.adapter
  const validators = options.validators ?? []
  const formatters = options.formatters
  const plugins: KernelPlugin[] = options.plugins ?? []
  const defaultMarkers = normalizeMarkers(options.markers)
  let markers: MarkerDelimiters = defaultMarkers
  let viewport: ViewportPort | null = null
  let preview: PreviewModel | null = null
  const listeners = new Set<(event: KernelEvent) => void>()

  let state: KernelState = {
    definition: null,
    data: {},
    source: null,
    validation: emptyValidation(),
  }

  function emit(event: KernelEvent) {
    for (const listener of listeners) listener(event)
    if (event.type !== 'state-changed') {
      for (const listener of listeners) listener({ type: 'state-changed' })
    }
  }

  function snapshot(): KernelState {
    return {
      definition: state.definition
        ? { ...state.definition, fields: state.definition.fields.map((f) => ({ ...f })) }
        : null,
      data: cloneData(state.data),
      source: state.source,
      validation: { ...state.validation, issues: [...state.validation.issues] },
    }
  }

  function requireDefinition(): TemplateDefinition {
    if (!state.definition || !state.source) {
      throw new Error('kernel has no loaded template')
    }
    return state.definition
  }

  function refreshValidation() {
    state.validation = validateState(state, validators)
    emit({ type: 'validated', result: state.validation })
  }

  function applyMarkers(next: MarkerDelimiters) {
    markers = normalizeMarkers(next)
    adapter.setMarkers?.(markers)
  }

  async function sourceHash(source: Source): Promise<string> {
    if (source.hash) return source.hash
    switch (source.kind) {
      case 'docx':
      case 'xlsx':
        return hashBytes(source.buffer)
      case 'html':
        return source.html !== undefined ? hashText(source.html) : hashBytes(source.buffer)
      case 'custom':
        return hashValue(source.data)
    }
  }

  function persistMarkers(): MarkerDelimiters | undefined {
    return isDefaultMarkers(markers) ? undefined : markers
  }

  function exportPayload() {
    if (!state.definition || !state.source) return {}
    return applyBeforeExport(plugins, formatData(state.definition, state.data, formatters), {
      definition: state.definition,
      source: state.source,
    })
  }

  function finishHydrate(patchedData: Record<string, unknown>) {
    state.data = cloneData(patchedData)
    refreshValidation()
  }

  function can(command: Command): boolean {
    switch (command.type) {
      case 'load':
      case 'hydrate':
        return command.source.kind === adapter.kind
      case 'insertField':
      case 'updateField':
      case 'removeField':
      case 'setValue':
      case 'setData':
      case 'resetData':
      case 'insertRow':
      case 'removeRow':
      case 'export':
        return Boolean(state.definition && state.source)
      default:
        return false
    }
  }

  async function dispatch(command: Command): Promise<DispatchResult> {
    if (!can(command)) {
      throw new Error(`cannot dispatch ${command.type}`)
    }

    switch (command.type) {
      case 'load': {
        applyMarkers(defaultMarkers)
        await adapter.load(command.source)
        const hash = await sourceHash(command.source)
        preview = adapter.getPreview()
        const fields: Field[] = []
        const names = new Set<string>()
        const discovered = applyAfterDiscover(plugins, await adapter.discoverFields(), command.source)
        for (const raw of discovered) {
          if (names.has(raw.name)) continue
          names.add(raw.name)
          const field: Field = { ...raw, id: raw.id ?? createId() }
          await adapter.insertAnchor(field)
          fields.push(field)
        }
        state = {
          definition: {
            version: 1,
            source: { kind: command.source.kind, hash },
            fields,
            markers: persistMarkers(),
          },
          data: {},
          source: { ...command.source, hash },
          validation: emptyValidation(),
        }
        const nextData = applyAfterHydrate(plugins, {
          definition: state.definition!,
          data: state.data,
          source: state.source!,
        })
        if (plugins.some((plugin) => plugin.afterHydrate) || nextData !== state.data) {
          finishHydrate(nextData)
          emit({ type: 'data-changed' })
        } else {
          emit({ type: 'state-changed' })
        }
        return { type: 'ok' }
      }

      case 'hydrate': {
        applyMarkers(command.definition.markers ?? DEFAULT_MARKERS)
        await adapter.load(command.source)
        const hash = command.source.hash ?? command.definition.source.hash
        preview = adapter.getPreview()
        const fields = command.definition.fields.map((f) => ({ ...f }))
        for (const field of fields) await adapter.insertAnchor(field)
        state = {
          definition: {
            ...command.definition,
            fields,
            markers: persistMarkers(),
          },
          data: cloneData(command.data ?? {}),
          source: { ...command.source, hash },
          validation: emptyValidation(),
        }
        finishHydrate(
          applyAfterHydrate(plugins, {
            definition: state.definition!,
            data: state.data,
            source: state.source!,
          }),
        )
        emit({ type: 'data-changed' })
        return { type: 'ok' }
      }

      case 'insertField': {
        const definition = requireDefinition()
        const field: Field = {
          ...command.field,
          id: command.field.id ?? createId(),
        }
        if (definition.fields.some((f) => f.name === field.name)) {
          throw new Error(`field name already exists: ${field.name}`)
        }
        await adapter.insertAnchor(field)
        definition.fields = [...definition.fields, field]
        refreshValidation()
        emit({ type: 'field-inserted', fieldId: field.id })
        return { type: 'ok' }
      }

      case 'updateField': {
        const definition = requireDefinition()
        const index = definition.fields.findIndex((f) => f.id === command.id)
        if (index < 0) throw new Error(`field not found: ${command.id}`)
        const next: Field = { ...definition.fields[index], ...command.patch, id: command.id }
        if (adapter.updateAnchor) await adapter.updateAnchor(next)
        definition.fields = definition.fields.map((f, i) => (i === index ? next : f))
        refreshValidation()
        emit({ type: 'field-updated', fieldId: command.id })
        return { type: 'ok' }
      }

      case 'removeField': {
        const definition = requireDefinition()
        const field = definition.fields.find((f) => f.id === command.id)
        if (!field) throw new Error(`field not found: ${command.id}`)
        await adapter.removeAnchor(command.id)
        definition.fields = definition.fields.filter((f) => f.id !== command.id)
        delete state.data[field.name]
        refreshValidation()
        emit({ type: 'field-removed', fieldId: command.id })
        return { type: 'ok' }
      }

      case 'setValue': {
        requireDefinition()
        state.data = setDataPath(state.data, command.path, command.value)
        refreshValidation()
        emit({ type: 'data-changed' })
        return { type: 'ok' }
      }

      case 'setData': {
        requireDefinition()
        state.data = cloneData(command.data)
        refreshValidation()
        emit({ type: 'data-changed' })
        return { type: 'ok' }
      }

      case 'resetData': {
        requireDefinition()
        state.data = {}
        refreshValidation()
        emit({ type: 'data-changed' })
        return { type: 'ok' }
      }

      case 'insertRow': {
        requireDefinition()
        state.data = insertTableRow(state.data, command.table, command.index, command.row)
        refreshValidation()
        emit({ type: 'data-changed' })
        return { type: 'ok' }
      }

      case 'removeRow': {
        requireDefinition()
        state.data = removeTableRow(state.data, command.table, command.index)
        refreshValidation()
        emit({ type: 'data-changed' })
        return { type: 'ok' }
      }

      case 'export': {
        const definition = requireDefinition()
        await adapter.bind(exportPayload())
        const format = command.format ?? adapter.kind
        const buffer = applyAfterExport(
          plugins,
          { buffer: await adapter.export(), format },
          { definition, source: state.source! },
        )
        emit({ type: 'exported', format, bytes: buffer.byteLength })
        return { type: 'exported', buffer, format }
      }
    }
  }

  return {
    getState: snapshot,
    getDefinition: () => snapshot().definition,
    getData: () => snapshot().data,
    getExportData: () => exportPayload(),
    getFormSchema: () => buildFormSchema(state),
    getView: () => buildView(state),
    getPreview: () => preview,
    getSource: () => state.source,
    getMarkers: () => ({ ...markers }),
    validate: () => validateState(state, validators),
    can,
    dispatch,
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    setViewport(next) {
      viewport = next
    },
  }
}
