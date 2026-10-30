export async function hashBytes(buffer: Uint8Array): Promise<string> {
  const subtle = globalThis.crypto?.subtle
  if (subtle) {
    const copy = new Uint8Array(buffer.byteLength)
    copy.set(buffer)
    const digest = await subtle.digest('SHA-256', copy)
    return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('')
  }
  let h = 2166136261
  for (const b of buffer) {
    h ^= b
    h = Math.imul(h, 16777619)
  }
  return (h >>> 0).toString(16).padStart(8, '0')
}

export async function hashText(text: string): Promise<string> {
  return hashBytes(new TextEncoder().encode(text))
}

export async function hashValue(value: unknown): Promise<string> {
  if (value instanceof Uint8Array) return hashBytes(value)
  if (typeof value === 'string') return hashText(value)
  let json: string | undefined
  try {
    json = JSON.stringify(value)
  } catch {
    throw new Error('custom source requires hash when data is not JSON-serializable')
  }
  if (json === undefined) {
    throw new Error('custom source requires hash when data is not JSON-serializable')
  }
  return hashText(json)
}
