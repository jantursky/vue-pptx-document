import type * as FsPromises from 'node:fs/promises'

/** Turns an image source, such as a URL or a path, into a base64 `data:image/...` URI. */
export type ImageLoader = (src: string) => Promise<string>

export interface FetchLimits {
  /** Milliseconds for the whole request, body included. */
  timeout: number
  maxBytes: number
}

export const FETCH_LIMITS: FetchLimits = { timeout: 10_000, maxBytes: 10 * 1024 * 1024 }

const FILE_TYPES: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  svg: 'image/svg+xml',
  webp: 'image/webp',
}

interface Runtime {
  process?: {
    versions?: { node?: string }
    release?: { name?: string }
    getBuiltinModule?: (id: 'node:fs/promises') => typeof FsPromises
  }
}

const runtime = (): Runtime['process'] => (globalThis as Runtime).process

/** The test PptxGenJS uses to decide whether it reads paths from the file system. */
export const isNode = (): boolean => Boolean(runtime()?.versions?.node) && runtime()?.release?.name === 'node'

const toBase64 = (bytes: Uint8Array): string => {
  let binary = ''
  // Chunks keep the spread under the engines' argument limit.
  for (let start = 0; start < bytes.length; start += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(start, start + 0x8000))
  }
  return btoa(binary)
}

const toDataUri = (type: string, bytes: Uint8Array): string => `data:${type};base64,${toBase64(bytes)}`

const tooLarge = (maxBytes: number): Error => new Error(`it is larger than ${maxBytes} bytes`)

/** Reads the body, stopping as soon as it passes `maxBytes`, whatever `content-length` claims. */
const readCapped = async (response: Response, maxBytes: number): Promise<Uint8Array> => {
  if (Number(response.headers.get('content-length')) > maxBytes) throw tooLarge(maxBytes)
  const chunks: Uint8Array[] = []
  const reader = response.body?.getReader()
  let size = 0
  for (;;) {
    const chunk = await reader?.read()
    if (!chunk || chunk.done) break
    size += chunk.value.length
    if (size > maxBytes) {
      await reader?.cancel()
      throw tooLarge(maxBytes)
    }
    chunks.push(chunk.value)
  }
  return new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer())
}

/** Fetches an image URL, following redirects, and accepts only a successful `image/*` response. */
export const fetchImage = async (url: string, limits: FetchLimits = FETCH_LIMITS): Promise<string> => {
  const response = await fetch(url, { signal: AbortSignal.timeout(limits.timeout) })
  if (!response.ok) throw new Error(`the server answered ${response.status}`)
  const type = response.headers.get('content-type')?.split(';')[0]?.trim().toLowerCase() ?? ''
  if (!type.startsWith('image/')) throw new Error(`it is not an image (${type || 'no content type'})`)
  return toDataUri(type, await readCapped(response, limits.maxBytes))
}

/**
 * Reads an image file in Node. `getBuiltinModule` rather than an import keeps `node:fs` out of
 * browser bundles, which bundlers would otherwise try to resolve.
 */
export const readImageFile = async (path: string): Promise<string> => {
  const extension = /\.(\w+)$/.exec(path)?.[1]?.toLowerCase() ?? ''
  const type = FILE_TYPES[extension]
  if (!type) throw new Error(`"${extension || path}" is not a known image file type`)
  const fs = runtime()?.getBuiltinModule?.('node:fs/promises')
  if (!fs) throw new Error('this runtime has no file system access')
  return toDataUri(type, await fs.readFile(path))
}

/** Node's loader when none is given: http(s) URLs are fetched, anything else is a file path. */
export const nodeLoader =
  (allowFiles: boolean): ImageLoader =>
  (src) => {
    if (/^https?:\/\//i.test(src)) return fetchImage(src)
    if (!allowFiles) return Promise.reject(new Error('reading files is turned off (allowFiles: false)'))
    return readImageFile(src)
  }
