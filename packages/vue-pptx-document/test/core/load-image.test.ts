import { afterEach, describe, expect, it, vi } from 'vitest'

import { FETCH_LIMITS, fetchImage, readImageFile } from '../../src/core/load-image'

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('fetchImage', () => {
  it('allows 10 seconds and 10 MB by default', () => {
    expect(FETCH_LIMITS).toEqual({ timeout: 10_000, maxBytes: 10 * 1024 * 1024 })
  })

  it('gives up when the server takes too long', async () => {
    vi.stubGlobal(
      'fetch',
      (_url: string, init: RequestInit) =>
        new Promise((_resolve, reject) => {
          init.signal!.addEventListener('abort', () => reject(init.signal!.reason as Error))
        }),
    )
    await expect(fetchImage('https://example.com/slow.png', { timeout: 10, maxBytes: 100 })).rejects.toThrow(
      'timed out',
    )
  })

  it('keeps the content type of the response', async () => {
    vi.stubGlobal('fetch', () =>
      Promise.resolve(new Response(new Uint8Array([1, 2, 3]), { headers: { 'content-type': 'Image/WebP' } })),
    )
    expect(await fetchImage('https://example.com/a.webp')).toBe('data:image/webp;base64,AQID')
  })
})

describe('readImageFile', () => {
  it('needs a runtime with Node built-ins', async () => {
    const builtin = Object.getOwnPropertyDescriptor(process, 'getBuiltinModule')!
    Object.defineProperty(process, 'getBuiltinModule', { value: undefined, configurable: true })
    try {
      await expect(readImageFile('logo.png')).rejects.toThrow('no file system access')
    } finally {
      Object.defineProperty(process, 'getBuiltinModule', builtin)
    }
  })
})
