import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import JSZip from 'jszip'
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { defineComponent, h, type VNodeChild } from 'vue'

import {
  Image,
  Master,
  Media,
  Presentation,
  Slide,
  Text,
  renderToBuffer,
  renderToModel,
  type PptxWarning,
  type WriteOptions,
} from '../../src'
import { loadSources } from '../../src/core/sources'

const PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=='
const PNG = `data:image/png;base64,${PNG_BASE64}`
const PNG_BYTES = Uint8Array.from(atob(PNG_BASE64), (char) => char.charCodeAt(0))
const JPEG = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////2Q=='

const deck = (children: () => VNodeChild, slideProps: Record<string, unknown> = {}) =>
  defineComponent(() => () => h(Presentation, null, () => h(Slide, slideProps, children)))

const picture = (src: string, x = 0) => h(Image, { src, x, y: 0, w: 1, h: 1 })

interface Written {
  slide: string
  media: Record<string, Uint8Array>
  warnings: PptxWarning[]
}

const write = async (component: ReturnType<typeof deck>, options: WriteOptions = {}): Promise<Written> => {
  const warnings: PptxWarning[] = []
  const bytes = await renderToBuffer(component, undefined, { onWarn: (warning) => warnings.push(warning), ...options })
  const zip = await JSZip.loadAsync(bytes)
  const media: Record<string, Uint8Array> = {}
  for (const [name, file] of Object.entries(zip.files)) {
    if (!file.dir && name.startsWith('ppt/media/')) media[name] = await file.async('uint8array')
  }
  return { slide: await zip.file('ppt/slides/slide1.xml')!.async('string'), media, warnings }
}

const pictures = (xml: string): number => xml.match(/<p:pic>/g)?.length ?? 0

const respond = (body: BodyInit | null, init: ResponseInit = {}) =>
  vi.fn<typeof fetch>(() => Promise.resolve(new Response(body, init)))

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('image sources', () => {
  it('leaves data URIs alone', async () => {
    const loadImage = vi.fn((src: string) => Promise.resolve(src))
    const model = await renderToModel(deck(() => picture(PNG)))
    expect(await loadSources(model, { loadImage })).toBe(model)
    const { slide, media } = await write(
      deck(() => picture(PNG)),
      { loadImage },
    )
    expect(loadImage).not.toHaveBeenCalled()
    expect(pictures(slide)).toBe(1)
    expect(Object.values(media)).toContainEqual(PNG_BYTES)
  })

  it('calls loadImage once per distinct source, for pictures, masters, backgrounds and covers', async () => {
    const loadImage = vi.fn((_src: string) => Promise.resolve(PNG))
    const Deck = defineComponent(
      () => () =>
        h(Presentation, null, () => [
          h(Master, { name: 'BRAND', background: 'brand.png' }, () => picture('logo.png')),
          h(Slide, { master: 'BRAND', background: 'brand.png' }, () => [
            picture('logo.png'),
            picture('logo.png', 1),
            picture('chart.png', 2),
            h(Text, { x: 0, y: 2, w: 3, h: 1 }, () => 'Caption'),
            h(Media, { src: 'data:video/mp4;base64,AAAA', cover: 'poster.png', x: 3, y: 0, w: 1, h: 1 }),
          ]),
        ]),
    )
    const { slide, warnings } = await write(Deck, { loadImage })
    expect(loadImage.mock.calls.map(([src]) => src).sort()).toEqual([
      'brand.png',
      'chart.png',
      'logo.png',
      'poster.png',
    ])
    expect(pictures(slide)).toBe(4)
    expect(slide).toContain('<p:bg>')
    expect(warnings).toEqual([])
  })

  it('skips an image that fails to load, with a warning', async () => {
    const loadImage = vi.fn((src: string) =>
      src === 'bad.png' ? Promise.reject(new Error('gone')) : Promise.resolve(PNG),
    )
    const { slide, warnings } = await write(
      deck(() => [picture('good.png'), picture('bad.png', 1)]),
      { loadImage },
    )
    expect(pictures(slide)).toBe(1)
    expect(warnings).toEqual([
      {
        code: 'load-failed',
        message: 'The image "bad.png" couldn\'t be loaded (gone); it is skipped.',
        slide: 1,
        object: '<Image>',
      },
    ])
  })

  it('rejects in strict mode', async () => {
    const loadImage = () => Promise.reject(new Error('gone'))
    await expect(
      renderToBuffer(
        deck(() => picture('bad.png')),
        undefined,
        { loadImage, strict: true },
      ),
    ).rejects.toThrow('load-failed: The image "bad.png" couldn\'t be loaded (gone); it is skipped. (slide 1, <Image>)')
  })

  it('accepts only base64 image data URIs from a loader', async () => {
    const loadImage = () => Promise.resolve('data:text/html,<p>no</p>')
    const { slide, warnings } = await write(
      deck(() => picture('page.png')),
      { loadImage },
    )
    expect(pictures(slide)).toBe(0)
    expect(warnings[0]!.message).toContain('did not return a base64 image data URI')
  })

  it('keeps the color of a background whose image fails, and drops a cover that fails', async () => {
    const loadImage = (src: string) => (src === 'poster.png' ? Promise.resolve(PNG) : Promise.reject(new Error('gone')))
    const model = await renderToModel(
      defineComponent(
        () => () =>
          h(Presentation, null, () => [
            h(Slide, { background: { color: '#ff0000', src: 'a.png' } }, () => [
              h(Media, { src: 'data:video/mp4;base64,AAAA', cover: 'missing.png', x: 0, y: 0, w: 1, h: 1 }),
              h(Media, { src: 'data:video/mp4;base64,AAAA', cover: 'poster.png', x: 1, y: 0, w: 1, h: 1 }),
            ]),
            h(Slide, { background: 'b.png' }),
          ]),
      ),
    )
    const loaded = await loadSources(model, { loadImage, onWarn: () => undefined })
    expect(loaded.slides[0]!.background).toEqual({ color: 'FF0000' })
    expect(loaded.slides[1]!.background).toBeUndefined()
    const [missing, poster] = loaded.slides[0]!.objects
    expect(missing).toMatchObject({ kind: 'media', options: { data: 'video/mp4;base64,AAAA' } })
    expect(missing).not.toHaveProperty('options.cover')
    expect(poster).toMatchObject({ kind: 'media', options: { cover: PNG } })
  })

  it('names a background file after its image type', async () => {
    const { media } = await write(deck(() => [], { background: JPEG }))
    expect(Object.keys(media).some((name) => name.endsWith('.jpeg'))).toBe(true)
  })

  it('leaves sources to PptxGenJS outside Node unless a loader is given', async () => {
    const model = await renderToModel(deck(() => picture('logo.png')))
    const release = Object.getOwnPropertyDescriptor(process, 'release')!
    Object.defineProperty(process, 'release', { value: { name: 'browser' }, configurable: true })
    try {
      expect(await loadSources(model, { allowFiles: false })).toBe(model)
      const loaded = await loadSources(model, { loadImage: () => Promise.resolve(PNG) })
      expect(loaded.slides[0]!.objects[0]).toMatchObject({ src: PNG, options: { data: PNG.slice(5) } })
    } finally {
      Object.defineProperty(process, 'release', release)
    }
  })
})

describe('files in Node', () => {
  let dir = ''
  let file = ''

  beforeAll(async () => {
    dir = await mkdtemp(join(tmpdir(), 'vue-pptx-document-'))
    file = join(dir, 'logo.PNG')
    await writeFile(file, PNG_BYTES)
  })

  afterAll(async () => {
    await rm(dir, { recursive: true, force: true })
  })

  it('reads a path by default', async () => {
    const { slide, media, warnings } = await write(deck(() => picture(file)))
    expect(warnings).toEqual([])
    expect(pictures(slide)).toBe(1)
    expect(Object.values(media)).toContainEqual(PNG_BYTES)
  })

  it('refuses paths with allowFiles: false', async () => {
    const { slide, media, warnings } = await write(
      deck(() => [picture(file), picture('/etc/passwd', 1), h(Media, { src: file, x: 2, y: 0, w: 1, h: 1 })]),
      { allowFiles: false },
    )
    expect(pictures(slide)).toBe(0)
    expect(media).toEqual({})
    expect(warnings.map(({ code, object }) => [code, object])).toEqual([
      ['load-failed', '<Image>'],
      ['load-failed', '<Image>'],
      ['load-failed', '<Media>'],
    ])
    expect(warnings[0]!.message).toContain('allowFiles: false')
  })

  it('refuses files that are not images', async () => {
    const { slide, warnings } = await write(deck(() => picture(join(dir, 'notes.txt'))))
    expect(pictures(slide)).toBe(0)
    expect(warnings[0]!.message).toContain('"txt" is not a known image file type')
  })

  it('reports a file that does not exist', async () => {
    const { warnings } = await write(deck(() => picture(join(dir, 'missing.png'))))
    expect(warnings[0]).toMatchObject({ code: 'load-failed', slide: 1, object: '<Image>' })
    expect(warnings[0]!.message).toContain('ENOENT')
  })
})

describe('URLs in Node', () => {
  const url = 'https://example.com/logo.png'

  it('fetches an image', async () => {
    const fetch = respond(PNG_BYTES, { headers: { 'content-type': 'image/png' } })
    vi.stubGlobal('fetch', fetch)
    const { slide, media, warnings } = await write(deck(() => [picture(url), picture('http://example.com/a.png', 1)]))
    expect(warnings).toEqual([])
    expect(pictures(slide)).toBe(2)
    expect(Object.values(media)).toContainEqual(PNG_BYTES)
    expect(fetch).toHaveBeenCalledWith(url, { signal: expect.any(AbortSignal) as AbortSignal })
  })

  it.each([
    [
      'an error status',
      respond('Not found', { status: 404, headers: { 'content-type': 'image/png' } }),
      'answered 404',
    ],
    ['a page', respond('<html></html>', { headers: { 'content-type': 'text/html; charset=utf-8' } }), 'text/html'],
    ['no content type', respond(PNG_BYTES), 'no content type'],
    [
      'a declared size over the limit',
      respond(PNG_BYTES, { headers: { 'content-type': 'image/png', 'content-length': String(11 * 1024 * 1024) } }),
      'larger than',
    ],
    [
      'a body over the limit',
      respond(new Uint8Array(10 * 1024 * 1024 + 1), { headers: { 'content-type': 'image/png' } }),
      'larger than',
    ],
    ['a network error', vi.fn<typeof fetch>(() => Promise.reject(new TypeError('fetch failed'))), 'fetch failed'],
  ])('skips %s', async (_case, fetch, reason) => {
    vi.stubGlobal('fetch', fetch)
    const { slide, warnings } = await write(deck(() => picture(url)))
    expect(pictures(slide)).toBe(0)
    expect(warnings).toHaveLength(1)
    expect(warnings[0]!.message).toContain(reason)
  })
})
