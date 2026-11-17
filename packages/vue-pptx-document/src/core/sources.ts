import { at, createContext, inContext, warn, type DiagnosticOptions } from './diagnostics'
import { isNode, nodeLoader, type ImageLoader } from './load-image'
import type { Background, DeckModel, MasterModel, SlideModel, SlideObject } from './model'

export interface SourceOptions {
  /**
   * Loads every picture, background and media cover that isn't a data URI yet, before
   * PptxGenJS sees it, and resolves to a base64 `data:image/...` URI. Called once per
   * distinct source. Without it, Node fetches http(s) URLs and reads other sources as file
   * paths, and the browser leaves them to PptxGenJS.
   */
  loadImage?: ImageLoader
  /**
   * Lets Node read images and media from the file system. On by default; turn it off when the
   * deck is built from input you don't trust.
   */
  allowFiles?: boolean
}

type Loaded = ReadonlyMap<string, string | Error>

const IMAGE_DATA_URI = /^data:image\/[\w.+-]+;base64,/i

const clip = (src: string): string => (src.length > 80 ? `${src.slice(0, 77)}...` : src)

/** Every image source of the deck that still has to be loaded. */
const imageSources = (model: DeckModel): Set<string> => {
  const sources = new Set<string>()
  const note = (src: string | undefined): void => {
    if (src && !src.startsWith('data:')) sources.add(src)
  }
  for (const page of [...model.masters, ...model.slides]) {
    note(page.background?.src)
    for (const object of page.objects) {
      if (object.kind === 'image') note(object.src)
      else if (object.kind === 'media') note(object.options.cover)
    }
  }
  return sources
}

const loadAll = async (sources: Set<string>, load: ImageLoader): Promise<Loaded> => {
  const entries = await Promise.all(
    [...sources].map(async (src): Promise<[string, string | Error]> => {
      try {
        const uri = await load(src)
        if (!IMAGE_DATA_URI.test(uri)) throw new Error('the loader did not return a base64 image data URI')
        return [src, uri]
      } catch (error) {
        return [src, error instanceof Error ? error : new Error(String(error))]
      }
    }),
  )
  return new Map(entries)
}

/** Puts the loaded data URIs in place of the sources, and drops what failed with a warning. */
const replaceSources = (model: DeckModel, loaded: Loaded, blockFiles: boolean): DeckModel => {
  const source = (src: string): string | undefined => {
    const result = loaded.get(src)
    if (result === undefined || typeof result === 'string') return result ?? src
    warn('load-failed', `The image "${clip(src)}" couldn't be loaded (${result.message}); it is skipped.`)
    return undefined
  }

  const background = (value: Background | undefined): Background | undefined => {
    if (!value?.src) return value
    const src = source(value.src)
    if (src) return { ...value, src }
    const { src: _src, ...rest } = value
    return rest.color === undefined ? undefined : rest
  }

  const object = (value: SlideObject): SlideObject | undefined => {
    if (value.kind === 'image') {
      const src = at({ object: '<Image>' }, () => source(value.src))
      if (src === value.src) return value
      if (!src) return undefined
      const { path: _path, ...options } = value.options
      return { ...value, src, options: { ...options, data: src.slice('data:'.length) } }
    }
    if (value.kind === 'media') {
      // PptxGenJS reads every media path that doesn't start with "http" from the disk.
      const { path, cover } = value.options
      if (blockFiles && path && !path.startsWith('http')) {
        at({ object: '<Media>' }, () =>
          warn('load-failed', `The media file "${clip(path)}" is skipped: reading files is turned off.`),
        )
        return undefined
      }
      const loadedCover = cover === undefined ? cover : at({ object: '<Media>' }, () => source(cover))
      if (loadedCover === cover) return value
      const { cover: _cover, ...options } = value.options
      return { ...value, options: loadedCover ? { ...options, cover: loadedCover } : options }
    }
    return value
  }

  const page = <P extends SlideModel | MasterModel>(value: P): P => ({
    ...value,
    background: background(value.background),
    objects: value.objects.flatMap((item) => object(item) ?? []),
  })

  return {
    ...model,
    masters: model.masters.map(page),
    slides: model.slides.map((slide, index) => at({ slide: index + 1 }, () => page(slide))),
  }
}

/**
 * Loads the deck's images into data URIs before PptxGenJS sees them, so PptxGenJS never reads
 * the file system or the network for them itself.
 */
export const loadSources = async (
  model: DeckModel,
  { loadImage, allowFiles = true, ...diagnostics }: SourceOptions & DiagnosticOptions,
): Promise<DeckModel> => {
  const node = isNode()
  const load = loadImage ?? (node ? nodeLoader(allowFiles) : undefined)
  const blockFiles = node && !allowFiles
  const sources = load ? imageSources(model) : new Set<string>()
  if (sources.size === 0 && !blockFiles) return model
  const loaded = load ? await loadAll(sources, load) : new Map<string, string | Error>()
  return inContext(createContext(diagnostics), () => replaceSources(model, loaded, blockFiles))
}
