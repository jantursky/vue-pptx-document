import type PptxGenJS from 'pptxgenjs'

import { at, createContext, inContext, locate, warn, type DiagnosticOptions } from './diagnostics'
import type { Background, DeckModel, SlideModel, SlideObject } from './model'

type Pptx = PptxGenJS
type Slide = PptxGenJS.Slide

let engine: Promise<new () => Pptx> | undefined

/** PptxGenJS loads on the first write, so a preview never pays for it. */
export const loadPptxGenJS = (): Promise<new () => Pptx> => {
  engine ??= import('pptxgenjs').then((module) => {
    const exported = module as unknown as { default: unknown }
    // The CommonJS build, as some bundlers and Node hand it over, nests the class once more.
    const candidate = (exported.default as { default?: unknown } | undefined)?.default ?? exported.default
    return candidate as new () => Pptx
  })
  return engine
}

/** PptxGenJS writes into the option objects it gets, so each call gets its own copy. */
const copy = <T>(value: T): T => (value === undefined ? value : (JSON.parse(JSON.stringify(value)) as T))

const background = (value: Background | undefined): PptxGenJS.BackgroundProps | undefined => {
  if (!value) return undefined
  if (value.src) return value.src.startsWith('data:') ? { data: value.src.slice(5) } : { path: value.src }
  return { color: value.color, transparency: value.transparency }
}

/** An object whose XML needs a fix after PptxGenJS wrote it, found again by a temporary name. */
interface Patch {
  kind: 'image' | 'table'
  /** The temporary `objectName`. */
  name: string
  /** The `objectName` the author set, if any. */
  original?: string
}

interface Writing {
  lang?: string
  patches: Patch[]
}

/** Patches wait here between building the presentation and writing it. */
const pending = new WeakMap<Pptx, Patch[]>()

/** Marks the object so `finalize` can find its XML, and returns the options to hand to PptxGenJS. */
const flag = <T extends { objectName?: string }>(writing: Writing, kind: Patch['kind'], options: T): T => {
  const name = `pptx-a11y-${writing.patches.length + 1}`
  writing.patches.push({ kind, name, ...(options.objectName ? { original: options.objectName } : {}) })
  return { ...options, objectName: name }
}

/** Text without a language of its own gets the deck's. */
const withLang = <T extends object>(writing: Writing, options: T): T =>
  writing.lang ? { lang: writing.lang, ...options } : options

const add = (slide: Slide, object: SlideObject, writing: Writing): void => {
  const position = { ...object.frame }
  switch (object.kind) {
    case 'text':
      slide.addText(copy(object.runs), withLang(writing, { ...copy(object.options), ...position }))
      break
    case 'shape':
      slide.addShape(object.shape as PptxGenJS.SHAPE_NAME, { ...copy(object.options), ...position })
      break
    case 'image': {
      const { decorative, ...options } = copy(object.options) as PptxGenJS.ImageProps & { decorative?: boolean }
      // PptxGenJS describes an image without altText by its path or data, so those get an empty
      // description after writing.
      const described = decorative !== true && Boolean(options.altText)
      slide.addImage({ ...(described ? options : flag(writing, 'image', options)), ...position })
      break
    }
    case 'table': {
      const { headerRows, ...options } = copy(object.options) as PptxGenJS.TableProps & { headerRows?: number }
      const rows = copy(object.rows)
      if (writing.lang) {
        for (const row of rows) for (const cell of row) cell.options = withLang(writing, cell.options ?? {})
      }
      const flagged = typeof headerRows === 'number' && headerRows > 0 ? flag(writing, 'table', options) : options
      slide.addTable(rows, { ...flagged, ...position })
      break
    }
    case 'chart':
      if (Array.isArray(object.type)) {
        // For combo charts PptxGenJS reads the options from the second argument.
        slide.addChart(copy(object.type) as PptxGenJS.IChartMulti[], { ...copy(object.options), ...position } as never)
      } else {
        slide.addChart(object.type as PptxGenJS.CHART_NAME, copy(object.data), { ...copy(object.options), ...position })
      }
      break
    case 'media':
      slide.addMedia({ ...copy(object.options), ...position })
      break
    case 'placeholder':
      // Placeholders belong to masters; on a slide they're plain text boxes.
      slide.addText(object.text ?? '', withLang(writing, { ...copy(object.options), ...position }))
      break
  }
}

type MasterObject = NonNullable<PptxGenJS.SlideMasterProps['objects']>[number]

/** PptxGenJS masters take text, rectangles, lines, images, charts and placeholders. */
const masterObject = (object: SlideObject, writing: Writing): MasterObject | undefined => {
  const position = { ...object.frame }
  switch (object.kind) {
    case 'text': {
      // A master's text box holds one string, so the runs' own formatting can't vary; the first
      // run's options apply to the whole box.
      const runs = copy(object.runs)
      const { breakLine: _breakLine, ...first } = (runs[0]?.options ?? {}) as Record<string, unknown>
      const text = runs.map((run) => run.text + (run.options?.breakLine ? '\n' : '')).join('')
      return { text: { text, options: withLang(writing, { ...first, ...copy(object.options), ...position }) } }
    }
    case 'chart':
      // Combo charts only go through slide.addChart.
      if (Array.isArray(object.type)) return undefined
      return {
        chart: { type: object.type, data: copy(object.data), opts: { ...copy(object.options), ...position } },
      } as never
    case 'shape':
      if (object.shape === 'line') return { line: { ...copy(object.options), ...position } }
      if (object.shape === 'rect') return { rect: { ...copy(object.options), ...position } }
      return undefined
    case 'image':
      return { image: { ...copy(object.options), ...position } }
    case 'placeholder':
      return {
        placeholder: {
          options: { ...copy(object.options), ...position },
          ...(object.text ? { text: object.text } : {}),
        },
      }
    default:
      return undefined
  }
}

const NAMES: Record<SlideObject['kind'], string> = {
  text: '<Text>',
  shape: '<Shape>',
  image: '<Image>',
  table: '<Table>',
  chart: '<Chart>',
  media: '<Media>',
  placeholder: '<Placeholder>',
}

/** Builds a PptxGenJS presentation from the model. */
export const toPptxGenJS = async (model: DeckModel, options: DiagnosticOptions = {}): Promise<Pptx> => {
  const PptxGenJS = await loadPptxGenJS()
  return inContext(createContext(options), () => build(new PptxGenJS(), model))
}

const build = (deck: Pptx, model: DeckModel): Pptx => {
  const writing: Writing = { patches: [], ...(model.lang ? { lang: model.lang } : {}) }
  pending.set(deck, writing.patches)
  if (model.layout.name === 'CUSTOM') {
    deck.defineLayout({ name: 'CUSTOM', width: model.layout.width, height: model.layout.height })
  }
  deck.layout = model.layout.name
  const { meta } = model
  // PptxGenJS fills in "PptxGenJS Presentation" and "PptxGenJS" otherwise.
  deck.title = meta.title ?? ''
  deck.subject = meta.subject ?? ''
  deck.company = meta.company ?? ''
  deck.author = meta.author ?? ''
  if (meta.revision) deck.revision = meta.revision
  if (model.rtl) deck.rtlMode = true
  if (model.theme) deck.theme = { ...model.theme }

  for (const master of model.masters) {
    const objects: MasterObject[] = []
    for (const object of master.objects) {
      const converted = masterObject(object, writing)
      if (converted) objects.push(converted)
      else
        at({ object: NAMES[object.kind] }, () =>
          warn('master-object-dropped', `A ${object.kind} can't be part of a master; it is skipped.`),
        )
    }
    const props: PptxGenJS.SlideMasterProps = { title: master.name, objects }
    const bg = background(master.background)
    if (bg) props.background = bg
    if (master.margin !== undefined) props.margin = master.margin
    if (master.slideNumber) props.slideNumber = copy(master.slideNumber)
    deck.defineSlideMaster(props)
  }

  const sections = new Set<string>()
  model.slides.forEach((model_, index) => at({ slide: index + 1 }, () => addSlide(model_)))
  return deck

  function addSlide(model_: SlideModel): void {
    if (model_.section && !sections.has(model_.section)) {
      sections.add(model_.section)
      deck.addSection({ title: model_.section })
    }
    const slide = deck.addSlide({
      ...(model_.master ? { masterName: model_.master } : {}),
      ...(model_.section ? { sectionTitle: model_.section } : {}),
    })
    const bg = background(model_.background)
    if (bg) slide.background = bg
    if (model_.color) slide.color = model_.color
    if (model_.hidden) slide.hidden = true
    if (model_.slideNumber) slide.slideNumber = copy(model_.slideNumber)
    for (const object of model_.objects) at({ object: NAMES[object.kind] }, () => add(slide, object, writing))
    if (model_.notes) slide.addNotes(model_.notes)
  }
}

export type OutputType = 'blob' | 'arraybuffer' | 'uint8array' | 'base64' | 'nodebuffer'

export interface WriteOptions extends DiagnosticOptions {
  /** Deflate the zip. On by default; turn it off for faster writes of large decks. */
  compression?: boolean
}

/**
 * PptxGenJS reads images and media while it writes the zip, so its errors don't say which object
 * failed. They name the path, though, which finds the object.
 */
const locateSource = (error: unknown, model: DeckModel): Error => {
  const message = error instanceof Error ? error.message : String(error)
  for (const [index, slide] of model.slides.entries()) {
    for (const object of slide.objects) {
      const source =
        object.kind === 'image'
          ? object.src
          : object.kind === 'media'
            ? (object.options.path ?? object.options.link)
            : undefined
      if (typeof source === 'string' && source !== '' && !source.startsWith('data:') && message.includes(source)) {
        return locate(error, { slide: index + 1, object: NAMES[object.kind] })
      }
    }
  }
  return locate(error, {})
}

const escapeXml = (value: string): string =>
  value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

/**
 * Fixes what PptxGenJS can't write: empty descriptions for decorative images (it falls back to the
 * image's path or data) and the header row flag of tables. The zip is only opened when a deck has such objects.
 */
const finalize = async (bytes: Uint8Array, patches: Patch[], outputType: OutputType, compression: boolean) => {
  const { default: JSZip } = await import('jszip')
  const zip = await JSZip.loadAsync(bytes)
  const kinds = { image: 'Picture', table: 'Table' }
  for (const path of Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))) {
    let xml = (await zip.file(path)?.async('string')) ?? ''
    const before = xml
    for (const patch of patches) {
      const name = `name="${patch.name}"`
      if (!xml.includes(name)) continue
      const restore = (id: string): string =>
        `name="${patch.original ? escapeXml(patch.original) : `${kinds[patch.kind]} ${Number(id) - 1}`}"`
      if (patch.kind === 'image') {
        xml = xml.replace(
          new RegExp(`<p:cNvPr id="(\\d+)" ${name} descr="[^"]*"`),
          (_match, id: string) => `<p:cNvPr id="${id}" ${restore(id)} descr=""`,
        )
      } else {
        xml = xml.replace(
          new RegExp(`<p:cNvPr id="(\\d+)" ${name}([\\s\\S]*?)<a:tblPr`),
          (_match, id: string, middle: string) => `<p:cNvPr id="${id}" ${restore(id)}${middle}<a:tblPr firstRow="1"`,
        )
      }
    }
    if (xml !== before) zip.file(path, xml)
  }
  return zip.generateAsync({ type: outputType, compression: compression ? 'DEFLATE' : 'STORE' })
}

export const writeModel = async (model: DeckModel, outputType: OutputType, options: WriteOptions = {}) => {
  const { compression, ...diagnostics } = options
  const deck = await toPptxGenJS(model, diagnostics)
  const patches = pending.get(deck) ?? []
  try {
    if (patches.length === 0) return await deck.write({ outputType, compression: compression ?? true })
    const bytes = (await deck.write({ outputType: 'uint8array', compression: compression ?? true })) as Uint8Array
    return await finalize(bytes, patches, outputType, compression ?? true)
  } catch (error) {
    throw locateSource(error, model)
  }
}
