import JSZip from 'jszip'

import type { DeckModel, SlideObject, TextRun } from '../core/model'
import { parsePptx, type ParseOptions, type PptxInput } from '../import/parse'
import { all, attr, child, parseXml } from '../import/xml'

export type { PptxInput } from '../import/parse'

/** One slide of a deck read back from a .pptx file. */
export interface PptxSlideReadout {
  /** The slide's XML part, as written in the file. */
  xml: string
  /** The slide's texts in reading order: object order, which is z-order. Paragraphs are joined by `\n`. */
  texts: string[]
  /** The speaker notes, or an empty string. */
  notes: string
}

/** A .pptx file read back for assertions. */
export interface PptxReadout {
  /** The deck the importer reads, in the model the writer and the preview use. */
  model: DeckModel
  slides: PptxSlideReadout[]
  /** Things the importer skipped or approximated. */
  warnings: string[]
}

const runsText = (runs: TextRun[]): string =>
  runs
    .map((run) => (run.options?.breakLine ? `${run.text}\n` : run.text))
    .join('')
    .replace(/\n$/, '')

const objectTexts = (object: SlideObject): string[] => {
  switch (object.kind) {
    case 'text':
      return [runsText(object.runs)]
    case 'table':
      return object.rows.flatMap((row) =>
        row.map((cell) => (typeof cell.text === 'string' ? cell.text : runsText(cell.text ?? []))),
      )
    case 'placeholder':
      return object.text ? [object.text] : []
    default:
      return []
  }
}

/** The slide parts of a package, in presentation order. */
const slideParts = async (zip: JSZip): Promise<string[]> => {
  const presentation = parseXml((await zip.file('ppt/presentation.xml')?.async('string')) ?? '')
  const rels = parseXml((await zip.file('ppt/_rels/presentation.xml.rels')?.async('string')) ?? '')
  const targets = new Map(all(rels, 'Relationship').map((rel) => [attr(rel, 'Id') ?? '', attr(rel, 'Target') ?? '']))
  const parts: string[] = []
  for (const id of all(child(presentation, 'p:sldIdLst'), 'p:sldId')) {
    const target = targets.get(attr(id, 'r:id') ?? '')
    if (target) parts.push(target.startsWith('/') ? target.slice(1) : `ppt/${target}`)
  }
  return parts
}

/**
 * Reads a .pptx file for assertions: the deck model, and for each slide its XML, texts and notes.
 * It works in Node without a DOM and accepts the same inputs as the importer.
 */
export const readPptx = async (input: PptxInput, options: ParseOptions = {}): Promise<PptxReadout> => {
  // A slide's own texts only, unless the caller asks for the inherited layout and master shapes.
  const { model, warnings } = await parsePptx(input, { masterShapes: false, ...options })
  const bytes = typeof Blob !== 'undefined' && input instanceof Blob ? await input.arrayBuffer() : input
  const zip = await JSZip.loadAsync(bytes)
  const parts = await slideParts(zip)
  const slides: PptxSlideReadout[] = []
  for (const [index, slide] of model.slides.entries()) {
    const part = parts[index]
    slides.push({
      xml: part ? ((await zip.file(part)?.async('string')) ?? '') : '',
      texts: slide.objects.flatMap(objectTexts),
      notes: slide.notes ?? '',
    })
  }
  return { model, slides, warnings }
}

/** The 1-based numbers of the slides that hold the text: a substring, or a match for a RegExp. */
export const findText = (deck: Pick<PptxReadout, 'slides'>, text: string | RegExp): number[] => {
  const matches = (value: string): boolean => {
    if (typeof text === 'string') return value.includes(text)
    // A global or sticky RegExp keeps state between calls to test().
    text.lastIndex = 0
    return text.test(value)
  }
  const found: number[] = []
  for (const [index, slide] of deck.slides.entries()) {
    if (slide.texts.some(matches)) found.push(index + 1)
  }
  return found
}

const round = (value: number): number => Math.round(value * 100) / 100

/** The size of a data URI payload in bytes, or the path itself. */
const imageLabel = (src: string): string => {
  const payload = /^data:[^,]*;base64,(.*)$/s.exec(src)?.[1]
  if (payload === undefined) return src.startsWith('data:') ? `${src.length} bytes` : src
  const padding = payload.endsWith('==') ? 2 : payload.endsWith('=') ? 1 : 0
  return `${Math.floor((payload.length * 3) / 4) - padding} bytes`
}

const outlineDetail = (object: SlideObject): string => {
  switch (object.kind) {
    case 'image':
      return imageLabel(object.src)
    case 'shape':
      return object.shape
    case 'chart':
      return typeof object.type === 'string' ? object.type : 'combo'
    case 'placeholder':
      return object.name
    default:
      return objectTexts(object).join(' | ')
  }
}

const outlineLine = (object: SlideObject): string => {
  const { x, y, w, h } = object.frame
  const head = `${object.kind} @ ${round(x)},${round(y)} ${round(w)}×${round(h)}`
  const detail = outlineDetail(object)
  return detail ? `  ${head}: ${detail.replace(/\n/g, '\\n')}` : `  ${head}`
}

/**
 * A stable plain-text outline of the deck for snapshot tests: one line per object with its
 * kind, frame (rounded to 0.01 inch) and text. It holds no ids, dates or image data.
 */
export const toOutline = (deck: Pick<PptxReadout, 'model' | 'slides'>): string => {
  const lines: string[] = []
  for (const [index, slide] of deck.model.slides.entries()) {
    lines.push(`Slide ${index + 1}${slide.title ? ` "${slide.title}"` : ''}${slide.hidden ? ' (hidden)' : ''}`)
    for (const object of slide.objects) lines.push(outlineLine(object))
    const notes = deck.slides[index]?.notes
    if (notes) lines.push(`  notes: ${notes.replace(/\n/g, '\\n')}`)
  }
  return lines.join('\n')
}
