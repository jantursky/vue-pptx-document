import JSZip from 'jszip'

import type {
  Background,
  ChartName,
  ChartSeries,
  ComboChart,
  DeckModel,
  Frame,
  SlideModel,
  SlideObject,
  TableCellModel,
  TextRun,
} from '../core/model'
import { LAYOUTS } from '../core/resolve'
import { EMU_PER_IN, round } from '../core/units'
import { all, attr, child, childrenOf, descendants, numAttr, parseXml, path, tagOf, textOf, type XmlNode } from './xml'

type Options = Record<string, unknown>

export interface ParseOptions {
  /** Copy the shapes and pictures of each slide's layout and master onto the slide. On by default. */
  masterShapes?: boolean
  /** Embed pictures as data URIs. Off: pictures become `media/<name>` paths. On by default. */
  embedImages?: boolean
}

export interface ParseResult {
  model: DeckModel
  /** Things the importer skipped or approximated. */
  warnings: string[]
}

export type PptxInput = ArrayBuffer | Uint8Array | Blob

const EMU_PER_PT = 12700

const inches = (emu: number | undefined): number => round((emu ?? 0) / EMU_PER_IN)

const MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/emf',
  wmf: 'image/wmf',
  mp4: 'video/mp4',
  m4v: 'video/mp4',
  mov: 'video/quicktime',
  mp3: 'audio/mpeg',
  wav: 'audio/wav',
}

/** Resolves a relationship target against the part that refers to it. */
const resolvePart = (from: string, target: string): string => {
  if (target.startsWith('/')) return target.slice(1)
  const parts = from.split('/').slice(0, -1)
  for (const segment of target.split('/')) {
    if (segment === '..') parts.pop()
    else if (segment !== '.') parts.push(segment)
  }
  return parts.join('/')
}

const relsPath = (part: string): string => {
  const parts = part.split('/')
  const file = parts.pop() ?? ''
  return [...parts, '_rels', `${file}.rels`].join('/')
}

interface Relationship {
  type: string
  target: string
  external: boolean
}

// Color helpers ------------------------------------------------------------------------------

const toHsl = (hex: string): [number, number, number] => {
  const [r, g, b] = [0, 2, 4].map((offset) => parseInt(hex.slice(offset, offset + 2), 16) / 255) as [
    number,
    number,
    number,
  ]
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  if (max === min) return [0, 0, l]
  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  const hue = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
  return [hue / 6, s, l]
}

const fromHsl = ([hue, s, l]: [number, number, number]): string => {
  const channel = (t: number): number => {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s
    const p = 2 * l - q
    let x = t
    if (x < 0) x += 1
    if (x > 1) x -= 1
    if (x < 1 / 6) return p + (q - p) * 6 * x
    if (x < 1 / 2) return q
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6
    return p
  }
  const values = s === 0 ? [l, l, l] : [channel(hue + 1 / 3), channel(hue), channel(hue - 1 / 3)]
  return values
    .map((value) =>
      Math.round(Math.max(0, Math.min(1, value)) * 255)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')
    .toUpperCase()
}

const PRESET_COLORS: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  gray: '808080',
  darkGray: 'A9A9A9',
  lightGray: 'D3D3D3',
  orange: 'FFA500',
}

interface ColorResult {
  color: string
  transparency?: number
}

class Importer {
  readonly warnings: string[] = []
  private readonly rels = new Map<string, Map<string, Relationship>>()
  private readonly parts = new Map<string, XmlNode | undefined>()
  private readonly images = new Map<string, string>()
  private theme: Partial<Record<string, string>> = {}

  constructor(
    private readonly zip: JSZip,
    private readonly options: Required<ParseOptions>,
  ) {}

  warn(message: string): void {
    if (!this.warnings.includes(message)) this.warnings.push(message)
  }

  async xml(part: string): Promise<XmlNode | undefined> {
    if (!this.parts.has(part)) {
      const file = this.zip.file(part)
      this.parts.set(part, file ? parseXml(await file.async('string')) : undefined)
    }
    return this.parts.get(part)
  }

  async relationships(part: string): Promise<Map<string, Relationship>> {
    const cached = this.rels.get(part)
    if (cached) return cached
    const map = new Map<string, Relationship>()
    const root = await this.xml(relsPath(part))
    for (const rel of all(root, 'Relationship')) {
      const id = attr(rel, 'Id')
      const target = attr(rel, 'Target')
      if (!id || !target) continue
      const external = attr(rel, 'TargetMode') === 'External'
      map.set(id, { type: attr(rel, 'Type') ?? '', target: external ? target : resolvePart(part, target), external })
    }
    this.rels.set(part, map)
    return map
  }

  async relatedOfType(part: string, type: string): Promise<string | undefined> {
    for (const rel of (await this.relationships(part)).values()) {
      if (rel.type.endsWith(`/${type}`)) return rel.target
    }
    return undefined
  }

  async image(part: string): Promise<string> {
    const cached = this.images.get(part)
    if (cached) return cached
    const extension = part.split('.').pop()?.toLowerCase() ?? ''
    let src: string
    if (this.options.embedImages) {
      const file = this.zip.file(part)
      const base64 = file ? await file.async('base64') : ''
      src = `data:${MIME[extension] ?? 'application/octet-stream'};base64,${base64}`
    } else src = part.replace(/^ppt\//, '')
    if (extension === 'emf' || extension === 'wmf')
      this.warn(`${part} is a ${extension.toUpperCase()} picture; browsers can't show it in the preview.`)
    this.images.set(part, src)
    return src
  }

  async loadTheme(presentationPart: string): Promise<void> {
    const masterPart = await this.relatedOfType(presentationPart, 'slideMaster')
    const themePart =
      (masterPart && (await this.relatedOfType(masterPart, 'theme'))) ??
      (await this.relatedOfType(presentationPart, 'theme'))
    if (!themePart) return
    const scheme = descendants(await this.xml(themePart), 'a:clrScheme')[0]
    for (const slot of childrenOf(scheme)) {
      const name = tagOf(slot)?.replace(/^a:/, '')
      const value = child(slot, 'a:srgbClr')
        ? attr(child(slot, 'a:srgbClr'), 'val')
        : attr(child(slot, 'a:sysClr'), 'lastClr')
      if (name && value) this.theme[name] = value.toUpperCase()
    }
    this.theme.tx1 ??= this.theme.dk1 ?? '000000'
    this.theme.bg1 ??= this.theme.lt1 ?? 'FFFFFF'
    this.theme.tx2 ??= this.theme.dk2 ?? '44546A'
    this.theme.bg2 ??= this.theme.lt2 ?? 'E7E6E6'
  }

  /** A color element (`a:srgbClr`, `a:schemeClr`, ...) with its modifiers applied. */
  color(node: XmlNode | undefined): ColorResult | undefined {
    const element = childrenOf(node).find((item) =>
      /^a:(srgbClr|schemeClr|sysClr|prstClr|scrgbClr)$/.test(tagOf(item) ?? ''),
    )
    if (!element) return undefined
    const tag = tagOf(element)
    let hex: string | undefined
    if (tag === 'a:srgbClr') hex = attr(element, 'val')
    else if (tag === 'a:schemeClr')
      hex =
        this.theme[attr(element, 'val') ?? ''] ??
        this.theme[(attr(element, 'val') ?? '').replace('tx', 'dk').replace('bg', 'lt')]
    else if (tag === 'a:sysClr') hex = attr(element, 'lastClr')
    else if (tag === 'a:prstClr') hex = PRESET_COLORS[attr(element, 'val') ?? '']
    if (!hex) return undefined
    hex = hex.toUpperCase()
    let transparency: number | undefined
    const modifiers = childrenOf(element)
    const value = (name: string): number | undefined => {
      const modifier = modifiers.find((item) => tagOf(item) === name)
      return modifier ? numAttr(modifier, 'val') : undefined
    }
    const lumMod = value('a:lumMod')
    const lumOff = value('a:lumOff')
    const tint = value('a:tint')
    const shade = value('a:shade')
    if (lumMod !== undefined || lumOff !== undefined || tint !== undefined || shade !== undefined) {
      const hsl = toHsl(hex)
      if (lumMod !== undefined) hsl[2] *= lumMod / 100000
      if (lumOff !== undefined) hsl[2] += lumOff / 100000
      if (tint !== undefined) hsl[2] = hsl[2] + (1 - hsl[2]) * (1 - tint / 100000)
      if (shade !== undefined) hsl[2] *= shade / 100000
      hsl[2] = Math.max(0, Math.min(1, hsl[2]))
      hex = fromHsl(hsl)
    }
    const alpha = value('a:alpha')
    if (alpha !== undefined && alpha < 100000) transparency = Math.round(100 - alpha / 1000)
    return transparency === undefined ? { color: hex } : { color: hex, transparency }
  }

  /** `a:solidFill` (or the first gradient stop) under `node`. */
  fill(node: XmlNode | undefined): ColorResult | 'none' | undefined {
    if (child(node, 'a:noFill')) return 'none'
    const solid = child(node, 'a:solidFill')
    if (solid) return this.color(solid)
    const gradient = child(node, 'a:gradFill')
    if (gradient) {
      this.warn('Gradient fills are imported as their first color.')
      return this.color(descendants(gradient, 'a:gs')[0])
    }
    if (child(node, 'a:blipFill') || child(node, 'a:pattFill')) this.warn('Picture and pattern fills are not imported.')
    return undefined
  }

  line(spPr: XmlNode | undefined): Options | undefined {
    const ln = child(spPr, 'a:ln')
    if (!ln) return undefined
    const fill = this.fill(ln)
    // Without a fill of its own, an outline comes from the shape style, which PptxGenJS doesn't write.
    if (fill === 'none' || fill === undefined) return undefined
    const width = numAttr(ln, 'w')
    const dash = attr(child(ln, 'a:prstDash'), 'val')
    const out: Options = { ...fill }
    if (width !== undefined) out.width = round(width / EMU_PER_PT)
    if (dash && dash !== 'solid') out.dashType = dash
    const head = attr(child(ln, 'a:headEnd'), 'type')
    const tail = attr(child(ln, 'a:tailEnd'), 'type')
    if (head && head !== 'none') out.beginArrowType = head
    if (tail && tail !== 'none') out.endArrowType = tail
    return out
  }
}

type Transform = (frame: Frame) => Frame

const identity: Transform = (frame) => frame

const xfrmFrame = (xfrm: XmlNode | undefined): Frame | undefined => {
  const off = child(xfrm, 'a:off')
  const ext = child(xfrm, 'a:ext')
  if (!off || !ext) return undefined
  return {
    x: inches(numAttr(off, 'x')),
    y: inches(numAttr(off, 'y')),
    w: inches(numAttr(ext, 'cx')),
    h: inches(numAttr(ext, 'cy')),
  }
}

const geometryOptions = (xfrm: XmlNode | undefined): Options => {
  const out: Options = {}
  const rot = numAttr(xfrm, 'rot')
  if (rot) out.rotate = round(rot / 60000)
  if (attr(xfrm, 'flipH') === '1') out.flipH = true
  if (attr(xfrm, 'flipV') === '1') out.flipV = true
  return out
}

const placeholderKey = (sp: XmlNode | undefined): { type: string; idx: string } | undefined => {
  const ph = descendants(sp, 'p:ph')[0]
  if (!ph) return undefined
  return { type: attr(ph, 'type') ?? 'body', idx: attr(ph, 'idx') ?? '' }
}

const ALIGN: Record<string, string> = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }
const ANCHOR: Record<string, string> = { t: 'top', ctr: 'middle', b: 'bottom' }

/** Character properties of an `a:rPr` or `a:defRPr`. */
const runProps = (importer: Importer, rPr: XmlNode | undefined, relationships: Map<string, Relationship>): Options => {
  const options: Options = {}
  if (!rPr) return options
  const size = numAttr(rPr, 'sz')
  if (size) options.fontSize = size / 100
  if (attr(rPr, 'b') === '1') options.bold = true
  if (attr(rPr, 'i') === '1') options.italic = true
  const underline = attr(rPr, 'u')
  if (underline && underline !== 'none') options.underline = { style: underline }
  const strike = attr(rPr, 'strike')
  if (strike && strike !== 'noStrike') options.strike = strike
  const baseline = numAttr(rPr, 'baseline')
  if (baseline && baseline > 0) options.superscript = true
  if (baseline && baseline < 0) options.subscript = true
  const spacing = numAttr(rPr, 'spc')
  if (spacing) options.charSpacing = spacing / 100
  const color = importer.fill(rPr)
  if (color && color !== 'none') {
    options.color = color.color
    if (color.transparency !== undefined) options.transparency = color.transparency
  }
  const highlight = importer.color(child(rPr, 'a:highlight'))
  if (highlight) options.highlight = highlight.color
  const face = attr(child(rPr, 'a:latin'), 'typeface')
  if (face && !face.startsWith('+')) options.fontFace = face
  const link = attr(child(rPr, 'a:hlinkClick'), 'r:id')
  const target = link ? relationships.get(link) : undefined
  if (target?.external) options.hyperlink = { url: target.target }
  return options
}

/** Reads `a:p` paragraphs into PptxGenJS runs. */
const readParagraphs = async (importer: Importer, txBody: XmlNode | undefined, part: string): Promise<TextRun[]> => {
  const runs: TextRun[] = []
  const relationships = await importer.relationships(part)
  // The shape's own list style sets the defaults its runs inherit, as PptxGenJS writes for slide numbers.
  const defaults = runProps(importer, path(txBody, 'a:lstStyle', 'a:lvl1pPr', 'a:defRPr'), relationships)
  const paragraphs = all(txBody, 'a:p')
  paragraphs.forEach((p, index) => {
    const pPr = child(p, 'a:pPr')
    const paragraph: Options = {}
    const align = ALIGN[attr(pPr, 'algn') ?? '']
    if (align) paragraph.align = align
    const level = numAttr(pPr, 'lvl')
    if (level) paragraph.indentLevel = level
    if (child(pPr, 'a:buChar')) {
      const char = attr(child(pPr, 'a:buChar'), 'char')
      paragraph.bullet =
        char && char !== '•' ? { code: char.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0') } : true
    } else if (child(pPr, 'a:buAutoNum')) {
      const type = attr(child(pPr, 'a:buAutoNum'), 'type')
      const start = numAttr(child(pPr, 'a:buAutoNum'), 'startAt')
      paragraph.bullet = {
        type: 'number',
        ...(type ? { numberType: type } : {}),
        ...(start ? { numberStartAt: start } : {}),
      }
    }
    const before = numAttr(path(pPr, 'a:spcBef', 'a:spcPts'), 'val')
    if (before) paragraph.paraSpaceBefore = before / 100
    const after = numAttr(path(pPr, 'a:spcAft', 'a:spcPts'), 'val')
    if (after) paragraph.paraSpaceAfter = after / 100
    const lineSpacing = numAttr(path(pPr, 'a:lnSpc', 'a:spcPct'), 'val')
    if (lineSpacing && lineSpacing !== 100000) paragraph.lineSpacingMultiple = round(lineSpacing / 100000)
    const start = runs.length
    for (const item of childrenOf(p)) {
      const tag = tagOf(item)
      if (tag === 'a:br') {
        const last = runs[runs.length - 1]
        if (last && runs.length > start) last.options = { ...last.options, breakLine: true }
        else runs.push({ text: '', options: { ...paragraph, breakLine: true } })
        continue
      }
      if (tag !== 'a:r' && tag !== 'a:fld') continue
      const options: Options = { ...defaults, ...paragraph, ...runProps(importer, child(item, 'a:rPr'), relationships) }
      if (tag === 'a:fld' && attr(item, 'type') === 'slidenum') options._field = 'slidenum'
      runs.push({ text: textOf(child(item, 'a:t')), options: options })
    }
    if (runs.length === start) runs.push({ text: '', options: { ...paragraph } })
    if (index < paragraphs.length - 1) {
      const last = runs[runs.length - 1]
      if (last) last.options = { ...last.options, breakLine: true }
    }
  })
  return runs.map((run) => (run.options && Object.keys(run.options).length > 0 ? run : { text: run.text }))
}

const bodyOptions = (txBody: XmlNode | undefined): Options => {
  const bodyPr = child(txBody, 'a:bodyPr')
  const out: Options = { valign: ANCHOR[attr(bodyPr, 'anchor') ?? 't'] ?? 'top' }
  const insets = ['tIns', 'rIns', 'bIns', 'lIns'].map((name) => numAttr(bodyPr, name))
  if (insets.some((value) => value !== undefined)) {
    const defaults = [45720, 91440, 45720, 91440]
    out.margin = insets.map((value, index) => round((value ?? defaults[index] ?? 0) / EMU_PER_PT))
  }
  if (attr(bodyPr, 'wrap') === 'none') out.wrap = false
  const vert = attr(bodyPr, 'vert')
  if (vert && vert !== 'horz') out.vert = vert
  if (child(bodyPr, 'a:normAutofit')) out.fit = 'shrink'
  if (child(bodyPr, 'a:spAutoFit')) out.fit = 'resize'
  return out
}

/** A text box holding nothing but the slide number field. */
const isSlideNumber = (object: SlideObject): boolean =>
  object.kind === 'text' &&
  object.runs.some((run) => (run.options as Options | undefined)?._field === 'slidenum') &&
  object.runs.every((run) => (run.options as Options | undefined)?._field === 'slidenum' || run.text.trim() === '')

const SLIDE_NUMBER_KEYS = ['fontFace', 'fontSize', 'color', 'bold', 'italic', 'align', 'valign', 'margin']

type SlideNumberProps = NonNullable<SlideModel['slideNumber']>

const slideNumberOf = (object: Extract<SlideObject, { kind: 'text' }>): SlideNumberProps => {
  const field = object.runs.find((run) => (run.options as Options | undefined)?._field === 'slidenum')
  const merged: Options = { ...(object.options as Options), ...(field?.options as Options | undefined) }
  const out: Options = { ...object.frame }
  for (const key of SLIDE_NUMBER_KEYS) if (merged[key] !== undefined) out[key] = merged[key]
  // PptxGenJS anchors a slide number to the top unless told otherwise.
  if (out.valign === 'top') delete out.valign
  return out
}

const hasText = (runs: TextRun[]): boolean => runs.some((run) => run.text.trim() !== '')

/** Finds the frame of a placeholder on the layout or master the slide inherits from. */
const inheritedFrame = async (
  importer: Importer,
  key: { type: string; idx: string },
  parts: string[],
): Promise<Frame | undefined> => {
  for (const part of parts) {
    const tree = path(await importer.xml(part), 'p:cSld', 'p:spTree')
    for (const sp of descendants(tree, 'p:sp')) {
      const other = placeholderKey(sp)
      if (!other) continue
      const matches =
        (key.idx !== '' && other.idx === key.idx) ||
        other.type === key.type ||
        (key.type === 'ctrTitle' && other.type === 'title')
      if (!matches) continue
      const frame = xfrmFrame(path(sp, 'p:spPr', 'a:xfrm'))
      if (frame) return frame
    }
  }
  return undefined
}

const shapeObject = async (
  importer: Importer,
  sp: XmlNode,
  part: string,
  transform: Transform,
  inheritFrom: string[],
  skipPlaceholders: boolean,
): Promise<SlideObject | undefined> => {
  const key = placeholderKey(sp)
  if (key && skipPlaceholders) return undefined
  const spPr = child(sp, 'p:spPr')
  const xfrm = child(spPr, 'a:xfrm')
  let frame = xfrmFrame(xfrm)
  if (!frame && key) frame = await inheritedFrame(importer, key, inheritFrom)
  if (!frame) return undefined
  frame = transform(frame)
  const geometry = geometryOptions(xfrm)
  const preset = attr(child(spPr, 'a:prstGeom'), 'prst')
  if (!preset && child(spPr, 'a:custGeom')) importer.warn('Freeform shapes are imported as rectangles.')
  const shape = preset ?? 'rect'
  const options: Options = { ...geometry }
  const fill = importer.fill(spPr)
  if (fill && fill !== 'none') options.fill = fill
  const line = importer.line(spPr)
  if (line) options.line = line
  if (shape === 'roundRect') {
    const adjust = attr(descendants(child(spPr, 'a:prstGeom'), 'a:gd')[0], 'fmla')
    const ratio = adjust ? Number(adjust.replace('val ', '')) / 100000 : 0.16667
    options.rectRadius = round(ratio * Math.min(frame.w, frame.h))
  }
  const txBody = child(sp, 'p:txBody')
  const runs = txBody ? await readParagraphs(importer, txBody, part) : []
  if (hasText(runs)) {
    const textOptions: Options = { ...options, ...bodyOptions(txBody) }
    if (shape !== 'rect') textOptions.shape = shape
    return { kind: 'text', frame, runs, options: textOptions }
  }
  if (key) return undefined
  if (!options.fill && !options.line) return undefined
  return { kind: 'shape', frame, shape, options: options }
}

const connectorObject = (importer: Importer, cxn: XmlNode, transform: Transform): SlideObject | undefined => {
  const spPr = child(cxn, 'p:spPr')
  const frame = xfrmFrame(child(spPr, 'a:xfrm'))
  if (!frame) return undefined
  const options: Options = {
    ...geometryOptions(child(spPr, 'a:xfrm')),
    line: importer.line(spPr) ?? { color: '000000', width: 1 },
  }
  return { kind: 'shape', frame: transform(frame), shape: 'line', options: options }
}

const pictureObject = async (
  importer: Importer,
  pic: XmlNode,
  part: string,
  transform: Transform,
): Promise<SlideObject | undefined> => {
  const spPr = child(pic, 'p:spPr')
  const frame = xfrmFrame(child(spPr, 'a:xfrm'))
  const embed = attr(path(pic, 'p:blipFill', 'a:blip'), 'r:embed')
  if (!frame || !embed) return undefined
  const target = (await importer.relationships(part)).get(embed)
  if (!target || target.external) return undefined
  const src = await importer.image(target.target)
  const options: Options = { ...geometryOptions(child(spPr, 'a:xfrm')) }
  const description = attr(path(pic, 'p:nvPicPr', 'p:cNvPr'), 'descr')
  if (description) options.altText = description
  if (attr(child(spPr, 'a:prstGeom'), 'prst') === 'ellipse') options.rounding = true
  if (
    descendants(path(pic, 'p:nvPicPr'), 'a:videoFile').length > 0 ||
    descendants(path(pic, 'p:nvPicPr'), 'p14:media').length > 0
  ) {
    importer.warn('Video and audio are imported as their poster picture.')
  }
  return { kind: 'image', frame: transform(frame), src, options: options }
}

const tableObject = async (importer: Importer, frame: Frame, tbl: XmlNode, part: string): Promise<SlideObject> => {
  const colW = all(child(tbl, 'a:tblGrid'), 'a:gridCol').map((col) => inches(numAttr(col, 'w')))
  const rowH: number[] = []
  const rows: TableCellModel[][] = []
  for (const tr of all(tbl, 'a:tr')) {
    rowH.push(inches(numAttr(tr, 'h')))
    const cells: TableCellModel[] = []
    for (const tc of all(tr, 'a:tc')) {
      if (attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1') continue
      const txBody = child(tc, 'a:txBody')
      let runs = await readParagraphs(importer, txBody, part)
      const options: Options = {}
      // One run's formatting belongs to the cell.
      if (runs.length === 1 && runs[0]?.options) {
        Object.assign(options, runs[0].options)
        runs = [{ text: runs[0].text }]
      }
      const plain = runs.every((run) => !run.options || Object.keys(run.options).every((key) => key === 'breakLine'))
      const colspan = numAttr(tc, 'gridSpan')
      if (colspan && colspan > 1) options.colspan = colspan
      const rowspan = numAttr(tc, 'rowSpan')
      if (rowspan && rowspan > 1) options.rowspan = rowspan
      const tcPr = child(tc, 'a:tcPr')
      const fill = importer.fill(tcPr)
      if (fill && fill !== 'none') options.fill = fill
      const anchor = ANCHOR[attr(tcPr, 'anchor') ?? '']
      if (anchor) options.valign = anchor
      const borders = ['a:lnT', 'a:lnR', 'a:lnB', 'a:lnL'].map((name) => {
        const ln = child(tcPr, name)
        if (!ln || child(ln, 'a:noFill')) return { type: 'none' }
        const color = importer.fill(ln)
        return {
          type: 'solid',
          ...(color && color !== 'none' ? { color: color.color } : {}),
          pt: round((numAttr(ln, 'w') ?? 12700) / EMU_PER_PT),
        }
      })
      if (borders.some((border) => border.type !== 'none')) options.border = borders
      cells.push({
        text: plain ? runs.map((run) => run.text + (run.options?.breakLine ? '\n' : '')).join('') : runs,
        ...(Object.keys(options).length > 0 ? { options: options } : {}),
      })
    }
    rows.push(cells)
  }
  return { kind: 'table', frame, rows, options: { colW, rowH } }
}

const CHART_TYPES: Record<string, ChartName> = {
  'c:barChart': 'bar',
  'c:bar3DChart': 'bar3D',
  'c:lineChart': 'line',
  'c:line3DChart': 'line',
  'c:areaChart': 'area',
  'c:area3DChart': 'area',
  'c:pieChart': 'pie',
  'c:pie3DChart': 'pie',
  'c:ofPieChart': 'pie',
  'c:doughnutChart': 'doughnut',
  'c:scatterChart': 'scatter',
  'c:bubbleChart': 'bubble',
  'c:radarChart': 'radar',
}

const cache = (node: XmlNode | undefined): string[] => {
  const multi = descendants(node, 'c:multiLvlStrCache')[0]
  const holder =
    multi ??
    descendants(node, 'c:strCache')[0] ??
    descendants(node, 'c:numCache')[0] ??
    descendants(node, 'c:strLit')[0] ??
    descendants(node, 'c:numLit')[0]
  // Multi-level categories: the first level holds the labels next to the axis.
  const points = all(multi ? child(multi, 'c:lvl') : holder, 'c:pt')
  const count = numAttr(child(holder, 'c:ptCount'), 'val') ?? points.length
  const out = new Array<string>(count).fill('')
  for (const point of points) out[numAttr(point, 'idx') ?? 0] = textOf(child(point, 'c:v'))
  return out
}

const chartObject = async (importer: Importer, frame: Frame, chartPart: string): Promise<SlideObject | undefined> => {
  const root = await importer.xml(chartPart)
  const chart = child(root, 'c:chart')
  const plotArea = child(chart, 'c:plotArea')
  const options: Options = {}
  const titleNode = child(chart, 'c:title')
  if (titleNode && attr(child(chart, 'c:autoTitleDeleted'), 'val') !== '1') {
    options.showTitle = true
    const text = descendants(titleNode, 'a:t')
      .map((t) => textOf(t))
      .join('')
    if (text) options.title = text
  }
  const legendNode = child(chart, 'c:legend')
  if (legendNode) {
    options.showLegend = true
    const pos = attr(child(legendNode, 'c:legendPos'), 'val')
    if (pos) options.legendPos = pos
  }
  const parts: ComboChart[] = []
  const colors: string[] = []
  const partColors = new Map<ComboChart, string[]>()
  // In a combo, a group drawn against the second value axis is on the secondary axes.
  const valueAxes = all(plotArea, 'c:valAx').map((axis) => attr(child(axis, 'c:axId'), 'val'))
  for (const group of childrenOf(plotArea)) {
    const type = CHART_TYPES[tagOf(group) ?? '']
    if (!type) continue
    const partOptions: Options = {}
    const start = colors.length
    const axes = all(group, 'c:axId').map((axis) => attr(axis, 'val'))
    if (
      valueAxes.length > 1 &&
      type !== 'scatter' &&
      type !== 'bubble' &&
      axes.includes(valueAxes[1]) &&
      !axes.includes(valueAxes[0])
    ) {
      partOptions.secondaryValAxis = true
      partOptions.secondaryCatAxis = true
    }
    if (type === 'bar' || type === 'bar3D') {
      const direction = attr(child(group, 'c:barDir'), 'val')
      if (direction) partOptions.barDir = direction
    }
    const grouping = attr(child(group, 'c:grouping'), 'val')
    if (grouping && grouping !== 'standard' && type !== 'line') partOptions.barGrouping = grouping
    if (type === 'doughnut') {
      const hole = numAttr(child(group, 'c:holeSize'), 'val')
      if (hole) partOptions.holeSize = hole
    }
    if (type === 'radar') {
      const style = attr(child(group, 'c:radarStyle'), 'val')
      if (style) partOptions.radarStyle = style
    }
    const labels = child(group, 'c:dLbls')
    if (attr(child(labels, 'c:showVal'), 'val') === '1') partOptions.showValue = true
    if (attr(child(labels, 'c:showPercent'), 'val') === '1') partOptions.showPercent = true
    const data: ChartSeries[] = []
    const series = all(group, 'c:ser')
    if (type === 'scatter' || type === 'bubble') {
      data.push({ name: 'X-Axis', values: cache(child(series[0], 'c:xVal')).map(Number) })
    }
    for (const ser of series) {
      const name = cache(child(ser, 'c:tx'))[0] ?? (textOf(path(ser, 'c:tx', 'c:v')) || undefined)
      const entry: ChartSeries = {}
      if (name) entry.name = name
      if (type === 'scatter' || type === 'bubble') {
        entry.values = cache(child(ser, 'c:yVal')).map(Number)
        if (type === 'bubble') entry.sizes = cache(child(ser, 'c:bubbleSize')).map(Number)
      } else {
        entry.labels = cache(child(ser, 'c:cat'))
        entry.values = cache(child(ser, 'c:val')).map(Number)
      }
      data.push(entry)
      if (type === 'pie' || type === 'doughnut') {
        for (const point of all(ser, 'c:dPt')) {
          const color = importer.fill(child(point, 'c:spPr'))
          if (color && color !== 'none') colors[numAttr(child(point, 'c:idx'), 'val') ?? colors.length] = color.color
        }
      } else {
        const color = importer.fill(child(ser, 'c:spPr')) ?? importer.fill(child(child(ser, 'c:spPr'), 'a:ln'))
        if (color && color !== 'none') colors.push(color.color)
      }
    }
    if (type === 'line' && all(group, 'c:ser').some((ser) => attr(child(ser, 'c:smooth'), 'val') === '1'))
      partOptions.lineSmooth = true
    const part: ComboChart = { type, data, options: partOptions }
    partColors.set(part, colors.slice(start).filter(Boolean))
    parts.push(part)
  }
  if (parts.length === 0) {
    importer.warn('A chart of a type PptxGenJS cannot write was skipped.')
    return undefined
  }
  const valAx = child(plotArea, 'c:valAx')
  const catAx = child(plotArea, 'c:catAx') ?? all(plotArea, 'c:valAx')[1]
  if (attr(child(valAx, 'c:delete'), 'val') === '1') options.valAxisHidden = true
  if (attr(child(catAx, 'c:delete'), 'val') === '1') options.catAxisHidden = true
  if (valAx && !child(valAx, 'c:majorGridlines')) options.valGridLine = { style: 'none' }
  const min = numAttr(path(valAx, 'c:scaling', 'c:min'), 'val')
  const max = numAttr(path(valAx, 'c:scaling', 'c:max'), 'val')
  if (min !== undefined) options.valAxisMinVal = min
  if (max !== undefined) options.valAxisMaxVal = max
  const format = attr(child(valAx, 'c:numFmt'), 'formatCode')
  if (format && format !== 'General') options.valAxisLabelFormatCode = format
  const filled = colors.filter(Boolean)
  if (filled.length > 0) options.chartColors = filled
  if (parts.length === 1) {
    const [only] = parts as [ComboChart]
    return { kind: 'chart', frame, type: only.type, data: only.data, options: { ...options, ...only.options } }
  }
  // Each part of a combo keeps its own colors.
  for (const part of parts) {
    const own = partColors.get(part) ?? []
    if (own.length > 0) part.options = { ...part.options, chartColors: own }
  }
  delete options.chartColors
  return { kind: 'chart', frame, type: parts, data: [], options: options }
}

const graphicFrameObject = async (
  importer: Importer,
  frameNode: XmlNode,
  part: string,
  transform: Transform,
): Promise<SlideObject | undefined> => {
  const frame = xfrmFrame(child(frameNode, 'p:xfrm'))
  if (!frame) return undefined
  const data = path(frameNode, 'a:graphic', 'a:graphicData')
  const uri = attr(data, 'uri') ?? ''
  if (uri.endsWith('/table')) return tableObject(importer, transform(frame), child(data, 'a:tbl') ?? {}, part)
  if (uri.endsWith('/chart')) {
    const id = attr(child(data, 'c:chart'), 'r:id')
    const target = id ? (await importer.relationships(part)).get(id) : undefined
    return target ? chartObject(importer, transform(frame), target.target) : undefined
  }
  if (uri.includes('diagram')) importer.warn('SmartArt is not imported.')
  else importer.warn(`An embedded object (${uri.split('/').pop() ?? 'unknown'}) is not imported.`)
  return undefined
}

const readTree = async (
  importer: Importer,
  tree: XmlNode | undefined,
  part: string,
  transform: Transform,
  inheritFrom: string[],
  skipPlaceholders: boolean,
): Promise<SlideObject[]> => {
  const out: SlideObject[] = []
  for (const node of childrenOf(tree)) {
    const tag = tagOf(node)
    let object: SlideObject | undefined
    if (tag === 'p:sp') object = await shapeObject(importer, node, part, transform, inheritFrom, skipPlaceholders)
    else if (tag === 'p:cxnSp') object = connectorObject(importer, node, transform)
    else if (tag === 'p:pic') object = await pictureObject(importer, node, part, transform)
    else if (tag === 'p:graphicFrame') object = await graphicFrameObject(importer, node, part, transform)
    else if (tag === 'p:grpSp') {
      const xfrm = path(node, 'p:grpSpPr', 'a:xfrm')
      const outer = xfrmFrame(xfrm)
      const chOff = child(xfrm, 'a:chOff')
      const chExt = child(xfrm, 'a:chExt')
      let inner = transform
      if (outer && chOff && chExt) {
        const cx = inches(numAttr(chOff, 'x'))
        const cy = inches(numAttr(chOff, 'y'))
        const sx = outer.w / (inches(numAttr(chExt, 'cx')) || 1)
        const sy = outer.h / (inches(numAttr(chExt, 'cy')) || 1)
        inner = (frame) =>
          transform({
            x: round(outer.x + (frame.x - cx) * sx),
            y: round(outer.y + (frame.y - cy) * sy),
            w: round(frame.w * sx),
            h: round(frame.h * sy),
          })
      }
      out.push(...(await readTree(importer, node, part, inner, inheritFrom, skipPlaceholders)))
    } else if (tag === 'mc:AlternateContent') {
      const choice = child(node, 'mc:Choice') ?? child(node, 'mc:Fallback')
      out.push(...(await readTree(importer, choice, part, transform, inheritFrom, skipPlaceholders)))
    }
    if (object) out.push(object)
  }
  return out
}

const backgroundOf = async (
  importer: Importer,
  root: XmlNode | undefined,
  part: string,
): Promise<Background | undefined> => {
  const bgPr = path(root, 'p:cSld', 'p:bg', 'p:bgPr')
  if (!bgPr) return undefined
  const embed = attr(path(bgPr, 'a:blipFill', 'a:blip'), 'r:embed')
  if (embed) {
    const target = (await importer.relationships(part)).get(embed)
    if (target && !target.external) return { src: await importer.image(target.target) }
  }
  const fill = importer.fill(bgPr)
  return fill && fill !== 'none' ? fill : undefined
}

const notesOf = async (importer: Importer, slidePart: string): Promise<string | undefined> => {
  const notesPart = await importer.relatedOfType(slidePart, 'notesSlide')
  if (!notesPart) return undefined
  const tree = path(await importer.xml(notesPart), 'p:cSld', 'p:spTree')
  for (const sp of all(tree, 'p:sp')) {
    if (placeholderKey(sp)?.type !== 'body') continue
    const text = all(child(sp, 'p:txBody'), 'a:p')
      .map((p) =>
        descendants(p, 'a:t')
          .map((t) => textOf(t))
          .join(''),
      )
      .join('\n')
      .trim()
    return text || undefined
  }
  return undefined
}

const layoutOf = (cx: number, cy: number): DeckModel['layout'] => {
  const width = round(cx / EMU_PER_IN)
  const height = round(cy / EMU_PER_IN)
  for (const layout of Object.values(LAYOUTS)) {
    if (Math.abs(layout.width - width) < 0.01 && Math.abs(layout.height - height) < 0.01) return layout
  }
  return { name: 'CUSTOM', width, height }
}

const toBytes = async (input: PptxInput): Promise<Uint8Array | ArrayBuffer> =>
  typeof Blob !== 'undefined' && input instanceof Blob ? input.arrayBuffer() : (input as Uint8Array | ArrayBuffer)

/** Reads a .pptx file into the model the writer and the preview use. */
export const parsePptx = async (input: PptxInput, options: ParseOptions = {}): Promise<ParseResult> => {
  const zip = await JSZip.loadAsync(await toBytes(input))
  const importer = new Importer(zip, {
    masterShapes: options.masterShapes ?? true,
    embedImages: options.embedImages ?? true,
  })
  const presentationPart = 'ppt/presentation.xml'
  const presentation = await importer.xml(presentationPart)
  if (!presentation) throw new Error('[vue-pptx-document] Not a PowerPoint file: ppt/presentation.xml is missing.')
  await importer.loadTheme(presentationPart)
  const size = child(presentation, 'p:sldSz')
  const layout = layoutOf(numAttr(size, 'cx') ?? 9144000, numAttr(size, 'cy') ?? 5143500)

  const sectionOf = new Map<string, string>()
  for (const section of descendants(presentation, 'p14:section')) {
    const name = attr(section, 'name')
    // PptxGenJS names the section it adds for slides without one Default-1, Default-2, ...
    if (!name || /^Default-\d+$/.test(name)) continue
    for (const id of descendants(section, 'p14:sldId')) sectionOf.set(attr(id, 'id') ?? '', name)
  }

  const core = await importer.xml('docProps/core.xml')
  const app = await importer.xml('docProps/app.xml')
  const meta: DeckModel['meta'] = {}
  const title = textOf(child(core, 'dc:title')).trim()
  const author = textOf(child(core, 'dc:creator')).trim()
  const subject = textOf(child(core, 'dc:subject')).trim()
  const company = textOf(child(app, 'Company')).trim()
  if (title) meta.title = title
  if (author) meta.author = author
  if (subject) meta.subject = subject
  if (company) meta.company = company

  const relationships = await importer.relationships(presentationPart)
  const slides: SlideModel[] = []
  for (const id of all(child(presentation, 'p:sldIdLst'), 'p:sldId')) {
    const target = relationships.get(attr(id, 'r:id') ?? '')
    if (!target) continue
    const part = target.target
    const root = await importer.xml(part)
    if (!root) continue
    const layoutPart = await importer.relatedOfType(part, 'slideLayout')
    const masterPart = layoutPart ? await importer.relatedOfType(layoutPart, 'slideMaster') : undefined
    const inheritFrom = [layoutPart, masterPart].filter((item): item is string => item !== undefined)
    const objects: SlideObject[] = []
    if (options.masterShapes !== false && attr(root, 'showMasterSp') !== '0') {
      for (const inherited of [masterPart, layoutPart]) {
        if (!inherited) continue
        const tree = path(await importer.xml(inherited), 'p:cSld', 'p:spTree')
        objects.push(...(await readTree(importer, tree, inherited, identity, [], true)))
      }
    }
    const tree = path(root, 'p:cSld', 'p:spTree')
    objects.push(...(await readTree(importer, tree, part, identity, inheritFrom, false)))
    let background = await backgroundOf(importer, root, part)
    for (const inherited of inheritFrom) {
      if (background) break
      background = await backgroundOf(importer, await importer.xml(inherited), inherited)
    }
    const slide: SlideModel = { objects: objects.filter((object) => !isSlideNumber(object)) }
    const number = objects.find(isSlideNumber)
    if (number?.kind === 'text') slide.slideNumber = slideNumberOf(number)
    if (background) slide.background = background
    if (attr(root, 'show') === '0') slide.hidden = true
    const section = sectionOf.get(attr(id, 'id') ?? '')
    if (section) slide.section = section
    const notes = await notesOf(importer, part)
    if (notes) slide.notes = notes
    slides.push(slide)
  }
  return { model: { layout, meta, masters: [], slides }, warnings: importer.warnings }
}
