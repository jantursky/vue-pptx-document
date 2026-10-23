import type { TextRun } from './model'
import { PT_PER_IN } from './units'

/** PowerPoint's default text size, used when neither the run nor the box sets one. */
export const DEFAULT_FONT_SIZE = 18
/** PowerPoint's default text box insets, in inches: left/right and top/bottom. */
export const DEFAULT_INSET_X = 0.1
export const DEFAULT_INSET_Y = 0.05
const LINE_HEIGHT = 1.2
/** PptxGenJS indents bulleted paragraphs by 27 points. */
export const BULLET_INDENT = 27 / PT_PER_IN

export interface Insets {
  top: number
  right: number
  bottom: number
  left: number
}

/** Text box insets in inches, from PptxGenJS's `margin` (points; arrays are left, right, bottom, top). */
export const textInsets = (margin: unknown): Insets => {
  const pt = (value: unknown): number => (typeof value === 'number' ? value / PT_PER_IN : 0)
  if (typeof margin === 'number') {
    const all = pt(margin)
    return { top: all, right: all, bottom: all, left: all }
  }
  if (Array.isArray(margin)) {
    return { left: pt(margin[0]), right: pt(margin[1]), bottom: pt(margin[2]), top: pt(margin[3]) }
  }
  return { top: DEFAULT_INSET_Y, right: DEFAULT_INSET_X, bottom: DEFAULT_INSET_Y, left: DEFAULT_INSET_X }
}

interface FontSpec {
  size: number
  bold: boolean
  italic: boolean
  face: string
}

type Measurer = (text: string, font: FontSpec) => number

let canvasMeasurer: Measurer | null | undefined

/** Text width in inches from a 2D canvas when there is one, so flex layouts match the preview. */
const fromCanvas = (): Measurer | null => {
  if (canvasMeasurer !== undefined) return canvasMeasurer
  canvasMeasurer = null
  try {
    const canvas =
      typeof OffscreenCanvas === 'function'
        ? new OffscreenCanvas(1, 1)
        : typeof document === 'undefined'
          ? null
          : document.createElement('canvas')
    const context = canvas?.getContext('2d') as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null
    if (context && typeof context.measureText === 'function') {
      canvasMeasurer = (text, font) => {
        context.font = `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${font.size}pt ${JSON.stringify(font.face)}, sans-serif`
        return context.measureText(text).width / 96
      }
    }
  } catch {
    canvasMeasurer = null
  }
  return canvasMeasurer
}

/** Average glyph widths in em, for Node and anywhere without a canvas. */
const estimate: Measurer = (text, font) => {
  let em = 0
  for (const char of text) {
    if (char === ' ') em += 0.28
    else if (/[il.,;:'!|]/.test(char)) em += 0.28
    else if (/[mwMW@]/.test(char)) em += 0.85
    else if (/[A-Z0-9]/.test(char)) em += 0.65
    else em += 0.52
  }
  if (font.bold) em *= 1.06
  return (em * font.size) / PT_PER_IN
}

const width: Measurer = (text, font) => (fromCanvas() ?? estimate)(text, font)

export interface TextBoxOptions {
  fontSize?: number
  fontFace?: string
  bold?: boolean
  italic?: boolean
  margin?: unknown
  lineSpacing?: number
  lineSpacingMultiple?: number
  paraSpaceBefore?: number
  paraSpaceAfter?: number
  wrap?: boolean
}

/** Splits runs into paragraphs at `breakLine` and at newlines inside the text. */
export const paragraphs = (runs: TextRun[]): TextRun[][] => {
  const out: TextRun[][] = [[]]
  for (const run of runs) {
    const parts = run.text.split('\n')
    parts.forEach((part, index) => {
      if (index > 0) out.push([])
      out[out.length - 1]?.push({ text: part, options: run.options })
    })
    if (run.options?.breakLine) out.push([])
  }
  if (out.length > 1 && out[out.length - 1]?.length === 0) out.pop()
  return out
}

/**
 * Estimates the size a text box needs, in inches. `maxWidth` is the box width the layout
 * offers; without it the text is measured on one line per paragraph.
 */
export const measureText = (
  runs: TextRun[],
  box: TextBoxOptions,
  maxWidth?: number,
): { width: number; height: number } => {
  const insets = textInsets(box.margin)
  const bulleted = runs.some((run) => run.options?.bullet)
  const available =
    maxWidth === undefined
      ? Infinity
      : Math.max(0, maxWidth - insets.left - insets.right - (bulleted ? BULLET_INDENT : 0))
  const wrap = box.wrap !== false
  let widest = 0
  let height = 0
  for (const paragraph of paragraphs(runs)) {
    let lineWidth = 0
    // The line's width without its trailing space, which doesn't take room at a wrap.
    let visible = 0
    let lineSize = 0
    let lines = 0
    let paragraphHeight = 0
    const first = paragraph[0]?.options
    const spacing = (size: number): number =>
      box.lineSpacing
        ? box.lineSpacing / PT_PER_IN
        : (size / PT_PER_IN) * LINE_HEIGHT * (box.lineSpacingMultiple ?? first?.lineSpacingMultiple ?? 1)
    const endLine = (): void => {
      const size = lineSize || (first?.fontSize ?? box.fontSize ?? DEFAULT_FONT_SIZE)
      paragraphHeight += spacing(size)
      widest = Math.max(widest, visible)
      lineWidth = 0
      visible = 0
      lineSize = 0
      lines++
    }
    for (const run of paragraph) {
      const font: FontSpec = {
        size: run.options?.fontSize ?? box.fontSize ?? DEFAULT_FONT_SIZE,
        bold: run.options?.bold ?? box.bold ?? false,
        italic: run.options?.italic ?? box.italic ?? false,
        face: run.options?.fontFace ?? box.fontFace ?? 'Arial',
      }
      for (const word of run.text.split(/(?<=\s)/)) {
        const trimmed = width(word.trimEnd(), font)
        if (wrap && lineWidth > 0 && lineWidth + trimmed > available) endLine()
        if (trimmed > 0) visible = lineWidth + trimmed
        lineWidth += width(word, font)
        lineSize = Math.max(lineSize, font.size)
      }
    }
    if (lineWidth > 0 || lines === 0) endLine()
    const before = (first?.paraSpaceBefore ?? box.paraSpaceBefore ?? 0) / PT_PER_IN
    const after = (first?.paraSpaceAfter ?? box.paraSpaceAfter ?? 0) / PT_PER_IN
    height += paragraphHeight + before + after
  }
  return {
    width: widest + insets.left + insets.right + (bulleted ? BULLET_INDENT : 0),
    height: height + insets.top + insets.bottom,
  }
}
