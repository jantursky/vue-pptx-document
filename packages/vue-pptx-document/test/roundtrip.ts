import type { DeckModel, SlideObject, TextRun } from '../src/core/model'

const FRAME_TOLERANCE = 0.01

const runsText = (runs: TextRun[]): string => runs.map((run) => run.text).join('')

const textsOf = (object: SlideObject): string[] => {
  if (object.kind === 'text') return [runsText(object.runs).replace(/\s+/g, ' ').trim()]
  if (object.kind === 'table') {
    return object.rows.flatMap((row) =>
      row.map((cell) => (typeof cell.text === 'string' ? cell.text : runsText(cell.text ?? [])).trim()),
    )
  }
  if (object.kind === 'placeholder') return object.text ? [object.text] : []
  return []
}

/**
 * Compares the model a deck resolves to with the model read back from the .pptx file it wrote,
 * and returns what differs: slide count, per-slide texts, object kinds, and frames within
 * 0.01 inch. An empty list means the writer and the importer agree.
 *
 * Not compared, because the importer doesn't read them back: the option keys beyond the frame
 * (the gaps are listed in `guides/import.mdx`), slide titles, and objects without a frame.
 */
export const compareModels = (written: DeckModel, read: DeckModel): string[] => {
  const differences: string[] = []
  if (written.slides.length !== read.slides.length) {
    return [`slides: written ${written.slides.length}, read ${read.slides.length}`]
  }
  for (const [index, slide] of written.slides.entries()) {
    const label = `slide ${index + 1}`
    const back = read.slides[index]!
    const kinds = slide.objects.map((object) => object.kind).join(',')
    const backKinds = back.objects.map((object) => object.kind).join(',')
    if (kinds !== backKinds) {
      differences.push(`${label}: kinds written [${kinds}], read [${backKinds}]`)
      continue
    }
    for (const [at, object] of slide.objects.entries()) {
      const other = back.objects[at]!
      const where = `${label} object ${at + 1} (${object.kind})`
      const texts = textsOf(object).join('|')
      const otherTexts = textsOf(other).join('|')
      if (texts !== otherTexts) differences.push(`${where}: text written "${texts}", read "${otherTexts}"`)
      for (const key of ['x', 'y', 'w', 'h'] as const) {
        if (Math.abs(object.frame[key] - other.frame[key]) > FRAME_TOLERANCE) {
          differences.push(`${where}: ${key} written ${object.frame[key]}, read ${other.frame[key]}`)
        }
      }
    }
  }
  return differences
}
