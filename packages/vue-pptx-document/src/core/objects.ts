import type PptxGenJS from 'pptxgenjs'

import { normalizeColors } from './colors'
import { warn } from './diagnostics'
import type {
  Background,
  ChartName,
  ChartSeries,
  ComboChart,
  Frame,
  SlideObject,
  TableCellModel,
  TextRun,
} from './model'
import { TYPES } from './nodes'
import type { TreeNode } from './snapshot'
import { toInches } from './units'

/** Props the components read themselves; everything else goes to PptxGenJS as an option. */
const GEOMETRY = new Set(['x', 'y', 'w', 'h', 'style'])

type Options = Record<string, unknown>

const omit = (props: Options, keys: Iterable<string>): Options => {
  const skip = new Set(keys)
  return Object.fromEntries(Object.entries(props).filter(([key, value]) => !skip.has(key) && value !== undefined))
}

/** `fill="#f35815"` and `line="#000"` are shorthands for `{ color }`. */
const expandFillLine = (options: Options): Options => {
  const out = { ...options }
  if (typeof out.fill === 'string') out.fill = { color: out.fill }
  if (typeof out.line === 'string') out.line = { color: out.line }
  return out
}

/** `href` is a shorthand for `hyperlink: { url }`; `underline` may be `true`. */
const expandRunShorthands = (options: Options): Options => {
  const out = { ...options }
  if (typeof out.href === 'string') {
    out.hyperlink = { url: out.href, ...(out.hyperlink as object | undefined) }
    delete out.href
  }
  if (out.underline === true) out.underline = { style: 'sng' }
  if (out.underline === false) delete out.underline
  return out
}

/** `margin` takes points in CSS order (top, right, bottom, left); PptxGenJS wants left, right, bottom, top. */
const expandMargin = (options: Options): Options => {
  const { margin } = options
  if (!Array.isArray(margin)) return options
  const [top = 0, right = top, bottom = top, left = right] = margin as number[]
  return { ...options, margin: [left, right, bottom, top] }
}

export const textOptions = (props: Options): PptxGenJS.TextPropsOptions =>
  normalizeColors(expandMargin(expandFillLine(expandRunShorthands(omit(props, GEOMETRY)))))

const splitLines = (text: string, options: Options): TextRun[] =>
  text.split('\n').map((line, index, lines) => ({
    text: line,
    options: index < lines.length - 1 ? { ...options, breakLine: true } : { ...options },
  }))

/**
 * Flattens `<Paragraph>`, `<Run>` and `<Br>` children into PptxGenJS text runs. Paragraph
 * props ride on each of their runs, and the paragraph's last run ends with `breakLine`.
 */
export const collectRuns = (children: TreeNode['children'], inherited: Options = {}): TextRun[] => {
  const out: TextRun[] = []
  const breakLast = (): void => {
    const last = out[out.length - 1]
    if (last) last.options = { ...last.options, breakLine: true }
    else out.push({ text: '', options: { breakLine: true } })
  }
  for (const child of children) {
    if (typeof child === 'string') out.push(...splitLines(child, inherited))
    else if (child.type === TYPES.run) {
      out.push(...collectRuns(child.children, { ...inherited, ...expandRunShorthands(omit(child.props, GEOMETRY)) }))
    } else if (child.type === TYPES.paragraph) {
      const runs = collectRuns(child.children, { ...inherited, ...expandRunShorthands(omit(child.props, GEOMETRY)) })
      if (out.length > 0 && !out[out.length - 1]?.options?.breakLine) breakLast()
      out.push(...(runs.length > 0 ? runs : [{ text: '', options: { ...inherited } }]))
      breakLast()
    } else if (child.type === TYPES.br) breakLast()
  }
  return out
}

/** Trims the whitespace a template leaves at the start and end of each paragraph. */
export const trimRuns = (runs: TextRun[]): TextRun[] => {
  const out = runs.map((run) => ({ ...run }))
  const groups: TextRun[][] = [[]]
  for (const run of out) {
    groups[groups.length - 1]?.push(run)
    if (run.options?.breakLine) groups.push([])
  }
  for (const group of groups) {
    for (const run of group) {
      run.text = run.text.replace(/^\s+/, '')
      if (run.text !== '') break
    }
    for (const run of [...group].reverse()) {
      run.text = run.text.replace(/\s+$/, '')
      if (run.text !== '') break
    }
  }
  const last = out[out.length - 1]
  if (last?.options?.breakLine) {
    const { breakLine: _, ...rest } = last.options
    last.options = rest
  }
  return out
    .filter((run) => run.text !== '' || run.options?.breakLine)
    .map((run) => (run.options && Object.keys(run.options).length > 0 ? run : { text: run.text }))
}

export const runsOf = (node: TreeNode): TextRun[] => normalizeColors(trimRuns(collectRuns(node.children)))

export const backgroundOf = (value: unknown): Background | undefined => {
  if (value == null || value === false) return undefined
  if (typeof value === 'string') {
    if (/^(data:|https?:|\/|\.\.?\/|blob:)/.test(value) || /\.(png|jpe?g|gif|svg|webp)$/i.test(value)) {
      return { src: value }
    }
    return normalizeColors({ color: value })
  }
  return normalizeColors(value as Background)
}

const asNumbers = (value: unknown, base: number): number[] | number | undefined => {
  if (Array.isArray(value)) return value.map((item) => toInches(item, base) ?? 0)
  return toInches(value, base)
}

const cellOf = (value: unknown): TableCellModel => {
  if (value == null) return { text: '' }
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean')
    return { text: String(value) }
  const cell = value as TableCellModel
  return { text: cell.text ?? '', ...(cell.options ? { options: expandFillLine(cell.options as Options) } : {}) }
}

/** Table rows from the `rows` prop, or from `<TableRow>` and `<TableCell>` children. */
export const tableRows = (node: TreeNode): TableCellModel[][] => {
  const fromProp = node.props.rows
  if (Array.isArray(fromProp)) {
    return normalizeColors(fromProp.map((row: unknown) => (Array.isArray(row) ? row.map(cellOf) : [cellOf(row)])))
  }
  const rows: TableCellModel[][] = []
  for (const row of node.children) {
    if (typeof row === 'string' || row.type !== TYPES.row) continue
    const rowOptions = omit(row.props, GEOMETRY)
    const cells: TableCellModel[] = []
    for (const cell of row.children) {
      if (typeof cell === 'string' || cell.type !== TYPES.cell) continue
      const runs = trimRuns(collectRuns(cell.children))
      const plain = runs.every((run) => !run.options || Object.keys(run.options).length === 0)
      const options = expandFillLine({ ...rowOptions, ...omit(cell.props, GEOMETRY) })
      cells.push({
        text: plain ? runs.map((run) => run.text).join('') : runs,
        ...(Object.keys(options).length > 0 ? { options: options } : {}),
      })
    }
    rows.push(cells)
  }
  return normalizeColors(rows)
}

const CHART_NAMES: Record<string, ChartName> = {
  area: 'area',
  bar: 'bar',
  bar3d: 'bar3D',
  bubble: 'bubble',
  bubble3d: 'bubble3D',
  doughnut: 'doughnut',
  line: 'line',
  pie: 'pie',
  radar: 'radar',
  scatter: 'scatter',
}

export const chartName = (value: unknown): ChartName => {
  const name = CHART_NAMES[String(value).toLowerCase()]
  if (!name) throw new Error(`[vue-pptx-document] Unknown chart type "${String(value)}".`)
  return name
}

const imageData = (src: string): PptxGenJS.DataOrPathProps =>
  src.startsWith('data:') ? { data: src.slice(5) } : { path: src }

/** Builds the slide object for a drawable node once its frame is known. */
export const toObject = (node: TreeNode, frame: Frame): SlideObject | undefined => {
  const { props } = node
  switch (node.type) {
    case TYPES.text: {
      return { kind: 'text', frame, runs: runsOf(node), options: textOptions(props) }
    }
    case TYPES.shape: {
      const shape = typeof props.type === 'string' ? props.type : 'rect'
      const runs = runsOf(node)
      if (runs.length > 0) {
        return { kind: 'text', frame, runs, options: { ...textOptions(omit(props, ['type'])), shape: shape as never } }
      }
      return {
        kind: 'shape',
        frame,
        shape,
        options: normalizeColors(expandFillLine(omit(props, [...GEOMETRY, 'type']))),
      }
    }
    case TYPES.image: {
      const src = typeof props.src === 'string' ? props.src : ''
      if (!src) {
        // PptxGenJS would fail later on an empty path, without saying which object it was.
        warn('missing-src', '<Image> without a `src` is skipped.')
        return undefined
      }
      return {
        kind: 'image',
        frame,
        src,
        options: { ...(omit(props, [...GEOMETRY, 'src']) as PptxGenJS.ImageProps), ...imageData(src) },
      }
    }
    case TYPES.table: {
      const options = omit(props, [...GEOMETRY, 'rows'])
      if (options.colW !== undefined) options.colW = asNumbers(options.colW, frame.w)
      if (options.rowH !== undefined) options.rowH = asNumbers(options.rowH, frame.h)
      return {
        kind: 'table',
        frame,
        rows: tableRows(node),
        // Table margins are already top, right, bottom, left in PptxGenJS, unlike text boxes.
        options: normalizeColors(expandFillLine(options)),
      }
    }
    case TYPES.chart: {
      const combo = Array.isArray(props.type)
      const type = combo
        ? (props.type as ComboChart[]).map((part) => ({ ...part, type: chartName(part.type) }))
        : chartName(props.type ?? 'bar')
      return {
        kind: 'chart',
        frame,
        type: normalizeColors(type),
        data: (Array.isArray(props.data) ? props.data : []) as ChartSeries[],
        options: normalizeColors(omit(props, [...GEOMETRY, 'type', 'data'])),
      }
    }
    case TYPES.media: {
      const src = typeof props.src === 'string' ? props.src : undefined
      const type = (props.type as PptxGenJS.MediaType | undefined) ?? 'video'
      const source = type === 'online' ? { link: src } : src ? imageData(src) : {}
      return {
        kind: 'media',
        frame,
        options: { ...omit(props, [...GEOMETRY, 'src', 'type']), type, ...source },
      }
    }
    case TYPES.placeholder: {
      const name = typeof props.name === 'string' ? props.name : 'body'
      const runs = runsOf(node)
      return {
        kind: 'placeholder',
        frame,
        name,
        ...(runs.length > 0 ? { text: runs.map((run) => run.text).join('\n') } : {}),
        options: {
          ...(textOptions(omit(props, ['name'])) as Options),
          name,
          type: (props.type as PptxGenJS.PLACEHOLDER_TYPE | undefined) ?? 'body',
        },
      }
    }
    default:
      return undefined
  }
}

/** A `<View>` with `fill`, `line` or `rectRadius` draws a rectangle behind its children. */
export const viewBackground = (node: TreeNode, frame: Frame): SlideObject | undefined => {
  const { fill, line, rectRadius, shadow, rotate } = node.props
  if (fill === undefined && line === undefined) return undefined
  const options = normalizeColors(
    expandFillLine(omit({ fill, line, rectRadius, shadow, rotate }, [])),
  ) as PptxGenJS.ShapeProps
  return { kind: 'shape', frame, shape: rectRadius ? 'roundRect' : 'rect', options }
}

export const DRAWABLE = new Set<string>([
  TYPES.text,
  TYPES.shape,
  TYPES.image,
  TYPES.table,
  TYPES.chart,
  TYPES.media,
  TYPES.placeholder,
])
