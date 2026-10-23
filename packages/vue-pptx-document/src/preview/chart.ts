import { h, type VNode } from 'vue'

import { cssColor } from '../core/colors'
import type { ChartName, ChartSeries, SlideObject } from '../core/model'
import { PT_PER_IN } from '../core/units'

type Options = Record<string, unknown>
type ChartObject = Extract<SlideObject, { kind: 'chart' }>

/** PptxGenJS's default series colors. */
const BAR_COLORS = ['C0504D', '4F81BD', '9BBB59', '8064A2', '4BACC6', 'F79646', '628FC6', 'C86360']
const PIE_COLORS = ['5DA5DA', 'FAA43A', '60BD68', 'F17CB0', 'B2912F', 'B276B2', 'DECF3F', 'F15854', 'A7A7A7']

const AXIS_FONT = 12
const TITLE_FONT = 18

interface Box {
  x: number
  y: number
  w: number
  h: number
}

interface Part {
  type: ChartName
  data: ChartSeries[]
  options: Options
}

const num = (value: unknown, fallback: number): number =>
  typeof value === 'number' && Number.isFinite(value) ? value : fallback
const r = (value: number): number => Math.round(value * 100) / 100

/** Formats a number with the subset of Excel format codes charts usually use. */
export const formatNumber = (value: number, code?: string): string => {
  if (!Number.isFinite(value)) return ''
  if (!code || code === 'General') return String(Math.round(value * 1000) / 1000)
  const percent = code.includes('%')
  const decimals = /\.(0+)/.exec(code)?.[1]?.length ?? 0
  const grouped = code.includes(',')
  const amount = percent ? value * 100 : value
  let text = Math.abs(amount).toFixed(decimals)
  if (grouped) {
    const [whole = '', fraction] = text.split('.')
    text = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (fraction ? `.${fraction}` : '')
  }
  const prefix = /^\s*"?([$€£¥])/.exec(code)?.[1] ?? ''
  const suffix = /([$€£¥])"?\s*$/.exec(code.replace('%', ''))?.[1] ?? ''
  return `${amount < 0 ? '-' : ''}${prefix}${text}${suffix}${percent ? '%' : ''}`
}

/** Round axis bounds and a step, as Office picks them. */
export const niceScale = (min: number, max: number, unit?: number): { min: number; max: number; step: number } => {
  if (min === max) {
    if (max === 0) return { min: 0, max: 1, step: 0.2 }
    if (min > 0) min = 0
    else max = 0
  }
  const span = max - min
  const raw = span / 5
  const magnitude = 10 ** Math.floor(Math.log10(raw))
  const step = unit ?? [1, 2, 2.5, 5, 10].map((m) => m * magnitude).find((m) => m >= raw) ?? magnitude * 10
  return { min: Math.floor(min / step) * step, max: Math.ceil(max / step - 1e-9) * step, step }
}

const textWidth = (text: string, size: number): number => text.length * size * 0.55

const color = (options: Options, index: number, pie: boolean): string => {
  const colors =
    Array.isArray(options.chartColors) && options.chartColors.length > 0
      ? (options.chartColors as string[])
      : pie
        ? PIE_COLORS
        : BAR_COLORS
  return cssColor(colors[index % colors.length], '#4F81BD')
}

const fillOpacity = (options: Options): number =>
  typeof options.chartColorsOpacity === 'number' ? options.chartColorsOpacity / 100 : 1

const label = (
  x: number,
  y: number,
  text: string,
  size: number,
  options: {
    anchor?: 'start' | 'middle' | 'end'
    color?: string
    bold?: boolean
    rotate?: number
    baseline?: string
    font?: string
  } = {},
): VNode =>
  h(
    'text',
    {
      x: r(x),
      y: r(y),
      'font-size': size,
      'font-weight': options.bold ? 700 : 400,
      'text-anchor': options.anchor ?? 'middle',
      'dominant-baseline': options.baseline ?? 'middle',
      fill: options.color ?? '#404040',
      transform: options.rotate ? `rotate(${options.rotate} ${r(x)} ${r(y)})` : undefined,
    },
    text,
  )

interface LegendEntry {
  name: string
  color: string
  line?: boolean
}

/** Draws the legend and returns the plot box that's left. */
const legend = (entries: LegendEntry[], options: Options, box: Box, out: VNode[]): Box => {
  if (!options.showLegend || entries.length === 0) return box
  const size = num(options.legendFontSize, AXIS_FONT)
  const textColor = cssColor(options.legendColor as string | undefined, '#404040')
  const swatch = size * 0.7
  const pos = (options.legendPos as string | undefined) ?? 'r'
  const itemWidth = (entry: LegendEntry): number => swatch + size * 0.4 + textWidth(entry.name, size) + size
  const draw = (entry: LegendEntry, x: number, y: number): void => {
    out.push(
      entry.line
        ? h('line', { x1: r(x), y1: r(y), x2: r(x + swatch), y2: r(y), stroke: entry.color, 'stroke-width': 2 })
        : h('rect', { x: r(x), y: r(y - swatch / 2), width: swatch, height: swatch, fill: entry.color }),
      label(x + swatch + size * 0.4, y, entry.name, size, { anchor: 'start', color: textColor }),
    )
  }
  if (pos === 'r' || pos === 'l') {
    const width = Math.min(box.w * 0.4, Math.max(...entries.map(itemWidth)))
    const lineHeight = size * 1.5
    const top = box.y + box.h / 2 - (entries.length * lineHeight) / 2 + lineHeight / 2
    const x = pos === 'r' ? box.x + box.w - width : box.x
    entries.forEach((entry, index) => draw(entry, x, top + index * lineHeight))
    return pos === 'r'
      ? { ...box, w: box.w - width - size }
      : { ...box, x: box.x + width + size, w: box.w - width - size }
  }
  const total = entries.reduce((sum, entry) => sum + itemWidth(entry), 0)
  const height = size * 1.8
  let x = pos === 'tr' ? box.x + box.w - total : box.x + (box.w - total) / 2
  const y = pos === 'b' ? box.y + box.h - height / 2 : box.y + height / 2
  for (const entry of entries) {
    draw(entry, x, y)
    x += itemWidth(entry)
  }
  return pos === 'b' ? { ...box, h: box.h - height } : { ...box, y: box.y + height, h: box.h - height }
}

const title = (options: Options, box: Box, out: VNode[]): Box => {
  if (!options.showTitle) return box
  const text = typeof options.title === 'string' ? options.title : 'Chart Title'
  const size = num(options.titleFontSize, TITLE_FONT)
  out.push(
    label(box.x + box.w / 2, box.y + size * 0.8, text, size, {
      color: cssColor(options.titleColor as string | undefined, '#404040'),
      bold: options.titleBold !== false,
    }),
  )
  return { ...box, y: box.y + size * 1.6, h: box.h - size * 1.6 }
}

// Pie and doughnut ----------------------------------------------------------------------------

const arc = (cx: number, cy: number, outer: number, inner: number, start: number, end: number): string => {
  const point = (radius: number, angle: number): string =>
    `${r(cx + radius * Math.sin(angle))} ${r(cy - radius * Math.cos(angle))}`
  const large = end - start > Math.PI ? 1 : 0
  if (end - start >= Math.PI * 2 - 1e-6) {
    const half = start + Math.PI
    return inner > 0
      ? `M${point(outer, start)} A${outer} ${outer} 0 1 1 ${point(outer, half)} A${outer} ${outer} 0 1 1 ${point(outer, start)} M${point(inner, start)} A${inner} ${inner} 0 1 0 ${point(inner, half)} A${inner} ${inner} 0 1 0 ${point(inner, start)} Z`
      : `M${point(outer, start)} A${outer} ${outer} 0 1 1 ${point(outer, half)} A${outer} ${outer} 0 1 1 ${point(outer, start)} Z`
  }
  if (inner <= 0) {
    return `M${r(cx)} ${r(cy)} L${point(outer, start)} A${outer} ${outer} 0 ${large} 1 ${point(outer, end)} Z`
  }
  return `M${point(outer, start)} A${outer} ${outer} 0 ${large} 1 ${point(outer, end)} L${point(inner, end)} A${inner} ${inner} 0 ${large} 0 ${point(inner, start)} Z`
}

const renderPie = (part: Part, options: Options, box: Box, out: VNode[]): void => {
  const doughnut = part.type === 'doughnut'
  const labels = part.data[0]?.labels ?? []
  const plot = legend(
    labels.map((name, index) => ({ name, color: color(options, index, true) })),
    options,
    box,
    out,
  )
  const cx = plot.x + plot.w / 2
  const cy = plot.y + plot.h / 2
  const outer = Math.max(0, Math.min(plot.w, plot.h) / 2 - 4)
  const hole = doughnut ? (outer * num(options.holeSize, 50)) / 100 : 0
  const series = doughnut ? part.data : part.data.slice(0, 1)
  const ring = series.length > 0 ? (outer - hole) / series.length : 0
  const start = (num(options.firstSliceAng, 0) * Math.PI) / 180
  const labelSize = num(options.dataLabelFontSize, AXIS_FONT)
  const labelColor = cssColor(options.dataLabelColor as string | undefined, '#404040')
  series.forEach((entry, seriesIndex) => {
    const values = (entry.values ?? []).map((value) => Math.max(0, num(value, 0)))
    const total = values.reduce((sum, value) => sum + value, 0) || 1
    const ringOuter = outer - seriesIndex * ring
    const ringInner = doughnut ? ringOuter - ring : 0
    let angle = start
    values.forEach((value, index) => {
      const sweep = (value / total) * Math.PI * 2
      out.push(
        h('path', {
          d: arc(cx, cy, ringOuter, ringInner, angle, angle + sweep),
          fill: color(options, index, true),
          'fill-opacity': fillOpacity(options),
          stroke: '#FFFFFF',
          'stroke-width': 1,
        }),
      )
      const parts: string[] = []
      if (options.showLabel) parts.push(labels[index] ?? '')
      if (options.showValue) parts.push(formatNumber(value, options.dataLabelFormatCode as string | undefined))
      if (options.showPercent) parts.push(`${Math.round((value / total) * 100)}%`)
      if (parts.length > 0 && sweep > 0) {
        const mid = angle + sweep / 2
        const radius = doughnut
          ? (ringOuter + ringInner) / 2
          : ringOuter * (options.dataLabelPosition === 'outEnd' ? 1.12 : 0.65)
        out.push(
          label(cx + radius * Math.sin(mid), cy - radius * Math.cos(mid), parts.join(' '), labelSize, {
            color: labelColor,
          }),
        )
      }
      angle += sweep
    })
  })
}

// Radar ---------------------------------------------------------------------------------------

const renderRadar = (part: Part, options: Options, box: Box, out: VNode[]): void => {
  const labels = part.data[0]?.labels ?? []
  const plot = legend(
    part.data.map((entry, index) => ({
      name: entry.name ?? `Series ${index + 1}`,
      color: color(options, index, false),
      line: true,
    })),
    options,
    box,
    out,
  )
  const size = num(options.catAxisLabelFontSize, AXIS_FONT)
  const cx = plot.x + plot.w / 2
  const cy = plot.y + plot.h / 2
  const radius = Math.max(0, Math.min(plot.w, plot.h) / 2 - size * 2)
  const all = part.data.flatMap((entry) => entry.values ?? [])
  const scale = niceScale(Math.min(0, ...all), Math.max(0, ...all), options.valAxisMajorUnit as number | undefined)
  const n = Math.max(labels.length, 1)
  const angleOf = (index: number): number => (index / n) * Math.PI * 2
  const at = (value: number, index: number): [number, number] => {
    const ratio = (value - scale.min) / (scale.max - scale.min || 1)
    return [cx + radius * ratio * Math.sin(angleOf(index)), cy - radius * ratio * Math.cos(angleOf(index))]
  }
  for (let tick = scale.min + scale.step; tick <= scale.max + 1e-9; tick += scale.step) {
    const ring = labels.map((_, index) => at(tick, index))
    out.push(
      h('polygon', {
        points: ring.map(([x, y]) => `${r(x)},${r(y)}`).join(' '),
        fill: 'none',
        stroke: '#D9D9D9',
        'stroke-width': 0.75,
      }),
    )
  }
  labels.forEach((name, index) => {
    const [x, y] = at(scale.max, index)
    out.push(h('line', { x1: r(cx), y1: r(cy), x2: r(x), y2: r(y), stroke: '#D9D9D9', 'stroke-width': 0.75 }))
    const [lx, ly] = [cx + (radius + size) * Math.sin(angleOf(index)), cy - (radius + size) * Math.cos(angleOf(index))]
    out.push(label(lx, ly, name, size))
  })
  const style = (options.radarStyle as string | undefined) ?? 'standard'
  part.data.forEach((entry, seriesIndex) => {
    const points = (entry.values ?? []).map((value, index) => at(num(value, 0), index))
    const stroke = color(options, seriesIndex, false)
    out.push(
      h('polygon', {
        points: points.map(([x, y]) => `${r(x)},${r(y)}`).join(' '),
        fill: style === 'filled' ? stroke : 'none',
        'fill-opacity': style === 'filled' ? 0.55 * fillOpacity(options) : undefined,
        stroke,
        'stroke-width': num(options.lineSize, 2),
      }),
    )
    if (style === 'marker') {
      for (const [x, y] of points) out.push(h('circle', { cx: r(x), cy: r(y), r: 3, fill: stroke }))
    }
  })
}

// Bar, line, area, scatter, bubble and combos -------------------------------------------------

const isBar = (type: ChartName): boolean => type === 'bar' || type === 'bar3D'

const smoothPath = (points: [number, number][]): string => {
  if (points.length < 3) return `M${points.map(([x, y]) => `${r(x)} ${r(y)}`).join(' L')}`
  let d = `M${r(points[0]?.[0] ?? 0)} ${r(points[0]?.[1] ?? 0)}`
  points.slice(0, -1).forEach((p1, i) => {
    const p0 = points[i - 1] ?? p1
    const p2 = points[i + 1] ?? p1
    const p3 = points[i + 2] ?? p2
    const c1x = p1[0] + (p2[0] - p0[0]) / 6
    const c1y = p1[1] + (p2[1] - p0[1]) / 6
    const c2x = p2[0] - (p3[0] - p1[0]) / 6
    const c2y = p2[1] - (p3[1] - p1[1]) / 6
    d += ` C${r(c1x)} ${r(c1y)} ${r(c2x)} ${r(c2y)} ${r(p2[0])} ${r(p2[1])}`
  })
  return d
}

const marker = (symbol: string, x: number, y: number, size: number, fill: string): VNode | null => {
  const s = size / 2
  switch (symbol) {
    case 'none':
      return null
    case 'square':
      return h('rect', { x: r(x - s), y: r(y - s), width: size, height: size, fill })
    case 'diamond':
      return h('polygon', {
        points: `${r(x)},${r(y - s)} ${r(x + s)},${r(y)} ${r(x)},${r(y + s)} ${r(x - s)},${r(y)}`,
        fill,
      })
    case 'triangle':
      return h('polygon', { points: `${r(x)},${r(y - s)} ${r(x + s)},${r(y + s)} ${r(x - s)},${r(y + s)}`, fill })
    case 'dash':
      return h('rect', { x: r(x - s), y: r(y - s / 4), width: size, height: s / 2, fill })
    case 'dot':
      return h('circle', { cx: r(x), cy: r(y), r: s / 2, fill })
    default:
      return h('circle', { cx: r(x), cy: r(y), r: s, fill })
  }
}

/** Values per series after stacking: `[base, top]` for each category. */
const stack = (series: ChartSeries[], grouping: string | undefined, categories: number): [number, number][][] => {
  const stacked = grouping === 'stacked' || grouping === 'percentStacked'
  if (!stacked)
    return series.map((entry) => Array.from({ length: categories }, (_, i) => [0, num(entry.values?.[i], 0)]))
  const totals = Array.from({ length: categories }, (_, i) =>
    series.reduce((sum, entry) => sum + Math.abs(num(entry.values?.[i], 0)), 0),
  )
  const positive = new Array<number>(categories).fill(0)
  const negative = new Array<number>(categories).fill(0)
  return series.map((entry) =>
    Array.from({ length: categories }, (_, i) => {
      let value = num(entry.values?.[i], 0)
      if (grouping === 'percentStacked') value = totals[i] ? value / (totals[i] ?? 1) : 0
      const base = (value >= 0 ? positive[i] : negative[i]) ?? 0
      if (value >= 0) positive[i] = base + value
      else negative[i] = base + value
      return [base, base + value]
    }),
  )
}

const renderCartesian = (parts: Part[], base: Options, box: Box, out: VNode[]): void => {
  // Combo charts set each axis on its own: valAxes[0] and catAxes[0] are the primary axes.
  const valAxes = Array.isArray(base.valAxes) ? (base.valAxes as Options[]) : []
  const catAxes = Array.isArray(base.catAxes) ? (base.catAxes as Options[]) : []
  const options: Options = { ...base, ...valAxes[0], ...catAxes[0] }
  const secondaryOptions: Options = { ...base, ...valAxes[1] }
  const first = parts[0]
  if (!first) return
  const xy = first.type === 'scatter' || first.type === 'bubble' || first.type === 'bubble3D'
  const horizontal = parts.every((part) => isBar(part.type)) && options.barDir === 'bar'
  const entries: LegendEntry[] = []
  let seriesIndex = 0
  const colorIndex = new Map<ChartSeries, number>()
  for (const part of parts) {
    const series = xy ? part.data.slice(1) : part.data
    for (const entry of series) {
      colorIndex.set(entry, seriesIndex)
      entries.push({
        name: entry.name ?? `Series ${seriesIndex + 1}`,
        color: color({ ...options, ...part.options }, seriesIndex, false),
        line: part.type === 'line' || part.type === 'scatter',
      })
      seriesIndex++
    }
  }
  const plotBox = legend(entries, options, box, out)
  const labels = xy ? (first.data[0]?.values ?? []).map(String) : (first.data[0]?.labels ?? []).map(String)
  const categories = Math.max(
    labels.length,
    ...parts.flatMap((part) => part.data.map((entry) => entry.values?.length ?? 0)),
  )

  // Value ranges, primary and secondary.
  const stacks = new Map<Part, [number, number][][]>()
  const range: Record<'primary' | 'secondary', [number, number]> = {
    primary: [Infinity, -Infinity],
    secondary: [Infinity, -Infinity],
  }
  for (const part of parts) {
    const grouping = (part.options.barGrouping ?? options.barGrouping) as string | undefined
    const series = xy ? part.data.slice(1) : part.data
    const stacked = stack(series, part.type === 'line' ? undefined : grouping, categories)
    stacks.set(part, stacked)
    const target = part.options.secondaryValAxis ? range.secondary : range.primary
    for (const values of stacked) {
      for (const [base, top] of values) {
        target[0] = Math.min(target[0], base, top)
        target[1] = Math.max(target[1], base, top)
      }
    }
    if (isBar(part.type) || part.type === 'area') {
      target[0] = Math.min(target[0], 0)
      target[1] = Math.max(target[1], 0)
    }
  }
  const percent = parts.some((part) => (part.options.barGrouping ?? options.barGrouping) === 'percentStacked')
  const valueScale = (which: 'primary' | 'secondary') => {
    const [low, high] = range[which]
    if (!Number.isFinite(low) || !Number.isFinite(high)) return null
    const axis = which === 'secondary' ? secondaryOptions : options
    const scale = niceScale(
      num(axis.valAxisMinVal, low),
      num(axis.valAxisMaxVal, high),
      typeof axis.valAxisMajorUnit === 'number' ? axis.valAxisMajorUnit : undefined,
    )
    if (typeof axis.valAxisMinVal === 'number') scale.min = axis.valAxisMinVal
    if (typeof axis.valAxisMaxVal === 'number') scale.max = axis.valAxisMaxVal
    if (percent) return { min: Math.max(scale.min, -1), max: Math.min(scale.max, 1), step: 0.2 }
    return scale
  }
  const primary = valueScale('primary') ?? { min: 0, max: 1, step: 0.2 }
  const secondary = valueScale('secondary')

  const axisSize = num(options.valAxisLabelFontSize, AXIS_FONT)
  const catSize = num(options.catAxisLabelFontSize, AXIS_FONT)
  const valueCode = (options.valAxisLabelFormatCode as string | undefined) ?? (percent ? '0%' : undefined)
  const secondaryCode = (secondaryOptions.valAxisLabelFormatCode as string | undefined) ?? valueCode
  const tickLabels = (scale: { min: number; max: number; step: number }, code = valueCode): string[] => {
    const out: string[] = []
    for (let tick = scale.min; tick <= scale.max + scale.step / 1000; tick += scale.step)
      out.push(formatNumber(tick, code))
    return out
  }
  const valHidden = options.valAxisHidden === true
  const secondaryHidden = secondaryOptions.valAxisHidden === true
  const catHidden = options.catAxisHidden === true
  const valLabelWidth = valHidden
    ? 0
    : Math.max(...tickLabels(primary).map((text) => textWidth(text, axisSize))) + axisSize * 0.6
  const secondaryTitle =
    valAxes[1] && secondaryOptions.showValAxisTitle && typeof secondaryOptions.valAxisTitle === 'string'
      ? secondaryOptions.valAxisTitle
      : undefined
  const secondaryWidth =
    (secondary && !secondaryHidden
      ? Math.max(...tickLabels(secondary, secondaryCode).map((text) => textWidth(text, axisSize))) + axisSize * 0.6
      : 0) + (secondary && secondaryTitle ? num(secondaryOptions.valAxisTitleFontSize, AXIS_FONT) * 1.4 : 0)
  const catLabelWidth = Math.max(0, ...labels.map((text) => textWidth(text, catSize))) + catSize * 0.6
  const valTitle =
    options.showValAxisTitle && typeof options.valAxisTitle === 'string' ? options.valAxisTitle : undefined
  const catTitle =
    options.showCatAxisTitle && typeof options.catAxisTitle === 'string' ? options.catAxisTitle : undefined
  const titleSize = num(options.valAxisTitleFontSize, AXIS_FONT)

  // Space for the axes: values on the left (or bottom, for horizontal bars).
  const left =
    (horizontal ? (catHidden ? 0 : Math.min(catLabelWidth, plotBox.w * 0.35)) : valLabelWidth) +
    (horizontal ? (catTitle ? titleSize * 1.4 : 0) : valTitle ? titleSize * 1.4 : 0)
  const bottom =
    (horizontal ? (valHidden ? 0 : axisSize * 1.6) : catHidden ? 0 : catSize * 1.6) +
    (horizontal ? (valTitle ? titleSize * 1.4 : 0) : catTitle ? titleSize * 1.4 : 0)
  const plot: Box = {
    x: plotBox.x + left + 4,
    y: plotBox.y + 6,
    w: Math.max(1, plotBox.w - left - 10 - secondaryWidth),
    h: Math.max(1, plotBox.h - bottom - 10),
  }

  const valueLength = horizontal ? plot.w : plot.h
  const toValue = (value: number, scale = primary): number => {
    const ratio = (value - scale.min) / (scale.max - scale.min || 1)
    return horizontal ? plot.x + ratio * valueLength : plot.y + plot.h - ratio * valueLength
  }
  const band = (horizontal ? plot.h : plot.w) / Math.max(categories, 1)
  // Horizontal bars run bottom to top, the first category lowest, as in PowerPoint.
  const center = (index: number): number =>
    horizontal ? plot.y + band * (categories - index - 0.5) : plot.x + band * (index + 0.5)

  // X scale for scatter and bubble.
  const xValues = (first.data[0]?.values ?? []).map((value) => num(value, 0))
  const xScale = xy ? niceScale(Math.min(0, ...xValues), Math.max(...xValues, 1)) : null
  const toX = (value: number): number =>
    plot.x + ((value - (xScale?.min ?? 0)) / ((xScale?.max ?? 1) - (xScale?.min ?? 0) || 1)) * plot.w

  // Gridlines and value axis.
  const grid = options.valGridLine as { style?: string; color?: string; size?: number } | undefined
  const gridColor = cssColor(grid?.color ?? (xy ? 'D9D9D9' : '888888'))
  const valueColor = cssColor(options.valAxisLabelColor as string | undefined, '#595959')
  tickLabels(primary).forEach((text, index) => {
    const tick = primary.min + index * primary.step
    const at = toValue(tick)
    if (grid?.style !== 'none') {
      out.push(
        horizontal
          ? h('line', {
              x1: r(at),
              y1: plot.y,
              x2: r(at),
              y2: plot.y + plot.h,
              stroke: gridColor,
              'stroke-width': (grid?.size ?? 1) * 0.5,
            })
          : h('line', {
              x1: plot.x,
              y1: r(at),
              x2: plot.x + plot.w,
              y2: r(at),
              stroke: gridColor,
              'stroke-width': (grid?.size ?? 1) * 0.5,
            }),
      )
    }
    if (!valHidden) {
      out.push(
        horizontal
          ? label(at, plot.y + plot.h + axisSize * 0.9, text, axisSize, { color: valueColor })
          : label(plot.x - axisSize * 0.4, at, text, axisSize, { anchor: 'end', color: valueColor }),
      )
    }
  })
  if (secondary && !secondaryHidden) {
    tickLabels(secondary, secondaryCode).forEach((text, index) => {
      const at = toValue(secondary.min + index * secondary.step, secondary)
      out.push(label(plot.x + plot.w + axisSize * 0.4, at, text, axisSize, { anchor: 'start', color: valueColor }))
    })
  }

  // Category axis.
  const catColor = cssColor(options.catAxisLabelColor as string | undefined, '#595959')
  const rotate = typeof options.catAxisLabelRotate === 'number' ? options.catAxisLabelRotate : undefined
  if (xy && xScale) {
    for (let tick = xScale.min; tick <= xScale.max + 1e-9; tick += xScale.step) {
      if (!catHidden)
        out.push(
          label(
            toX(tick),
            plot.y + plot.h + catSize * 0.9,
            formatNumber(tick, options.catLabelFormatCode as string | undefined),
            catSize,
            { color: catColor },
          ),
        )
    }
  } else if (!catHidden) {
    labels.forEach((text, index) => {
      out.push(
        horizontal
          ? label(plot.x - catSize * 0.4, center(index), text, catSize, { anchor: 'end', color: catColor })
          : label(center(index), plot.y + plot.h + catSize * 0.9, text, catSize, {
              color: catColor,
              ...(rotate ? { rotate, anchor: rotate < 0 ? 'end' : 'start' } : {}),
            }),
      )
    })
  }
  const zero = toValue(Math.max(primary.min, Math.min(0, primary.max)))
  if (options.catAxisLineShow !== false) {
    out.push(
      horizontal
        ? h('line', {
            x1: r(zero),
            y1: plot.y,
            x2: r(zero),
            y2: plot.y + plot.h,
            stroke: '#BFBFBF',
            'stroke-width': 0.75,
          })
        : h('line', {
            x1: plot.x,
            y1: r(zero),
            x2: plot.x + plot.w,
            y2: r(zero),
            stroke: '#BFBFBF',
            'stroke-width': 0.75,
          }),
    )
  }
  if (valTitle) {
    out.push(
      horizontal
        ? label(plot.x + plot.w / 2, plotBox.y + plotBox.h - titleSize * 0.5, valTitle, titleSize, {
            color: valueColor,
          })
        : label(plotBox.x + titleSize * 0.6, plot.y + plot.h / 2, valTitle, titleSize, {
            color: valueColor,
            rotate: -90,
          }),
    )
  }
  if (secondary && secondaryTitle) {
    const size = num(secondaryOptions.valAxisTitleFontSize, AXIS_FONT)
    out.push(
      label(plotBox.x + plotBox.w - size * 0.6, plot.y + plot.h / 2, secondaryTitle, size, {
        color: valueColor,
        rotate: 90,
      }),
    )
  }
  if (catTitle) {
    out.push(
      horizontal
        ? label(plotBox.x + titleSize * 0.6, plot.y + plot.h / 2, catTitle, titleSize, { color: catColor, rotate: -90 })
        : label(plot.x + plot.w / 2, plotBox.y + plotBox.h - titleSize * 0.5, catTitle, titleSize, { color: catColor }),
    )
  }

  // Series.
  const barParts = parts.filter((part) => isBar(part.type))
  const barSlots = barParts.reduce((sum, part) => {
    const grouping = (part.options.barGrouping ?? options.barGrouping) as string | undefined
    return sum + (grouping === 'stacked' || grouping === 'percentStacked' ? 1 : part.data.length)
  }, 0)
  const gap =
    num(
      options.barGapWidthPct,
      barParts.some((part) => String(part.options.barGrouping ?? options.barGrouping).includes('tacked')) ? 50 : 150,
    ) / 100
  const barWidth = band / Math.max(barSlots + gap, 1)
  let slot = 0
  const labelSize = num(options.dataLabelFontSize, 10)
  const labelColor = cssColor(options.dataLabelColor as string | undefined, '#404040')
  const dataCode = (options.dataLabelFormatCode as string | undefined) ?? (percent ? '0%' : undefined)

  for (const part of parts) {
    const partOptions = { ...options, ...part.options }
    const scale = part.options.secondaryValAxis && secondary ? secondary : primary
    const grouping = partOptions.barGrouping as string | undefined
    const stacked = grouping === 'stacked' || grouping === 'percentStacked'
    const values = stacks.get(part) ?? []
    const series = xy ? part.data.slice(1) : part.data
    series.forEach((entry, index) => {
      const fill = color(partOptions, colorIndex.get(entry) ?? index, false)
      const showValue = partOptions.showValue === true
      if (isBar(part.type)) {
        const offset = (stacked ? slot : slot + index) * barWidth + (gap * barWidth) / 2
        values[index]?.forEach(([base, top], category) => {
          const start = center(category) - band / 2 + offset
          const a = toValue(base, scale)
          const b = toValue(top, scale)
          const rect = horizontal
            ? { x: r(Math.min(a, b)), y: r(start), width: r(Math.abs(b - a)), height: r(barWidth) }
            : { x: r(start), y: r(Math.min(a, b)), width: r(barWidth), height: r(Math.abs(b - a)) }
          if (part.type === 'bar3D') {
            const depth = barWidth * 0.3
            out.push(
              h('polygon', {
                points: horizontal
                  ? `${rect.x},${rect.y} ${r(rect.x + depth)},${r(rect.y - depth)} ${r(rect.x + rect.width + depth)},${r(rect.y - depth)} ${r(rect.x + rect.width)},${rect.y}`
                  : `${rect.x},${rect.y} ${r(rect.x + depth)},${r(rect.y - depth)} ${r(rect.x + rect.width + depth)},${r(rect.y - depth)} ${r(rect.x + rect.width)},${rect.y}`,
                fill,
                'fill-opacity': 0.75,
              }),
              h('polygon', {
                points: `${r(rect.x + rect.width)},${rect.y} ${r(rect.x + rect.width + depth)},${r(rect.y - depth)} ${r(rect.x + rect.width + depth)},${r(rect.y + rect.height - depth)} ${r(rect.x + rect.width)},${r(rect.y + rect.height)}`,
                fill,
                'fill-opacity': 0.55,
              }),
            )
          }
          out.push(h('rect', { ...rect, fill, 'fill-opacity': fillOpacity(partOptions) }))
          if (showValue) {
            const value = top - base
            const text = formatNumber(value, dataCode)
            out.push(
              horizontal
                ? label(Math.max(a, b) + labelSize * 0.3, start + barWidth / 2, text, labelSize, {
                    anchor: 'start',
                    color: labelColor,
                  })
                : label(start + barWidth / 2, Math.min(a, b) - labelSize * 0.6, text, labelSize, { color: labelColor }),
            )
          }
        })
      } else if (part.type === 'line' || part.type === 'area') {
        const points = (values[index] ?? []).map(([, top], category): [number, number] => [
          center(category),
          toValue(top, scale),
        ])
        if (part.type === 'area') {
          const bases = (values[index] ?? []).map(([base], category): [number, number] => [
            center(category),
            toValue(base, scale),
          ])
          const outline = [...points, ...bases.reverse()]
          out.push(
            h('polygon', {
              points: outline.map(([x, y]) => `${r(x)},${r(y)}`).join(' '),
              fill,
              'fill-opacity': (stacked ? 1 : 0.85) * fillOpacity(partOptions),
            }),
          )
        } else {
          const width = num(partOptions.lineSize, 2)
          if (width > 0) {
            out.push(
              h('path', {
                d: partOptions.lineSmooth
                  ? smoothPath(points)
                  : `M${points.map(([x, y]) => `${r(x)} ${r(y)}`).join(' L')}`,
                fill: 'none',
                stroke: fill,
                'stroke-width': width,
                'stroke-dasharray': partOptions.lineDash && partOptions.lineDash !== 'solid' ? '4 3' : undefined,
              }),
            )
          }
          const symbol = (partOptions.lineDataSymbol as string | undefined) ?? 'circle'
          const size = num(partOptions.lineDataSymbolSize, 6)
          for (const [x, y] of points) {
            const node = marker(symbol, x, y, size, fill)
            if (node) out.push(node)
          }
        }
        if (showValue) {
          ;(values[index] ?? []).forEach(([base, top], category) => {
            out.push(
              label(
                center(category),
                toValue(top, scale) - labelSize * 0.8,
                formatNumber(top - base, dataCode),
                labelSize,
                { color: labelColor },
              ),
            )
          })
        }
      } else {
        // Scatter and bubble: the first series holds the X values.
        const ys = (entry.values ?? []).map((value) => num(value, 0))
        const points = ys.map((y, i): [number, number] => [toX(xValues[i] ?? i), toValue(y, scale)])
        if (part.type === 'scatter') {
          const width = num(partOptions.lineSize, 2)
          if (width > 0 && points.length > 1) {
            out.push(
              h('path', {
                d: partOptions.lineSmooth
                  ? smoothPath(points)
                  : `M${points.map(([x, y]) => `${r(x)} ${r(y)}`).join(' L')}`,
                fill: 'none',
                stroke: fill,
                'stroke-width': width,
              }),
            )
          }
          const symbol = (partOptions.lineDataSymbol as string | undefined) ?? 'circle'
          for (const [x, y] of points) {
            const node = marker(symbol, x, y, num(partOptions.lineDataSymbolSize, 6), fill)
            if (node) out.push(node)
          }
        } else {
          const sizes = series.flatMap((item) => item.sizes ?? []).map((value) => Math.abs(num(value, 0)))
          const largest = Math.max(...sizes, 1)
          const maxRadius = Math.min(plot.w, plot.h) * 0.12
          points.forEach(([x, y], i) => {
            const radius = Math.sqrt(Math.abs(num(entry.sizes?.[i], 1)) / largest) * maxRadius
            out.push(
              h('circle', {
                cx: r(x),
                cy: r(y),
                r: r(radius),
                fill,
                'fill-opacity': (part.type === 'bubble3D' ? 0.9 : 0.75) * fillOpacity(partOptions),
                stroke: fill,
              }),
            )
          })
        }
        if (showValue) {
          points.forEach(([x, y], i) =>
            out.push(
              label(x, y - labelSize * 0.9, formatNumber(ys[i] ?? 0, dataCode), labelSize, { color: labelColor }),
            ),
          )
        }
      }
    })
    if (isBar(part.type)) slot += stacked ? 1 : series.length
  }
}

/** Draws a chart as SVG in a viewBox measured in points. */
export const renderChart = (object: ChartObject, fontFace?: string): VNode => {
  const width = Math.max(object.frame.w * PT_PER_IN, 1)
  const height = Math.max(object.frame.h * PT_PER_IN, 1)
  const options = object.options as Options
  const parts: Part[] = Array.isArray(object.type)
    ? object.type.map((part) => ({
        type: part.type,
        data: part.data,
        options: (part.options ?? {}) as Options,
      }))
    : [{ type: object.type, data: object.data, options: {} }]
  const out: VNode[] = []
  const fill = typeof options.fill === 'string' ? cssColor(options.fill) : undefined
  if (fill) out.push(h('rect', { width, height, fill }))
  let box: Box = { x: 6, y: 6, w: width - 12, h: height - 12 }
  box = title(options, box, out)
  const first = parts[0]
  if (first && parts.length === 1 && (first.type === 'pie' || first.type === 'doughnut'))
    renderPie(first, options, box, out)
  else if (first && parts.length === 1 && first.type === 'radar') renderRadar(first, options, box, out)
  else if (first) renderCartesian(parts, options, box, out)
  const name = Array.isArray(object.type) ? 'combo' : object.type
  return h(
    'svg',
    {
      viewBox: `0 0 ${r(width)} ${r(height)}`,
      width: '100%',
      height: '100%',
      role: 'img',
      'aria-label':
        (options.altText as string | undefined) ??
        (typeof options.title === 'string' ? options.title : `${name} chart`),
      'data-chart': name,
      'font-family': `${JSON.stringify(fontFace ?? 'Arial')}, Arial, sans-serif`,
      style: { display: 'block', overflow: 'visible' },
    },
    out,
  )
}
