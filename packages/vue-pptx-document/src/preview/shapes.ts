/**
 * SVG outlines for PptxGenJS's preset shapes, in a box of `w` by `h` (any unit). Shapes not
 * listed here are previewed as their bounding rectangle with a dashed outline.
 */
type PathFn = (w: number, h: number, r: number) => string

const poly = (points: [number, number][]): string =>
  `M${points.map(([x, y]) => `${round(x)} ${round(y)}`).join(' L')} Z`

const round = (value: number): number => Math.round(value * 10000) / 10000

const regular =
  (sides: number, rotation = -Math.PI / 2): PathFn =>
  (w, h) =>
    poly(
      Array.from({ length: sides }, (_, i) => {
        const angle = rotation + (i * 2 * Math.PI) / sides
        return [w / 2 + (w / 2) * Math.cos(angle), h / 2 + (h / 2) * Math.sin(angle)]
      }),
    )

const star =
  (points: number, inner: number): PathFn =>
  (w, h) =>
    poly(
      Array.from({ length: points * 2 }, (_, i) => {
        const angle = -Math.PI / 2 + (i * Math.PI) / points
        const radius = i % 2 === 0 ? 1 : inner
        return [w / 2 + (w / 2) * radius * Math.cos(angle), h / 2 + (h / 2) * radius * Math.sin(angle)]
      }),
    )

const ellipse: PathFn = (w, h) =>
  `M0 ${h / 2} A${w / 2} ${h / 2} 0 1 0 ${w} ${h / 2} A${w / 2} ${h / 2} 0 1 0 0 ${h / 2} Z`

const roundRect: PathFn = (w, h, r) => {
  const radius = Math.max(0, Math.min(r, w / 2, h / 2))
  if (radius === 0) return `M0 0 H${w} V${h} H0 Z`
  return [
    `M${radius} 0 H${w - radius}`,
    `A${radius} ${radius} 0 0 1 ${w} ${radius} V${h - radius}`,
    `A${radius} ${radius} 0 0 1 ${w - radius} ${h} H${radius}`,
    `A${radius} ${radius} 0 0 1 0 ${h - radius} V${radius}`,
    `A${radius} ${radius} 0 0 1 ${radius} 0 Z`,
  ].join(' ')
}

const arrow =
  (direction: 'right' | 'left' | 'up' | 'down'): PathFn =>
  (w, h) => {
    const horizontal = direction === 'right' || direction === 'left'
    const long = horizontal ? w : h
    const short = horizontal ? h : w
    const head = Math.min(long * 0.5, short * 0.5)
    const shaft = short * 0.25
    // Pointing right in (long, short) space, then mapped.
    const points: [number, number][] = [
      [0, shaft],
      [long - head, shaft],
      [long - head, 0],
      [long, short / 2],
      [long - head, short],
      [long - head, short - shaft],
      [0, short - shaft],
    ]
    const map = ([a, b]: [number, number]): [number, number] => {
      if (direction === 'right') return [a, b]
      if (direction === 'left') return [long - a, b]
      if (direction === 'down') return [b, a]
      return [b, long - a]
    }
    return poly(points.map(map))
  }

const doubleArrow =
  (horizontal: boolean): PathFn =>
  (w, h) => {
    const long = horizontal ? w : h
    const short = horizontal ? h : w
    const head = Math.min(long * 0.3, short * 0.5)
    const shaft = short * 0.25
    const points: [number, number][] = [
      [0, short / 2],
      [head, 0],
      [head, shaft],
      [long - head, shaft],
      [long - head, 0],
      [long, short / 2],
      [long - head, short],
      [long - head, short - shaft],
      [head, short - shaft],
      [head, short],
    ]
    return poly(points.map(([a, b]) => (horizontal ? [a, b] : [b, a])))
  }

const cross: PathFn = (w, h) => {
  const x = w * 0.35
  const y = h * 0.35
  return poly([
    [x, 0],
    [w - x, 0],
    [w - x, y],
    [w, y],
    [w, h - y],
    [w - x, h - y],
    [w - x, h],
    [x, h],
    [x, h - y],
    [0, h - y],
    [0, y],
    [x, y],
  ])
}

const SHAPES: Record<string, PathFn> = {
  rect: (w, h) => `M0 0 H${w} V${h} H0 Z`,
  roundRect,
  round2SameRect: roundRect,
  snip1Rect: (w, h) => {
    const s = Math.min(w, h) * 0.2
    return poly([
      [0, 0],
      [w - s, 0],
      [w, s],
      [w, h],
      [0, h],
    ])
  },
  ellipse,
  flowChartConnector: ellipse,
  triangle: (w, h) =>
    poly([
      [w / 2, 0],
      [w, h],
      [0, h],
    ]),
  rtTriangle: (w, h) =>
    poly([
      [0, 0],
      [w, h],
      [0, h],
    ]),
  diamond: (w, h) =>
    poly([
      [w / 2, 0],
      [w, h / 2],
      [w / 2, h],
      [0, h / 2],
    ]),
  flowChartDecision: (w, h) =>
    poly([
      [w / 2, 0],
      [w, h / 2],
      [w / 2, h],
      [0, h / 2],
    ]),
  parallelogram: (w, h) =>
    poly([
      [w * 0.25, 0],
      [w, 0],
      [w * 0.75, h],
      [0, h],
    ]),
  flowChartInputOutput: (w, h) =>
    poly([
      [w * 0.2, 0],
      [w, 0],
      [w * 0.8, h],
      [0, h],
    ]),
  trapezoid: (w, h) =>
    poly([
      [w * 0.25, 0],
      [w * 0.75, 0],
      [w, h],
      [0, h],
    ]),
  pentagon: regular(5),
  hexagon: (w, h) =>
    poly([
      [w * 0.25, 0],
      [w * 0.75, 0],
      [w, h / 2],
      [w * 0.75, h],
      [w * 0.25, h],
      [0, h / 2],
    ]),
  heptagon: regular(7),
  octagon: (w, h) => {
    const s = Math.min(w, h) * 0.29
    return poly([
      [s, 0],
      [w - s, 0],
      [w, s],
      [w, h - s],
      [w - s, h],
      [s, h],
      [0, h - s],
      [0, s],
    ])
  },
  decagon: regular(10),
  dodecagon: regular(12),
  star4: star(4, 0.38),
  star5: star(5, 0.38),
  star6: star(6, 0.5),
  star7: star(7, 0.5),
  star8: star(8, 0.6),
  star10: star(10, 0.6),
  star12: star(12, 0.65),
  rightArrow: arrow('right'),
  leftArrow: arrow('left'),
  upArrow: arrow('up'),
  downArrow: arrow('down'),
  leftRightArrow: doubleArrow(true),
  upDownArrow: doubleArrow(false),
  chevron: (w, h) => {
    const s = Math.min(w * 0.5, h * 0.5)
    return poly([
      [0, 0],
      [w - s, 0],
      [w, h / 2],
      [w - s, h],
      [0, h],
      [s, h / 2],
    ])
  },
  homePlate: (w, h) => {
    const s = Math.min(w * 0.5, h * 0.5)
    return poly([
      [0, 0],
      [w - s, 0],
      [w, h / 2],
      [w - s, h],
      [0, h],
    ])
  },
  plus: cross,
  mathPlus: cross,
  mathMinus: (w, h) =>
    poly([
      [w * 0.1, h * 0.35],
      [w * 0.9, h * 0.35],
      [w * 0.9, h * 0.65],
      [w * 0.1, h * 0.65],
    ]),
  donut: (w, h) => {
    const t = Math.min(w, h) * 0.25
    const inner = `M${t} ${h / 2} A${w / 2 - t} ${h / 2 - t} 0 1 1 ${w - t} ${h / 2} A${w / 2 - t} ${h / 2 - t} 0 1 1 ${t} ${h / 2} Z`
    return `${ellipse(w, h, 0)} ${inner}`
  },
  heart: (w, h) =>
    `M${w / 2} ${h * 0.25} C${w * 0.5} 0 0 0 0 ${h * 0.3} C0 ${h * 0.6} ${w * 0.4} ${h * 0.75} ${w / 2} ${h} C${w * 0.6} ${h * 0.75} ${w} ${h * 0.6} ${w} ${h * 0.3} C${w} 0 ${w * 0.5} 0 ${w / 2} ${h * 0.25} Z`,
  cloud: (w, h) =>
    `M${w * 0.25} ${h * 0.85} A${w * 0.2} ${h * 0.2} 0 0 1 ${w * 0.15} ${h * 0.45} A${w * 0.22} ${h * 0.25} 0 0 1 ${w * 0.45} ${h * 0.2} A${w * 0.22} ${h * 0.25} 0 0 1 ${w * 0.8} ${h * 0.35} A${w * 0.2} ${h * 0.22} 0 0 1 ${w * 0.8} ${h * 0.85} Z`,
  flowChartProcess: (w, h) => `M0 0 H${w} V${h} H0 Z`,
  flowChartAlternateProcess: (w, h) => roundRect(w, h, Math.min(w, h) * 0.17),
  flowChartTerminator: (w, h) => roundRect(w, h, Math.min(w, h) / 2),
  flowChartDocument: (w, h) =>
    `M0 0 H${w} V${h * 0.8} C${w * 0.75} ${h * 0.6} ${w * 0.5} ${h * 1.05} ${w * 0.25} ${h * 0.95} C${w * 0.12} ${h * 0.9} 0 ${h * 0.85} 0 ${h * 0.8} Z`,
  can: (w, h) => {
    const e = h * 0.1
    return `M0 ${e} A${w / 2} ${e} 0 0 1 ${w} ${e} V${h - e} A${w / 2} ${e} 0 0 1 0 ${h - e} Z M0 ${e} A${w / 2} ${e} 0 0 0 ${w} ${e}`
  },
  pie: (w, h) => `M${w / 2} ${h / 2} L${w} ${h / 2} A${w / 2} ${h / 2} 0 1 1 ${w / 2} 0 Z`,
  arc: (w, h) => `M${w / 2} 0 A${w / 2} ${h / 2} 0 0 1 ${w} ${h / 2}`,
  lightningBolt: (w, h) =>
    poly([
      [w * 0.4, 0],
      [w * 0.65, h * 0.35],
      [w * 0.55, h * 0.4],
      [w * 0.85, h * 0.7],
      [w * 0.75, h * 0.73],
      [w, h],
      [w * 0.45, h * 0.78],
      [w * 0.55, h * 0.74],
      [w * 0.2, h * 0.48],
      [w * 0.32, h * 0.43],
      [0, h * 0.18],
    ]),
  wedgeRectCallout: (w, h) =>
    poly([
      [0, 0],
      [w, 0],
      [w, h * 0.75],
      [w * 0.4, h * 0.75],
      [w * 0.2, h],
      [w * 0.25, h * 0.75],
      [0, h * 0.75],
    ]),
  wedgeRoundRectCallout: (w, h) =>
    `${roundRect(w, h * 0.75, Math.min(w, h) * 0.1)} M${w * 0.25} ${h * 0.74} L${w * 0.2} ${h} L${w * 0.4} ${h * 0.74} Z`,
  frame: (w, h) => {
    const t = Math.min(w, h) * 0.12
    return `M0 0 H${w} V${h} H0 Z M${t} ${t} V${h - t} H${w - t} V${t} Z`
  },
}

export const OPEN_SHAPES = new Set(['line', 'arc'])

export const isKnownShape = (name: string): boolean => name in SHAPES || name === 'line'

/** The SVG path for a preset shape, or undefined when the preview doesn't know it. */
export const shapePath = (name: string, w: number, h: number, radius = 0): string | undefined =>
  SHAPES[name]?.(w, h, radius)
