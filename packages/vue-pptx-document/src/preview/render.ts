import { h, type CSSProperties, type VNode } from 'vue'

import { cssColor, opacity } from '../core/colors'
import { BULLET_INDENT, DEFAULT_FONT_SIZE, paragraphs, textInsets } from '../core/measure'
import type { DeckModel, Frame, MasterModel, SlideModel, SlideObject, TableCellModel, TextRun } from '../core/model'
import { PT_PER_IN } from '../core/units'
import { renderChart } from './chart'
import { OPEN_SHAPES, shapePath } from './shapes'

type Options = Record<string, unknown>

/** Converts inches to container query units, so a slide scales with its width. */
export interface Scale {
  /** Inches to a CSS length. */
  in: (inches: number) => string
  /** Points to a CSS length. */
  pt: (points: number) => string
  width: number
  height: number
}

export const createScale = (width: number, height: number): Scale => ({
  in: (inches) => `${round((inches / width) * 100)}cqw`,
  pt: (points) => `${round((points / PT_PER_IN / width) * 100)}cqw`,
  width,
  height,
})

const round = (value: number): number => Math.round(value * 10000) / 10000

const position = (frame: Frame, scale: Scale, options: Options = {}): CSSProperties => {
  const transforms: string[] = []
  if (typeof options.rotate === 'number' && options.rotate !== 0) transforms.push(`rotate(${options.rotate}deg)`)
  if (options.flipH) transforms.push('scaleX(-1)')
  if (options.flipV) transforms.push('scaleY(-1)')
  return {
    position: 'absolute',
    left: `${round((frame.x / scale.width) * 100)}%`,
    top: `${round((frame.y / scale.height) * 100)}%`,
    width: `${round((frame.w / scale.width) * 100)}%`,
    height: `${round((frame.h / scale.height) * 100)}%`,
    boxSizing: 'border-box',
    ...(transforms.length > 0 ? { transform: transforms.join(' ') } : {}),
  }
}

const fillOf = (fill: unknown): string | undefined => {
  if (!fill || typeof fill !== 'object') return undefined
  const { color, transparency, type } = fill as { color?: string; transparency?: number; type?: string }
  if (type === 'none' || !color) return undefined
  return withAlpha(cssColor(color), opacity(transparency))
}

const withAlpha = (color: string, alpha: number): string => {
  if (alpha >= 1 || !color.startsWith('#')) return color
  const hex = color.slice(1)
  const [r, g, b] = [0, 2, 4].map((offset) => parseInt(hex.slice(offset, offset + 2), 16))
  return `rgba(${r}, ${g}, ${b}, ${alpha})`
}

interface Stroke {
  color: string
  width: number
  dash?: string
}

/** A shape or text box outline, with its width in points. */
const strokeOf = (line: unknown): Stroke | undefined => {
  if (!line || typeof line !== 'object') return undefined
  const options = line as {
    color?: string
    width?: number
    size?: number
    pt?: number
    type?: string
    dashType?: string
    transparency?: number
  }
  if (options.type === 'none') return undefined
  const width = options.width ?? options.size ?? options.pt ?? 1
  const dash = options.dashType && options.dashType !== 'solid' ? options.dashType : undefined
  return {
    color: withAlpha(cssColor(options.color ?? '333333'), opacity(options.transparency)),
    width,
    ...(dash ? { dash: /dot/i.test(dash) ? '1 2' : '4 3' } : {}),
  }
}

const shadowOf = (shadow: unknown, scale: Scale): string | undefined => {
  if (!shadow || typeof shadow !== 'object') return undefined
  const options = shadow as {
    type?: string
    blur?: number
    offset?: number
    angle?: number
    color?: string
    opacity?: number
  }
  if (options.type === 'none') return undefined
  const angle = ((options.angle ?? 45) * Math.PI) / 180
  const offset = options.offset ?? 2
  const x = scale.pt(Math.cos(angle) * offset)
  const y = scale.pt(Math.sin(angle) * offset)
  return `drop-shadow(${x} ${y} ${scale.pt(options.blur ?? 3)} ${withAlpha(cssColor(options.color ?? '000000'), options.opacity ?? 0.35)})`
}

/** The SVG for a preset shape filling its box, behind any text. */
const renderShapeSvg = (name: string, frame: Frame, options: Options, scale: Scale, defaultFill?: string): VNode => {
  // The viewBox is in points, so stroke widths and radii read in their own units.
  const vw = Math.max(frame.w * PT_PER_IN, 0.01)
  const vh = Math.max(frame.h * PT_PER_IN, 0.01)
  const fill = fillOf(options.fill) ?? defaultFill ?? 'none'
  const stroke = strokeOf(options.line)
  const open = OPEN_SHAPES.has(name)
  const strokeProps = stroke
    ? { stroke: stroke.color, 'stroke-width': stroke.width, 'stroke-dasharray': stroke.dash }
    : open
      ? { stroke: '#333333', 'stroke-width': 1 }
      : {}
  const filter = shadowOf(options.shadow, scale)
  const svg = (children: VNode[]): VNode =>
    h(
      'svg',
      {
        viewBox: `0 0 ${vw} ${vh}`,
        preserveAspectRatio: 'none',
        width: '100%',
        height: '100%',
        overflow: 'visible',
        style: { position: 'absolute', inset: 0, ...(filter ? { filter } : {}) },
        'aria-hidden': 'true',
      },
      children,
    )
  if (name === 'line') {
    return svg([
      h('line', {
        x1: 0,
        y1: 0,
        x2: vw,
        y2: vh,
        ...strokeProps,
        'vector-effect': 'non-scaling-stroke',
      }),
    ])
  }
  const radius = typeof options.rectRadius === 'number' ? options.rectRadius * PT_PER_IN : Math.min(vw, vh) * 0.16
  const d = shapePath(name, vw, vh, radius)
  if (!d) {
    return svg([
      h('rect', {
        width: vw,
        height: vh,
        fill,
        stroke: stroke?.color ?? '#999999',
        'stroke-width': stroke?.width ?? 1,
        'stroke-dasharray': '4 3',
      }),
      h('title', `${name} (previewed as a rectangle)`),
    ])
  }
  return svg([h('path', { d, fill: open ? 'none' : fill, 'fill-rule': 'evenodd', ...strokeProps })])
}

interface TextContext {
  scale: Scale
  slideColor?: string | undefined
  fontFace?: string | undefined
}

const runStyle = (options: Options, box: Options, context: TextContext): CSSProperties => {
  // eslint-disable-next-line @typescript-eslint/no-unnecessary-type-parameters -- names the option's type at each call
  const pick = <T>(key: string): T | undefined => (options[key] ?? box[key]) as T | undefined
  const size = pick<number>('fontSize') ?? DEFAULT_FONT_SIZE
  const underline = pick<{ style?: string } | boolean>('underline')
  const strike = pick<unknown>('strike')
  const decorations = [
    underline && (typeof underline !== 'object' || underline.style !== 'none') ? 'underline' : '',
    strike ? 'line-through' : '',
  ].filter(Boolean)
  const superscript = pick<boolean>('superscript')
  const subscript = pick<boolean>('subscript')
  const highlight = pick<string>('highlight')
  const charSpacing = pick<number>('charSpacing')
  const color = withAlpha(
    cssColor(pick<string>('color') ?? context.slideColor ?? '000000', '#000000'),
    opacity(pick<number>('transparency')),
  )
  return {
    fontFamily: `${JSON.stringify(pick<string>('fontFace') ?? context.fontFace ?? 'Arial')}, Arial, sans-serif`,
    fontSize: context.scale.pt(superscript || subscript ? size * 0.7 : size),
    fontWeight: pick<boolean>('bold') ? 700 : 400,
    fontStyle: pick<boolean>('italic') ? 'italic' : 'normal',
    color,
    ...(decorations.length > 0 ? { textDecoration: decorations.join(' ') } : {}),
    ...(superscript ? { verticalAlign: 'super' } : subscript ? { verticalAlign: 'sub' } : {}),
    ...(highlight ? { backgroundColor: cssColor(highlight) } : {}),
    ...(charSpacing ? { letterSpacing: context.scale.pt(charSpacing) } : {}),
  }
}

const bulletText = (bullet: unknown, index: number): string => {
  if (!bullet) return ''
  if (typeof bullet === 'object') {
    const options = bullet as {
      type?: string
      code?: string
      characterCode?: string
      startAt?: number
      numberStartAt?: number
      numberType?: string
    }
    if (options.type === 'number') {
      const n = (options.numberStartAt ?? options.startAt ?? 1) + index
      const type = options.numberType ?? 'arabicPeriod'
      const value = type.startsWith('alphaLc')
        ? String.fromCharCode(96 + n)
        : type.startsWith('alphaUc')
          ? String.fromCharCode(64 + n)
          : type.startsWith('roman')
            ? toRoman(n, type.startsWith('romanLc'))
            : String(n)
      return type.endsWith('ParenR')
        ? `${value})`
        : type.endsWith('ParenBoth')
          ? `(${value})`
          : type.endsWith('Plain')
            ? value
            : `${value}.`
    }
    const code = options.code ?? options.characterCode
    if (code) return String.fromCharCode(parseInt(code, 16))
  }
  return '•'
}

const toRoman = (n: number, lower: boolean): string => {
  const table: [number, string][] = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ]
  let rest = n
  let out = ''
  for (const [value, symbol] of table) {
    while (rest >= value) {
      out += symbol
      rest -= value
    }
  }
  return lower ? out.toLowerCase() : out
}

/**
 * A link target the preview may put in an `href`: web, mail and phone links, and relative ones.
 * Decks can come from anywhere, so `javascript:` and other schemes never become clickable.
 */
export const safeUrl = (url: unknown): string | undefined => {
  if (typeof url !== 'string' || url.trim() === '') return undefined
  const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(url.replace(/[\s\p{Cc}]/gu, ''))?.[1]?.toLowerCase()
  return scheme === undefined || ['http', 'https', 'mailto', 'tel'].includes(scheme) ? url : undefined
}

const runNode = (run: TextRun, box: Options, context: TextContext, key: number): VNode => {
  const options = (run.options ?? {}) as Options
  const style = runStyle(options, box, context)
  const hyperlink = (options.hyperlink ?? box.hyperlink) as { url?: string; tooltip?: string } | undefined
  const href = safeUrl(hyperlink?.url)
  if (href) {
    return h(
      'a',
      {
        key,
        href,
        title: hyperlink?.tooltip,
        target: '_blank',
        rel: 'noopener noreferrer',
        style: { ...style, color: style.color },
      },
      run.text,
    )
  }
  return h('span', { key, style }, run.text)
}

/** Paragraphs of a text box, with bullets, numbering and alignment. */
export const renderParagraphs = (runs: TextRun[], box: Options, context: TextContext): VNode[] => {
  const counters = new Map<number, { scheme: string; count: number }>()
  return paragraphs(runs).map((paragraph, index) => {
    const first = (paragraph[0]?.options ?? {}) as Options
    // eslint-disable-next-line @typescript-eslint/no-unnecessary-type-parameters -- names the option's type at each call
    const pick = <T>(key: string): T | undefined => (first[key] ?? box[key]) as T | undefined
    const align = pick<string>('align')
    const bullet = pick<unknown>('bullet')
    const level = pick<number>('indentLevel') ?? 0
    const size = (first.fontSize ?? box.fontSize ?? DEFAULT_FONT_SIZE) as number
    const lineSpacing = pick<number>('lineSpacing')
    const multiple = pick<number>('lineSpacingMultiple') ?? 1
    // Numbering continues while paragraphs at a level keep the same scheme; anything else restarts it.
    const scheme =
      typeof bullet === 'object' && bullet && (bullet as { type?: string }).type === 'number'
        ? ((bullet as { numberType?: string }).numberType ?? 'arabicPeriod')
        : undefined
    let marker = ''
    if (scheme) {
      const previous = counters.get(level)
      const count = previous?.scheme === scheme ? previous.count : 0
      counters.set(level, { scheme, count: count + 1 })
      marker = bulletText(bullet, count)
    } else {
      counters.delete(level)
      if (bullet) marker = bulletText(bullet, 0)
    }
    const indent = (level > 0 ? level * BULLET_INDENT : 0) + (bullet ? BULLET_INDENT : 0)
    const empty = paragraph.every((run) => run.text === '')
    return h(
      'p',
      {
        key: index,
        style: {
          margin: 0,
          position: 'relative',
          textAlign: (align === 'center'
            ? 'center'
            : align === 'right'
              ? 'right'
              : align === 'justify'
                ? 'justify'
                : 'left') as CSSProperties['textAlign'],
          lineHeight: lineSpacing ? context.scale.pt(lineSpacing) : String(1.2 * multiple),
          paddingLeft: indent ? context.scale.in(indent) : undefined,
          paddingTop: pick<number>('paraSpaceBefore')
            ? context.scale.pt(pick<number>('paraSpaceBefore') ?? 0)
            : undefined,
          paddingBottom: pick<number>('paraSpaceAfter')
            ? context.scale.pt(pick<number>('paraSpaceAfter') ?? 0)
            : undefined,
          fontSize: context.scale.pt(size),
        },
      },
      [
        marker
          ? h(
              'span',
              {
                'aria-hidden': 'true',
                style: {
                  ...runStyle(first, box, context),
                  position: 'absolute',
                  left: context.scale.in(level * BULLET_INDENT),
                  width: context.scale.in(BULLET_INDENT),
                },
              },
              marker,
            )
          : null,
        ...(empty ? [h('br')] : paragraph.map((run, key) => runNode(run, box, context, key))),
      ],
    )
  })
}

const VALIGN: Record<string, string> = { top: 'flex-start', middle: 'center', bottom: 'flex-end' }

const renderText = (object: Extract<SlideObject, { kind: 'text' }>, context: TextContext): VNode => {
  const options = object.options as Options
  const { scale } = context
  const insets = textInsets(options.margin)
  const shape = typeof options.shape === 'string' ? options.shape : undefined
  const stroke = strokeOf(options.line)
  const fill = fillOf(options.fill)
  const vertical = typeof options.vert === 'string' && options.vert !== 'horz'
  // Text boxes default to the middle, as PptxGenJS writes them.
  const valign = VALIGN[(options.valign as string | undefined) ?? 'middle'] ?? 'center'
  return h(
    'div',
    {
      'data-pptx-object': 'text',
      style: {
        ...position(object.frame, scale, options),
        display: 'flex',
        flexDirection: 'column',
        justifyContent: valign,
        padding: `${scale.in(insets.top)} ${scale.in(insets.right)} ${scale.in(insets.bottom)} ${scale.in(insets.left)}`,
        whiteSpace: options.wrap === false ? 'pre' : 'pre-wrap',
        overflowWrap: 'break-word',
        ...(shape && shape !== 'rect'
          ? {}
          : {
              ...(fill ? { backgroundColor: fill } : {}),
              ...(stroke
                ? { border: `${scale.pt(stroke.width)} ${stroke.dash ? 'dashed' : 'solid'} ${stroke.color}` }
                : {}),
            }),
        ...(vertical ? { writingMode: 'vertical-rl' as const } : {}),
      },
    },
    [
      shape && shape !== 'rect' ? renderShapeSvg(shape, object.frame, options, scale) : null,
      h('div', { style: { position: 'relative' } }, renderParagraphs(object.runs, options, context)),
    ],
  )
}

const renderImage = (object: Extract<SlideObject, { kind: 'image' }>, scale: Scale): VNode => {
  const options = object.options as Options
  const sizing = options.sizing as { type?: string } | undefined
  const fit =
    sizing?.type === 'contain' ? 'contain' : sizing?.type === 'cover' || sizing?.type === 'crop' ? 'cover' : 'fill'
  const image = h('img', {
    src: object.src,
    alt: (options.altText as string | undefined) ?? '',
    draggable: false,
    style: {
      width: '100%',
      height: '100%',
      objectFit: fit,
      display: 'block',
      borderRadius: options.rounding ? '50%' : undefined,
      opacity: opacity(options.transparency),
    },
  })
  const link = options.hyperlink as { url?: string } | undefined
  return h('div', { 'data-pptx-object': 'image', style: position(object.frame, scale, options) }, [
    safeUrl(link?.url)
      ? h('a', { href: safeUrl(link?.url), target: '_blank', rel: 'noopener noreferrer' }, [image])
      : image,
  ])
}

interface Border {
  type?: string
  color?: string
  pt?: number
}

const cssBorder = (border: Border | undefined, scale: Scale): string | undefined => {
  if (!border || border.type === 'none') return undefined
  return `${scale.pt(border.pt ?? 1)} ${border.type === 'dash' ? 'dashed' : 'solid'} ${cssColor(border.color ?? '666666')}`
}

const renderTable = (object: Extract<SlideObject, { kind: 'table' }>, context: TextContext): VNode => {
  const { scale } = context
  const options = object.options as Options
  const columns = Math.max(
    0,
    ...object.rows.map((row) => row.reduce((sum, cell) => sum + (cell.options?.colspan ?? 1), 0)),
  )
  const colW = options.colW
  const widths = Array.isArray(colW)
    ? (colW as number[])
    : typeof colW === 'number'
      ? Array.from({ length: columns }, () => colW)
      : Array.from({ length: columns }, () => object.frame.w / Math.max(columns, 1))
  const rowH = options.rowH
  const heightOf = (index: number): number | undefined =>
    Array.isArray(rowH) ? (rowH as number[])[index] : typeof rowH === 'number' ? rowH : undefined
  const tableBorder = options.border as Border | Border[] | undefined
  const cell = (model: TableCellModel, key: number): VNode => {
    const cellOptions = { ...options, ...(model.options as Options | undefined) }
    delete cellOptions.colW
    delete cellOptions.rowH
    const margin = cellOptions.margin
    const [top, right, bottom, left] = Array.isArray(margin)
      ? (margin as number[])
      : typeof margin === 'number'
        ? [margin, margin, margin, margin]
        : [0.05, 0.1, 0.05, 0.1]
    const border = (cellOptions.border ?? tableBorder) as Border | Border[] | undefined
    const [bt, br, bb, bl] = Array.isArray(border) ? border : [border, border, border, border]
    const runs: TextRun[] =
      typeof model.text === 'string' || model.text === undefined ? [{ text: model.text ?? '' }] : model.text
    const fill = fillOf(cellOptions.fill)
    return h(
      'td',
      {
        key,
        colspan: model.options?.colspan,
        rowspan: model.options?.rowspan,
        style: {
          padding: `${scale.in(top ?? 0)} ${scale.in(right ?? 0)} ${scale.in(bottom ?? 0)} ${scale.in(left ?? 0)}`,
          verticalAlign:
            cellOptions.valign === 'middle' ? 'middle' : cellOptions.valign === 'bottom' ? 'bottom' : 'top',
          backgroundColor: fill,
          borderTop: cssBorder(bt, scale),
          borderRight: cssBorder(br, scale),
          borderBottom: cssBorder(bb, scale),
          borderLeft: cssBorder(bl, scale),
          whiteSpace: 'pre-wrap',
        },
      },
      renderParagraphs(runs, { fontSize: 12, ...cellOptions }, context),
    )
  }
  return h(
    'div',
    { 'data-pptx-object': 'table', style: { ...position(object.frame, scale), height: 'auto' } },
    h('table', { style: { width: '100%', borderCollapse: 'collapse', tableLayout: 'fixed' } }, [
      h(
        'colgroup',
        widths.map((width, index) => h('col', { key: index, style: { width: scale.in(width) } })),
      ),
      h(
        'tbody',
        object.rows.map((row, index) =>
          h(
            'tr',
            { key: index, style: { height: heightOf(index) ? scale.in(heightOf(index) ?? 0) : undefined } },
            row.map(cell),
          ),
        ),
      ),
    ]),
  )
}

const renderMedia = (object: Extract<SlideObject, { kind: 'media' }>, scale: Scale): VNode => {
  const options = object.options as Options & {
    type?: string
    path?: string
    data?: string
    link?: string
    cover?: string
  }
  const src = options.data ? `data:${options.data}` : options.path
  const box = { ...position(object.frame, scale), backgroundColor: '#000000' }
  if (options.type === 'online' || !src) {
    return h(
      'div',
      {
        'data-pptx-object': 'media',
        style: { ...box, display: 'grid', placeItems: 'center', color: '#FFFFFF', fontSize: scale.pt(12) },
      },
      [
        safeUrl(options.link)
          ? h(
              'a',
              {
                href: safeUrl(options.link),
                target: '_blank',
                rel: 'noopener noreferrer',
                style: { color: 'inherit' },
              },
              '▶ Online video',
            )
          : '▶ Media',
      ],
    )
  }
  const poster = options.cover
    ? options.cover.startsWith('data:')
      ? options.cover
      : `data:${options.cover}`
    : undefined
  return h('div', { 'data-pptx-object': 'media', style: box }, [
    h(options.type === 'audio' ? 'audio' : 'video', {
      src,
      controls: true,
      preload: 'none',
      poster,
      style: { width: '100%', height: '100%', display: 'block' },
    }),
  ])
}

const renderObject = (object: SlideObject, context: TextContext, key: number): VNode | null => {
  const { scale } = context
  switch (object.kind) {
    case 'text':
      return h(renderText(object, context), { key })
    case 'shape':
      return h(
        'div',
        {
          key,
          'data-pptx-object': 'shape',
          'data-shape': object.shape,
          style: position(object.frame, scale, object.options as Options),
        },
        [renderShapeSvg(object.shape, object.frame, object.options as Options, scale)],
      )
    case 'image':
      return h(renderImage(object, scale), { key })
    case 'table':
      return h(renderTable(object, context), { key })
    case 'chart':
      return h('div', { key, 'data-pptx-object': 'chart', style: position(object.frame, scale) }, [
        renderChart(object, context.fontFace),
      ])
    case 'media':
      return h(renderMedia(object, scale), { key })
    case 'placeholder':
      return object.text
        ? h(
            renderText(
              { kind: 'text', frame: object.frame, runs: [{ text: object.text }], options: object.options },
              context,
            ),
            { key },
          )
        : null
  }
}

/** Text placed into a master's placeholder takes the placeholder's frame and options. */
const fillPlaceholders = (objects: SlideObject[], master: MasterModel | undefined): SlideObject[] =>
  objects.map((object) => {
    if (object.kind !== 'text' || typeof object.options.placeholder !== 'string' || !master) return object
    const target = master.objects.find(
      (item) => item.kind === 'placeholder' && item.name === object.options.placeholder,
    )
    if (!target) return object
    return { ...object, frame: target.frame, options: { ...(target.options as Options), ...object.options } }
  })

const backgroundStyle = (slide: SlideModel, master: MasterModel | undefined): CSSProperties => {
  const background = slide.background ?? master?.background
  if (background?.src) {
    return {
      backgroundColor: '#FFFFFF',
      backgroundImage: `url(${JSON.stringify(background.src)})`,
      backgroundSize: 'cover',
      backgroundPosition: 'center',
    }
  }
  return {
    backgroundColor: withAlpha(cssColor(background?.color ?? 'FFFFFF', '#FFFFFF'), opacity(background?.transparency)),
  }
}

export interface RenderSlideOptions {
  /** Shown in place of the slide number field. */
  number: number
}

/** The slide as positioned HTML and SVG. Lengths are container query units of the slide's width. */
export const renderSlide = (model: DeckModel, index: number, { number }: RenderSlideOptions): VNode[] => {
  const slide = model.slides[index]
  if (!slide) return []
  const master = model.masters.find((item) => item.name === slide.master)
  const scale = createScale(model.layout.width, model.layout.height)
  const context: TextContext = { scale, slideColor: slide.color, fontFace: model.theme?.bodyFontFace }
  const nodes: VNode[] = [
    h('div', {
      key: 'background',
      'data-pptx-background': '',
      style: { position: 'absolute', inset: 0, ...backgroundStyle(slide, master) },
    }),
  ]
  master?.objects.forEach((object, key) => {
    if (object.kind === 'placeholder') return
    const node = renderObject(object, context, key)
    if (node) nodes.push(h('div', { key: `master-${key}`, 'data-pptx-master': '' }, [node]))
  })
  fillPlaceholders(slide.objects, master).forEach((object, key) => {
    const node = renderObject(object, context, key)
    if (node) nodes.push(h(node, { key: `object-${key}` }))
  })
  const slideNumber = slide.slideNumber ?? master?.slideNumber
  if (slideNumber) {
    const {
      x = 0,
      y = 0,
      w = 0.6,
      h: height = 0.3,
      ...rest
    } = slideNumber as Options & { x?: number; y?: number; w?: number; h?: number }
    nodes.push(
      h(
        renderText(
          {
            kind: 'text',
            frame: { x, y, w, h: height },
            runs: [{ text: String(number) }],
            options: { fontSize: 10, ...rest },
          },
          context,
        ),
        { key: 'slide-number', 'data-pptx-slide-number': '' },
      ),
    )
  }
  return nodes
}
