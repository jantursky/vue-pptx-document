import type PptxGenJS from 'pptxgenjs'

import { normalizeColors } from './colors'
import { layoutFlex } from './flex'
import { measureText } from './measure'
import type { DeckLayout, DeckModel, Frame, MasterModel, SlideModel, SlideObject } from './model'
import { TYPES } from './nodes'
import { backgroundOf, DRAWABLE, runsOf, tableRows, textOptions, toObject } from './objects'
import type { TreeNode } from './snapshot'
import { round, toInches } from './units'

/** PptxGenJS's built-in layouts, in inches. */
export const LAYOUTS = {
  '16x9': { name: 'LAYOUT_16x9', width: 10, height: 5.625 },
  '16x10': { name: 'LAYOUT_16x10', width: 10, height: 6.25 },
  '4x3': { name: 'LAYOUT_4x3', width: 10, height: 7.5 },
  wide: { name: 'LAYOUT_WIDE', width: 13.333, height: 7.5 },
} as const satisfies Record<string, DeckLayout>

export type LayoutName = keyof typeof LAYOUTS

export const resolveLayout = (value: unknown): DeckLayout => {
  if (value == null) return LAYOUTS['16x9']
  if (typeof value === 'string') {
    const key = value.replace(/^LAYOUT_/i, '').toLowerCase()
    const layout = (LAYOUTS as Record<string, DeckLayout>)[key]
    if (!layout)
      throw new Error(`[vue-pptx-document] Unknown layout "${value}". Use 16x9, 16x10, 4x3, wide or { width, height }.`)
    return layout
  }
  const { width, height } = value as { width: unknown; height: unknown }
  const w = toInches(width)
  const h = toInches(height)
  if (!w || !h) throw new Error('[vue-pptx-document] A custom layout needs a width and a height.')
  return { name: 'CUSTOM', width: w, height: h }
}

const roundFrame = (frame: Frame): Frame => ({
  x: round(frame.x),
  y: round(frame.y),
  w: round(frame.w),
  h: round(frame.h),
})

/** The height an object takes when it sets no `h`: text and tables fit their content. */
const autoHeight = (node: TreeNode, w: number, remaining: number): number => {
  if (node.type === TYPES.text || (node.type === TYPES.shape && runsOf(node).length > 0)) {
    return measureText(runsOf(node), textOptions(node.props), w).height
  }
  if (node.type === TYPES.table) {
    const rows = tableRows(node)
    const rowH = node.props.rowH
    if (Array.isArray(rowH)) return rowH.reduce<number>((sum, item) => sum + (toInches(item) ?? 0), 0)
    const fontSize = typeof node.props.fontSize === 'number' ? node.props.fontSize : 12
    return rows.length * (toInches(rowH) ?? (fontSize / 72) * 1.2 + 0.1)
  }
  return remaining
}

/** An object's frame from `x`, `y`, `w` and `h`, relative to `box`. Percentages are of the box. */
const absoluteFrame = (node: TreeNode, box: Frame): Frame => {
  const { props } = node
  const x = toInches(props.x, box.w) ?? 0
  const y = toInches(props.y, box.h) ?? 0
  let w = toInches(props.w, box.w)
  if (w === undefined && node.type === TYPES.table && Array.isArray(props.colW)) {
    w = props.colW.reduce<number>((sum, item) => sum + (toInches(item, box.w) ?? 0), 0)
  }
  w ??= Math.max(0, box.w - x)
  const h = toInches(props.h, box.h) ?? autoHeight(node, w, Math.max(0, box.h - y))
  return { x: box.x + x, y: box.y + y, w, h }
}

const warned = new Set<string>()
const warn = (message: string): void => {
  if (warned.has(message)) return
  warned.add(message)
  console.warn(`[vue-pptx-document] ${message}`)
}

/** Lays out a slide's (or master's) children: absolutely, or with Yoga for flex containers. */
const resolveObjects = async (container: TreeNode, box: Frame, flex: boolean): Promise<SlideObject[]> => {
  const out: SlideObject[] = []
  if (flex) {
    await layoutFlex(container, box, out)
    return out.map((object) => ({ ...object, frame: roundFrame(object.frame) }))
  }
  for (const child of container.children) {
    if (typeof child === 'string' || child.type === TYPES.notes || isHidden(child)) continue
    if (child.type === TYPES.view) {
      await layoutFlex(child, absoluteFrame(child, box), out)
    } else if (DRAWABLE.has(child.type)) {
      const object = toObject(child, absoluteFrame(child, box))
      if (object) out.push(object)
    } else {
      warn(`<${child.type}> can't be placed directly on a slide; it is skipped.`)
    }
  }
  return out.map((object) => ({ ...object, frame: roundFrame(object.frame) }))
}

/** `style.display: 'none'` hides an object on absolute slides as it does in flex containers. */
const isHidden = (node: TreeNode): boolean => (node.props.style as { display?: string } | undefined)?.display === 'none'

const slideNumberOf = (value: unknown, layout: DeckLayout): PptxGenJS.SlideNumberProps | undefined => {
  if (!value) return undefined
  const defaults = {
    x: layout.width - 1,
    y: layout.height - 0.45,
    w: 0.6,
    h: 0.3,
    fontSize: 10,
    align: 'right' as const,
  }
  if (value === true) return defaults
  const options = { ...defaults, ...(value as Record<string, unknown>) }
  for (const key of ['x', 'y', 'w', 'h'] as const) {
    const base = key === 'x' || key === 'w' ? layout.width : layout.height
    options[key] = toInches(options[key], base) ?? defaults[key]
  }
  return normalizeColors(options)
}

/** Speaker notes: the slide's `notes` prop, then every `<Notes>` on the slide or inside its views, in order. */
const notesOf = (slide: TreeNode): string | undefined => {
  const parts: string[] = []
  if (typeof slide.props.notes === 'string') parts.push(slide.props.notes)
  const collect = (node: TreeNode): void => {
    for (const child of node.children) {
      if (typeof child === 'string') continue
      if (child.type === TYPES.view) collect(child)
      if (child.type !== TYPES.notes) continue
      const runs = runsOf(child)
      parts.push(runs.map((run) => run.text + (run.options?.breakLine ? '\n' : '')).join(''))
    }
  }
  collect(slide)
  return parts.length > 0 ? parts.join('\n') : undefined
}

const defined = <T extends object>(value: T): T =>
  Object.fromEntries(Object.entries(value).filter(([, item]) => item !== undefined)) as T

/** Turns the snapshot of a `<Presentation>` into the resolved model. */
export const resolve = async (tree: TreeNode): Promise<DeckModel> => {
  const { props } = tree
  const layout = resolveLayout(props.layout)
  const box: Frame = { x: 0, y: 0, w: layout.width, h: layout.height }
  const masters: MasterModel[] = []
  const slides: SlideModel[] = []
  for (const child of tree.children) {
    if (typeof child === 'string') continue
    const flex = child.props.layout === 'flex'
    if (child.type === TYPES.master) {
      const name = typeof child.props.name === 'string' ? child.props.name : `MASTER_${masters.length + 1}`
      masters.push(
        defined({
          name,
          background: backgroundOf(child.props.background),
          margin: child.props.margin as PptxGenJS.Margin | undefined,
          slideNumber: slideNumberOf(child.props.slideNumber, layout),
          objects: await resolveObjects(child, box, flex),
        }),
      )
    } else if (child.type === TYPES.slide) {
      slides.push(
        defined({
          background: backgroundOf(child.props.background),
          color: normalizeColors({ color: child.props.color as string | undefined }).color,
          hidden: child.props.hidden === true ? true : undefined,
          master: typeof child.props.master === 'string' ? child.props.master : undefined,
          section: typeof child.props.section === 'string' ? child.props.section : undefined,
          notes: notesOf(child),
          slideNumber: slideNumberOf(child.props.slideNumber, layout),
          objects: await resolveObjects(child, box, flex),
        }),
      )
    } else {
      warn(`<${child.type}> must be inside a <Slide> or <Master>; it is skipped.`)
    }
  }
  for (const slide of slides) {
    if (slide.master && !masters.some((master) => master.name === slide.master)) {
      warn(`No <Master name="${slide.master}">; the slide uses the default master.`)
    }
  }
  const theme = defined({
    headFontFace:
      (props.headFontFace as string | undefined) ?? (props.theme as PptxGenJS.ThemeProps | undefined)?.headFontFace,
    bodyFontFace:
      (props.bodyFontFace as string | undefined) ?? (props.theme as PptxGenJS.ThemeProps | undefined)?.bodyFontFace,
  })
  return defined({
    layout,
    meta: defined({
      title: props.title as string | undefined,
      author: props.author as string | undefined,
      company: props.company as string | undefined,
      subject: props.subject as string | undefined,
      revision: props.revision === undefined ? undefined : String(props.revision as string | number),
    }),
    rtl: props.rtl === true ? true : undefined,
    theme: Object.keys(theme).length > 0 ? theme : undefined,
    masters,
    slides,
  })
}
