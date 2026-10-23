import type * as YogaModule from 'yoga-layout/load'
import type { Config, MeasureFunction, Node as YogaNode, Yoga as YogaInstance } from 'yoga-layout/load'

import { measureText } from './measure'
import type { Frame, SlideObject } from './model'
import { TYPES } from './nodes'
import { DRAWABLE, runsOf, textOptions, toObject, viewBackground } from './objects'
import type { TreeNode } from './snapshot'
import { isPercent, PX_PER_IN, toInches } from './units'

/** Flexbox style for `<View>`, `<Slide layout="flex">` and their children. Lengths as in `x`/`w`. */
export interface FlexStyle {
  display?: 'flex' | 'none'
  position?: 'relative' | 'absolute'
  flexDirection?: 'row' | 'column' | 'row-reverse' | 'column-reverse'
  flexWrap?: 'nowrap' | 'wrap' | 'wrap-reverse'
  flex?: number
  flexGrow?: number
  flexShrink?: number
  flexBasis?: number | string
  justifyContent?: 'flex-start' | 'center' | 'flex-end' | 'space-between' | 'space-around' | 'space-evenly'
  alignItems?: 'flex-start' | 'center' | 'flex-end' | 'stretch' | 'baseline'
  alignSelf?: 'auto' | 'flex-start' | 'center' | 'flex-end' | 'stretch' | 'baseline'
  alignContent?: 'flex-start' | 'center' | 'flex-end' | 'stretch' | 'space-between' | 'space-around' | 'space-evenly'
  width?: number | string
  height?: number | string
  minWidth?: number | string
  minHeight?: number | string
  maxWidth?: number | string
  maxHeight?: number | string
  aspectRatio?: number
  gap?: number | string
  rowGap?: number | string
  columnGap?: number | string
  padding?: number | string
  paddingTop?: number | string
  paddingRight?: number | string
  paddingBottom?: number | string
  paddingLeft?: number | string
  paddingHorizontal?: number | string
  paddingVertical?: number | string
  margin?: number | string
  marginTop?: number | string
  marginRight?: number | string
  marginBottom?: number | string
  marginLeft?: number | string
  marginHorizontal?: number | string
  marginVertical?: number | string
  top?: number | string
  right?: number | string
  bottom?: number | string
  left?: number | string
}

type Enums = typeof YogaModule

/** The Yoga instance and its enums, which live on the module rather than the instance. */
interface Yoga {
  Node: YogaInstance['Node']
  Config: YogaInstance['Config']
  enums: Enums
}

let yoga: Promise<Yoga> | undefined

/** Yoga's WebAssembly build loads on the first flex layout, so absolute-only decks never fetch it. */
export const loadYoga = (): Promise<Yoga> => {
  yoga ??= import('yoga-layout/load').then(async (enums) => {
    const instance = await enums.loadYoga()
    return { Node: instance.Node, Config: instance.Config, enums }
  })
  return yoga
}

type YogaValue = number | `${number}%` | 'auto' | undefined

/** Converts a length to Yoga's units: pixels at 96 per inch, or a percentage Yoga resolves itself. */
const px = (value: unknown): YogaValue => {
  if (value === 'auto') return 'auto'
  if (isPercent(value)) return value.trim() as `${number}%`
  const inches = toInches(value)
  return inches === undefined ? undefined : inches * PX_PER_IN
}

const pick = <T>(value: unknown, map: Record<string, T>): T | undefined =>
  typeof value === 'string' ? map[value] : undefined

type Size = Parameters<YogaNode['setMinWidth']>[0]

const applyStyle = (Y: Yoga, node: YogaNode, style: FlexStyle): void => {
  const { Align, Display, Edge, FlexDirection, Gutter, Justify, PositionType, Wrap } = Y.enums
  const align: Record<string, number> = {
    auto: Align.Auto,
    'flex-start': Align.FlexStart,
    start: Align.FlexStart,
    center: Align.Center,
    'flex-end': Align.FlexEnd,
    end: Align.FlexEnd,
    stretch: Align.Stretch,
    baseline: Align.Baseline,
    'space-between': Align.SpaceBetween,
    'space-around': Align.SpaceAround,
    'space-evenly': Align.SpaceEvenly,
  }
  const justify: Record<string, number> = {
    'flex-start': Justify.FlexStart,
    start: Justify.FlexStart,
    center: Justify.Center,
    'flex-end': Justify.FlexEnd,
    end: Justify.FlexEnd,
    'space-between': Justify.SpaceBetween,
    'space-around': Justify.SpaceAround,
    'space-evenly': Justify.SpaceEvenly,
  }
  const direction = pick(style.flexDirection, {
    row: FlexDirection.Row,
    column: FlexDirection.Column,
    'row-reverse': FlexDirection.RowReverse,
    'column-reverse': FlexDirection.ColumnReverse,
  })
  if (direction !== undefined) node.setFlexDirection(direction)
  const wrap = pick(style.flexWrap, {
    nowrap: Wrap.NoWrap,
    wrap: Wrap.Wrap,
    'wrap-reverse': Wrap.WrapReverse,
  })
  if (wrap !== undefined) node.setFlexWrap(wrap)
  if (style.display === 'none') node.setDisplay(Display.None)
  if (style.position === 'absolute') node.setPositionType(PositionType.Absolute)
  if (typeof style.justifyContent === 'string' && style.justifyContent in justify) {
    node.setJustifyContent(justify[style.justifyContent] as never)
  }
  if (typeof style.alignItems === 'string' && style.alignItems in align)
    node.setAlignItems(align[style.alignItems] as never)
  if (typeof style.alignSelf === 'string' && style.alignSelf in align)
    node.setAlignSelf(align[style.alignSelf] as never)
  if (typeof style.alignContent === 'string' && style.alignContent in align) {
    node.setAlignContent(align[style.alignContent] as never)
  }
  // `flex: 1` means what it does in CSS: grow and shrink from a zero basis.
  if (typeof style.flex === 'number') {
    node.setFlexGrow(style.flex)
    node.setFlexShrink(1)
    node.setFlexBasis(0)
  }
  if (typeof style.flexGrow === 'number') node.setFlexGrow(style.flexGrow)
  if (typeof style.flexShrink === 'number') node.setFlexShrink(style.flexShrink)
  if (style.flexBasis !== undefined) node.setFlexBasis(px(style.flexBasis))
  if (typeof style.aspectRatio === 'number') node.setAspectRatio(style.aspectRatio)

  const width = px(style.width)
  if (width !== undefined) node.setWidth(width)
  const height = px(style.height)
  if (height !== undefined) node.setHeight(height)
  const sizes = [
    ['minWidth', (value: Size) => node.setMinWidth(value)],
    ['minHeight', (value: Size) => node.setMinHeight(value)],
    ['maxWidth', (value: Size) => node.setMaxWidth(value)],
    ['maxHeight', (value: Size) => node.setMaxHeight(value)],
  ] as const
  for (const [key, set] of sizes) {
    const value = px(style[key])
    if (value !== undefined && value !== 'auto') set(value)
  }

  const gutters = [
    ['gap', Gutter.All],
    ['rowGap', Gutter.Row],
    ['columnGap', Gutter.Column],
  ] as const
  for (const [key, gutter] of gutters) {
    const value = px(style[key])
    if (value !== undefined && value !== 'auto') node.setGap(gutter, value)
  }

  const edges = [
    ['', Edge.All],
    ['Horizontal', Edge.Horizontal],
    ['Vertical', Edge.Vertical],
    ['Top', Edge.Top],
    ['Right', Edge.Right],
    ['Bottom', Edge.Bottom],
    ['Left', Edge.Left],
  ] as const
  const record = style as Record<string, unknown>
  for (const [suffix, edge] of edges) {
    const padding = px(record[`padding${suffix}`])
    if (padding !== undefined && padding !== 'auto') node.setPadding(edge, padding)
    const margin = px(record[`margin${suffix}`])
    if (margin !== undefined) node.setMargin(edge, margin)
  }
  const insets = [
    ['top', Edge.Top],
    ['right', Edge.Right],
    ['bottom', Edge.Bottom],
    ['left', Edge.Left],
  ] as const
  for (const [key, edge] of insets) {
    const value = px(style[key])
    if (value !== undefined && value !== 'auto') node.setPosition(edge, value)
  }
}

/** The size a leaf takes when neither its props nor its style set one, in inches. */
const intrinsic = (node: TreeNode, width: number | undefined): { width: number; height: number } => {
  switch (node.type) {
    case TYPES.text:
    case TYPES.shape: {
      const runs = runsOf(node)
      if (runs.length === 0) return { width: 1, height: 1 }
      return measureText(runs, textOptions(node.props), width)
    }
    case TYPES.table: {
      const rows = Array.isArray(node.props.rows)
        ? node.props.rows.length
        : node.children.filter((child) => typeof child !== 'string').length
      const fontSize = typeof node.props.fontSize === 'number' ? node.props.fontSize : 12
      const rowH = toInches(node.props.rowH)
      return { width: Math.min(width ?? 6, 6), height: rows * (rowH ?? (fontSize / 72) * 1.2 + 0.1) }
    }
    case TYPES.chart:
      return { width: Math.min(width ?? 4, 4), height: 3 }
    case TYPES.media:
      return { width: Math.min(width ?? 4, 4), height: 2.25 }
    default:
      return { width: 1, height: 1 }
  }
}

interface Entry {
  tree: TreeNode
  node: YogaNode
  children: Entry[]
}

const styleOf = (tree: TreeNode): FlexStyle => {
  const style = { ...((tree.props.style as FlexStyle | undefined) ?? {}) }
  // `w`/`h` work in flex too, and `x`/`y` place an absolutely positioned child.
  if (style.width === undefined && tree.props.w !== undefined) style.width = tree.props.w as number
  if (style.height === undefined && tree.props.h !== undefined) style.height = tree.props.h as number
  if (style.position === 'absolute') {
    if (style.left === undefined && tree.props.x !== undefined) style.left = tree.props.x as number
    if (style.top === undefined && tree.props.y !== undefined) style.top = tree.props.y as number
  }
  return style
}

const build = (Y: Yoga, config: Config, tree: TreeNode, isContainer: boolean): Entry => {
  const node = Y.Node.create(config)
  const style = styleOf(tree)
  applyStyle(Y, node, style)
  const entry: Entry = { tree, node, children: [] }
  if (isContainer) {
    for (const child of tree.children) {
      if (typeof child === 'string') continue
      if (child.type !== TYPES.view && !DRAWABLE.has(child.type)) continue
      const childEntry = build(Y, config, child, child.type === TYPES.view)
      node.insertChild(childEntry.node, entry.children.length)
      entry.children.push(childEntry)
    }
  } else if (style.width === undefined || style.height === undefined) {
    const { MeasureMode } = Y.enums
    const measure: MeasureFunction = (width, widthMode) => {
      const offered = widthMode === MeasureMode.Undefined ? undefined : width / PX_PER_IN
      const size = intrinsic(tree, offered)
      const w = widthMode === MeasureMode.Exactly ? width : size.width * PX_PER_IN
      return {
        width: widthMode === MeasureMode.AtMost ? Math.min(w, width) : w,
        height: size.height * PX_PER_IN,
      }
    }
    node.setMeasureFunc(measure)
  }
  return entry
}

const emit = (entry: Entry, originX: number, originY: number, out: SlideObject[], isRoot: boolean): void => {
  const layout = entry.node.getComputedLayout()
  const x = isRoot ? originX : originX + layout.left / PX_PER_IN
  const y = isRoot ? originY : originY + layout.top / PX_PER_IN
  const frame: Frame = { x, y, w: layout.width / PX_PER_IN, h: layout.height / PX_PER_IN }
  if ((entry.tree.props.style as FlexStyle | undefined)?.display === 'none') return
  if (isRoot || entry.tree.type === TYPES.view) {
    const background = entry.tree.type === TYPES.view ? viewBackground(entry.tree, frame) : undefined
    if (background) out.push(background)
    for (const child of entry.children) emit(child, x, y, out, false)
    return
  }
  const object = toObject(entry.tree, frame)
  if (object) out.push(object)
}

/**
 * Lays out `container`'s children with Yoga inside `frame` (inches) and appends the
 * resulting objects in document order.
 */
export const layoutFlex = async (container: TreeNode, frame: Frame, out: SlideObject[]): Promise<void> => {
  const Y = await loadYoga()
  const config = Y.Config.create()
  // Keep fractional pixels: slides are measured in inches, and rounding would move objects.
  config.setPointScaleFactor(0)
  const root = build(Y, config, { ...container, props: { ...container.props, x: undefined, y: undefined } }, true)
  root.node.setWidth(frame.w * PX_PER_IN)
  root.node.setHeight(frame.h * PX_PER_IN)
  const { Direction } = Y.enums
  try {
    root.node.calculateLayout(frame.w * PX_PER_IN, frame.h * PX_PER_IN, Direction.LTR)
    emit(root, frame.x, frame.y, out, true)
  } finally {
    root.node.freeRecursive()
    config.free()
  }
}
