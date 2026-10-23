import type PptxGenJS from 'pptxgenjs'

import type { FlexStyle } from '../core/flex'
import type { ChartName, ChartSeries, ComboChart } from '../core/model'
import type { LayoutName } from '../core/resolve'
import type { Length } from '../core/units'

export type { FlexStyle, Length }

/** A color as CSS writes it (`#f35815`, `#fff`, `white`) or a theme slot (`accent1`). */
export type Color = string

/** Where an object sits. Numbers are inches; strings take a unit or a percentage of the parent. */
export interface PositionProps {
  x?: Length
  y?: Length
  w?: Length
  h?: Length
  /** Flexbox style, read when the object is inside a `<View>` or a `<Slide layout="flex">`. */
  style?: FlexStyle
}

export interface RunProps {
  bold?: boolean
  italic?: boolean
  underline?: boolean | PptxGenJS.TextBaseProps['underline']
  strike?: boolean | 'dblStrike' | 'sngStrike'
  subscript?: boolean
  superscript?: boolean
  color?: Color
  highlight?: Color
  fontFace?: string
  /** Points. */
  fontSize?: number
  charSpacing?: number
  transparency?: number
  /** A shorthand for `hyperlink: { url }`. */
  href?: string
  hyperlink?: PptxGenJS.HyperlinkProps
  lang?: string
}

export interface ParagraphProps extends RunProps {
  align?: PptxGenJS.HAlign
  bullet?: PptxGenJS.TextBaseProps['bullet']
  indentLevel?: number
  /** Points. */
  paraSpaceBefore?: number
  /** Points. */
  paraSpaceAfter?: number
  /** Points. */
  lineSpacing?: number
  lineSpacingMultiple?: number
  rtlMode?: boolean
}

export type FillProp = Color | PptxGenJS.ShapeFillProps
export type LineProp = Color | PptxGenJS.ShapeLineProps

export interface TextProps extends PositionProps, ParagraphProps {
  valign?: PptxGenJS.VAlign
  fill?: FillProp
  line?: LineProp
  /** Inset in points: one number, or `[top, right, bottom, left]` as in CSS. */
  margin?: number | [number, number, number, number]
  fit?: 'none' | 'shrink' | 'resize'
  wrap?: boolean
  rotate?: number
  shadow?: PptxGenJS.ShadowProps
  glow?: PptxGenJS.TextGlowProps
  outline?: { color: Color; size: number }
  shape?: PptxGenJS.SHAPE_NAME
  rectRadius?: number
  vert?: PptxGenJS.TextPropsOptions['vert']
  /** Fills a master's placeholder by name. */
  placeholder?: string
  /** Any other PptxGenJS text option. */
  [option: string]: unknown
}

export interface ShapeProps extends PositionProps {
  /** A PptxGenJS shape name: `rect`, `roundRect`, `ellipse`, `line`, `triangle`, `rightArrow`, ... */
  type?: PptxGenJS.SHAPE_NAME | (string & {})
  fill?: FillProp
  line?: LineProp
  rectRadius?: number
  rotate?: number
  flipH?: boolean
  flipV?: boolean
  shadow?: PptxGenJS.ShadowProps
  hyperlink?: PptxGenJS.HyperlinkProps
  [option: string]: unknown
}

export interface ImageProps extends PositionProps {
  /** A URL, a path (in Node) or a data URI. */
  src: string
  altText?: string
  sizing?: PptxGenJS.ImageProps['sizing']
  rounding?: boolean
  rotate?: number
  flipH?: boolean
  flipV?: boolean
  transparency?: number
  hyperlink?: PptxGenJS.HyperlinkProps
  [option: string]: unknown
}

export type TableCellValue = string | number | boolean | null | undefined | PptxGenJS.TableCell

export interface TableProps extends PositionProps {
  rows?: TableCellValue[][]
  /** Column widths: one for all, or one per column. */
  colW?: Length | Length[]
  rowH?: Length | Length[]
  fontSize?: number
  fontFace?: string
  color?: Color
  fill?: FillProp
  border?:
    PptxGenJS.BorderProps | [PptxGenJS.BorderProps, PptxGenJS.BorderProps, PptxGenJS.BorderProps, PptxGenJS.BorderProps]
  align?: PptxGenJS.HAlign
  valign?: PptxGenJS.VAlign
  /** Cell margin in inches, as PptxGenJS reads it for tables. */
  margin?: number | [number, number, number, number]
  autoPage?: boolean
  autoPageRepeatHeader?: boolean
  autoPageHeaderRows?: number
  [option: string]: unknown
}

export interface TableCellProps extends ParagraphProps {
  colspan?: number
  rowspan?: number
  fill?: FillProp
  border?: TableProps['border']
  valign?: PptxGenJS.VAlign
  margin?: number | [number, number, number, number]
  [option: string]: unknown
}

export type ChartType = ChartName | Lowercase<ChartName>

export interface ChartProps extends PositionProps {
  /** A chart type, or several `{ type, data, options }` for a combo chart. */
  type: ChartType | ComboChart[]
  data?: ChartSeries[]
  title?: string
  showTitle?: boolean
  showLegend?: boolean
  legendPos?: 'b' | 'l' | 'r' | 't' | 'tr'
  showValue?: boolean
  showPercent?: boolean
  chartColors?: Color[]
  barDir?: 'bar' | 'col'
  barGrouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard'
  /** Any other PptxGenJS chart option. */
  [option: string]: unknown
}

export interface MediaProps extends PositionProps {
  type?: PptxGenJS.MediaType
  /** A URL or data URI; for `online`, the video link (YouTube embed URL). */
  src?: string
  /** A poster image as a data URI. */
  cover?: string
  extn?: string
  [option: string]: unknown
}

export interface ViewProps extends PositionProps {
  /** Draws a rectangle behind the children. */
  fill?: FillProp
  line?: LineProp
  rectRadius?: number
  shadow?: PptxGenJS.ShadowProps
  /** Rotation of the background rectangle, in degrees. */
  rotate?: number
}

export type BackgroundProp = Color | { color?: Color; transparency?: number; src?: string }

export interface SlideNumberOptions extends PositionProps, RunProps {
  align?: PptxGenJS.HAlign
}

export interface SlideProps {
  background?: BackgroundProp
  /** Default text color. */
  color?: Color
  hidden?: boolean
  /** The `name` of a `<Master>`. */
  master?: string
  /** Starts or continues a section with this title. */
  section?: string
  /** Speaker notes. `<Notes>` children work too. */
  notes?: string
  slideNumber?: boolean | SlideNumberOptions
  /** `flex` lays the slide's children out with Yoga, as a column by default. */
  layout?: 'absolute' | 'flex'
  style?: FlexStyle
}

export interface MasterProps {
  name: string
  background?: BackgroundProp
  margin?: PptxGenJS.Margin
  slideNumber?: boolean | SlideNumberOptions
  layout?: 'absolute' | 'flex'
  style?: FlexStyle
}

export interface PlaceholderProps extends TextProps {
  name: string
  type?: PptxGenJS.PLACEHOLDER_TYPE
}

export interface PresentationProps {
  /** `16x9` (10 x 5.625 in, the default), `16x10`, `4x3`, `wide`, or a custom size. */
  layout?: LayoutName | { width: Length; height: Length }
  title?: string
  author?: string
  company?: string
  subject?: string
  revision?: string | number
  rtl?: boolean
  headFontFace?: string
  bodyFontFace?: string
  theme?: PptxGenJS.ThemeProps
  /** Called with the Blob after each `toBlob()` render. */
  onRender?: (args: { blob: Blob }) => void
}
