import type PptxGenJS from 'pptxgenjs'

/**
 * The resolved presentation: every object has its final frame in inches and options in
 * PptxGenJS's own shape. The writer hands it to PptxGenJS, the preview draws it as HTML,
 * and the importer produces one from a .pptx file.
 */
export interface Frame {
  x: number
  y: number
  w: number
  h: number
}

export interface TextRun {
  text: string
  options?: PptxGenJS.TextPropsOptions
}

export type ChartName =
  'area' | 'bar' | 'bar3D' | 'bubble' | 'bubble3D' | 'doughnut' | 'line' | 'pie' | 'radar' | 'scatter'

export interface ChartSeries {
  name?: string
  labels?: string[]
  values?: number[]
  sizes?: number[]
}

export interface ComboChart {
  type: ChartName
  data: ChartSeries[]
  options?: PptxGenJS.IChartOpts
}

export interface TableCellModel {
  text?: string | TextRun[]
  options?: PptxGenJS.TableCellProps
}

export type SlideObject =
  | { kind: 'text'; frame: Frame; runs: TextRun[]; options: PptxGenJS.TextPropsOptions }
  | { kind: 'shape'; frame: Frame; shape: string; options: PptxGenJS.ShapeProps }
  | { kind: 'image'; frame: Frame; src: string; options: PptxGenJS.ImageProps }
  | { kind: 'table'; frame: Frame; rows: TableCellModel[][]; options: PptxGenJS.TableProps }
  | {
      kind: 'chart'
      frame: Frame
      type: ChartName | ComboChart[]
      data: ChartSeries[]
      options: PptxGenJS.IChartOpts
    }
  | { kind: 'media'; frame: Frame; options: PptxGenJS.MediaProps }
  | { kind: 'placeholder'; frame: Frame; name: string; text?: string; options: PptxGenJS.PlaceholderProps }

export type SlideObjectKind = SlideObject['kind']

export interface Background {
  color?: string
  transparency?: number
  /** An image URL or data URI. */
  src?: string
}

export interface SlideModel {
  background?: Background
  /** Default text color. */
  color?: string
  hidden?: boolean
  master?: string
  section?: string
  notes?: string
  slideNumber?: PptxGenJS.SlideNumberProps
  objects: SlideObject[]
}

export interface MasterModel {
  name: string
  background?: Background
  margin?: PptxGenJS.Margin
  slideNumber?: PptxGenJS.SlideNumberProps
  objects: SlideObject[]
}

export interface DeckLayout {
  /** PptxGenJS layout name, such as LAYOUT_16x9, or CUSTOM for a defined size. */
  name: string
  width: number
  height: number
}

export interface DeckMeta {
  title?: string
  author?: string
  company?: string
  subject?: string
  revision?: string
}

export interface DeckModel {
  layout: DeckLayout
  meta: DeckMeta
  rtl?: boolean
  theme?: PptxGenJS.ThemeProps
  masters: MasterModel[]
  slides: SlideModel[]
}
