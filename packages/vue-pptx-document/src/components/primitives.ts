import { defineComponent, h, type DefineComponent, type SlotsType, type VNode } from 'vue'

import { TYPES } from '../core/nodes'
import type {
  ChartProps,
  ImageProps,
  MasterProps,
  MediaProps,
  ParagraphProps,
  PlaceholderProps,
  PresentationProps,
  RunProps,
  ShapeProps,
  SlideProps,
  TableCellProps,
  TableProps,
  TextProps,
  ViewProps,
} from './types'

/** Boolean props a template may write as bare attributes (`<Text bold>`), so Vue casts them. */
const FLAGS = [
  'bold',
  'italic',
  'underline',
  'strike',
  'subscript',
  'superscript',
  'hidden',
  'rounding',
  'flipH',
  'flipV',
  'wrap',
  'rtl',
  'rtlMode',
  'showTitle',
  'showLegend',
  'showValue',
  'showPercent',
  'showLabel',
  'autoPage',
  'autoPageRepeatHeader',
  'slideNumber',
  'bullet',
  'showValAxisTitle',
  'showCatAxisTitle',
  'catAxisHidden',
  'valAxisHidden',
  'showLeaderLines',
  'showSerName',
  'showDataTable',
] as const

interface DefaultSlot {
  default?: () => VNode[]
}

// Each primitive renders an element of its type. Props go through as attributes, except the
// flags it declares, which Vue casts so `<Text bold>` means `bold: true`.
const primitive = <Props extends object>(name: string, type: string, flags: readonly string[] = []) =>
  defineComponent(
    (props: Record<string, unknown>, { attrs, slots }) =>
      () => {
        const set = Object.fromEntries(Object.entries(props).filter(([, value]) => value !== undefined))
        return h(type, { ...attrs, ...set }, slots.default?.())
      },
    {
      name,
      inheritAttrs: false,
      // Boolean first, so a bare attribute is true; `strike` also takes 'sngStrike' and `bullet` an object.
      props: Object.fromEntries(
        flags.map((flag) => [flag, { type: [Boolean, String, Object], default: undefined }]),
      ) as never,
      slots: Object as SlotsType<DefaultSlot>,
    },
  ) as unknown as DefineComponent<Props>

const RUN_FLAGS = ['bold', 'italic', 'underline', 'strike', 'subscript', 'superscript', 'rtlMode'] as const
const TEXT_FLAGS = [...RUN_FLAGS, 'bullet'] as const

export const Presentation = primitive<PresentationProps>('Presentation', TYPES.presentation, ['rtl'])
export const Master = primitive<MasterProps>('Master', TYPES.master, ['slideNumber'])
export const Placeholder = primitive<PlaceholderProps>('Placeholder', TYPES.placeholder, [...TEXT_FLAGS, 'wrap'])
export const Slide = primitive<SlideProps>('Slide', TYPES.slide, ['hidden', 'slideNumber'])
export const View = primitive<ViewProps>('View', TYPES.view)
export const Text = primitive<TextProps>('Text', TYPES.text, [...TEXT_FLAGS, 'wrap'])
export const Paragraph = primitive<ParagraphProps>('Paragraph', TYPES.paragraph, TEXT_FLAGS)
export const Run = primitive<RunProps>('Run', TYPES.run, RUN_FLAGS)
export const Br = primitive<Record<string, never>>('Br', TYPES.br)
export const Shape = primitive<ShapeProps>('Shape', TYPES.shape, [...TEXT_FLAGS, 'flipH', 'flipV'])
export const Image = primitive<ImageProps>('Image', TYPES.image, ['rounding', 'flipH', 'flipV'])
export const Chart = primitive<ChartProps>('Chart', TYPES.chart, [
  'showTitle',
  'showLegend',
  'showValue',
  'showPercent',
  'showLabel',
  'showValAxisTitle',
  'showCatAxisTitle',
  'catAxisHidden',
  'valAxisHidden',
  'showLeaderLines',
  'showSerName',
  'showDataTable',
])
export const Table = primitive<TableProps>('Table', TYPES.table, ['autoPage', 'autoPageRepeatHeader', ...RUN_FLAGS])
export const TableRow = primitive<TableCellProps>('TableRow', TYPES.row, RUN_FLAGS)
export const TableCell = primitive<TableCellProps>('TableCell', TYPES.cell, TEXT_FLAGS)
export const Media = primitive<MediaProps>('Media', TYPES.media)
export const Notes = primitive<Record<string, never>>('Notes', TYPES.notes)

export { FLAGS }
