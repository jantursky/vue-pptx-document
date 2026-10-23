/** Element types the primitives render. Each one becomes a node of the same type in the tree. */
export const TYPES = {
  presentation: 'PRESENTATION',
  master: 'MASTER',
  placeholder: 'PLACEHOLDER',
  slide: 'SLIDE',
  view: 'VIEW',
  text: 'TEXT',
  paragraph: 'PARAGRAPH',
  run: 'RUN',
  br: 'BR',
  shape: 'SHAPE',
  image: 'IMAGE',
  chart: 'CHART',
  table: 'TABLE',
  row: 'TABLE_ROW',
  cell: 'TABLE_CELL',
  media: 'MEDIA',
  notes: 'NOTES',
} as const

export type NodeType = (typeof TYPES)[keyof typeof TYPES]

export const ROOT = 'ROOT'

/** Elements whose children may be plain text. Text anywhere else is dropped with a warning. */
export const TEXT_PARENTS = new Set<string>([
  TYPES.text,
  TYPES.paragraph,
  TYPES.run,
  TYPES.shape,
  TYPES.cell,
  TYPES.notes,
  TYPES.placeholder,
])

export interface PptxElement {
  kind: 'element'
  type: string
  props: Record<string, unknown>
  children: PptxNode[]
  parent: PptxElement | null
}

export interface PptxTextNode {
  kind: 'text'
  value: string
  parent: PptxElement | null
}

export interface PptxCommentNode {
  kind: 'comment'
  value: string
  parent: PptxElement | null
}

export type PptxNode = PptxElement | PptxTextNode | PptxCommentNode

export interface PptxRoot extends PptxElement {
  type: typeof ROOT
  onChange: (() => void) | null
}

export const createElementNode = (type: string): PptxElement => ({
  kind: 'element',
  type,
  props: {},
  children: [],
  parent: null,
})

export const createRoot = (): PptxRoot => ({ ...createElementNode(ROOT), type: ROOT, onChange: null })
