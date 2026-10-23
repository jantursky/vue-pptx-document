import { TEXT_PARENTS, TYPES, type PptxElement, type PptxNode } from './nodes'

/** A plain copy of the live tree: no parents, no comments, text as strings. */
export interface TreeNode {
  type: string
  props: Record<string, unknown>
  children: (TreeNode | string)[]
}

const warned = new Set<string>()
const warnOnce = (message: string): void => {
  if (warned.has(message)) return
  warned.add(message)
  console.warn(`[vue-pptx-document] ${message}`)
}

const convert = (element: PptxElement): TreeNode => ({
  type: element.type,
  props: { ...element.props },
  children: element.children.flatMap((child: PptxNode): (TreeNode | string)[] => {
    if (child.kind === 'comment') return []
    if (child.kind === 'element') return [convert(child)]
    if (child.value.trim() === '') return []
    if (!TEXT_PARENTS.has(element.type)) {
      warnOnce(`Text "${child.value.trim().slice(0, 40)}" inside <${element.type}> is dropped. Wrap it in <Text>.`)
      return []
    }
    return [child.value]
  }),
})

/** Copies the tree under the mount container. The root component must render one `<Presentation>`. */
export const snapshot = (root: PptxElement): TreeNode => {
  const first = root.children.find((child): child is PptxElement => child.kind === 'element')
  if (first?.type !== TYPES.presentation) {
    throw new Error('[vue-pptx-document] The root component must render a single <Presentation>.')
  }
  return convert(first)
}
