import { checkProps } from './check-props'
import { at, warn } from './diagnostics'
import { tag, TEXT_PARENTS, TYPES, type PptxElement, type PptxNode } from './nodes'

/** A plain copy of the live tree: no parents, no comments, text as strings. */
export interface TreeNode {
  type: string
  props: Record<string, unknown>
  children: (TreeNode | string)[]
}

const convert = (element: PptxElement): TreeNode => {
  if (process.env.NODE_ENV !== 'production') {
    at({ object: tag(element.type) }, () => checkProps(element.type, element.props))
  }
  return copy(element)
}

const copy = (element: PptxElement): TreeNode => ({
  type: element.type,
  props: { ...element.props },
  children: element.children.flatMap((child: PptxNode): (TreeNode | string)[] => {
    if (child.kind === 'comment') return []
    if (child.kind === 'element') return [convert(child)]
    if (child.value.trim() === '') return []
    if (!TEXT_PARENTS.has(element.type)) {
      at({ object: tag(element.type) }, () =>
        warn(
          'dropped-text',
          `Text "${child.value.trim().slice(0, 40)}" inside ${tag(element.type)} is dropped. Wrap it in <Text>.`,
        ),
      )
      return []
    }
    return [child.value]
  }),
})

/** Slides are numbered as `resolve` numbers them: direct `<Slide>` children of the presentation. */
const convertDeck = (presentation: PptxElement): TreeNode => {
  let slide = 0
  return {
    ...convert({ ...presentation, children: [] }),
    children: presentation.children.flatMap((child) => {
      if (child.kind !== 'element' || child.type !== TYPES.slide)
        return copy({ ...presentation, children: [child] }).children
      slide += 1
      return [at({ slide }, () => convert(child))]
    }),
  }
}

/** Copies the tree under the mount container. The root component must render one `<Presentation>`. */
export const snapshot = (root: PptxElement): TreeNode => {
  const first = root.children.find((child): child is PptxElement => child.kind === 'element')
  if (first?.type !== TYPES.presentation) {
    throw new Error('[vue-pptx-document] The root component must render a single <Presentation>.')
  }
  return convertDeck(first)
}
