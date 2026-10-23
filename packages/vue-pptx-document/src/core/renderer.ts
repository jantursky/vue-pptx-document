import { camelize, createRenderer, type RendererOptions } from 'vue'

import { createElementNode, ROOT, type PptxElement, type PptxNode, type PptxRoot } from './nodes'

/** Walks up to the mount container and tells it the tree changed. */
const notify = (node: PptxNode | null): void => {
  let current = node
  while (current?.parent) current = current.parent
  if (current?.kind === 'element' && current.type === ROOT) (current as PptxRoot).onChange?.()
}

/** Template attributes arrive as written (`flex-direction`); the layout reads camelCase. */
const camelizeKeys = (style: unknown): unknown => {
  if (style == null || typeof style !== 'object' || Array.isArray(style)) return style
  return Object.fromEntries(Object.entries(style).map(([key, value]) => [camelize(key), value]))
}

const detach = (child: PptxNode): void => {
  const { parent } = child
  if (!parent) return
  const index = parent.children.indexOf(child)
  if (index !== -1) parent.children.splice(index, 1)
  child.parent = null
}

export const nodeOps: RendererOptions<PptxNode, PptxElement> = {
  createElement: (type) => createElementNode(type),
  createText: (value) => ({ kind: 'text', value, parent: null }),
  createComment: (value) => ({ kind: 'comment', value, parent: null }),
  setText(node, value) {
    if (node.kind === 'element') return
    node.value = value
    notify(node)
  },
  setElementText(element, value) {
    for (const child of element.children) child.parent = null
    element.children = value === '' ? [] : [{ kind: 'text', value, parent: element }]
    notify(element)
  },
  insert(child, parent, anchor) {
    detach(child)
    const index = anchor ? parent.children.indexOf(anchor) : -1
    if (index === -1) parent.children.push(child)
    else parent.children.splice(index, 0, child)
    child.parent = parent
    notify(parent)
  },
  remove(child) {
    const { parent } = child
    detach(child)
    notify(parent)
  },
  patchProp(element, key, _prev, next) {
    // `class` has no meaning in a slide; Vue normalizes it onto every vnode anyway.
    if (key === 'class') return
    const name = camelize(key)
    if (next === undefined) Reflect.deleteProperty(element.props, name)
    else element.props[name] = key === 'style' ? camelizeKeys(next) : next
    notify(element)
  },
  parentNode: (node) => node.parent,
  nextSibling(node) {
    const siblings = node.parent?.children
    if (!siblings) return null
    return siblings[siblings.indexOf(node) + 1] ?? null
  },
}

export const renderer = createRenderer<PptxNode, PptxElement>(nodeOps)
