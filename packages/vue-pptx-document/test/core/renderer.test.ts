import { describe, expect, it, vi } from 'vitest'
import { defineComponent, h, nextTick, ref } from 'vue'

import { Presentation, Slide, Text } from '../../src'
import { createElementNode, createRoot } from '../../src/core/nodes'
import { nodeOps, renderer } from '../../src/core/renderer'
import { snapshot } from '../../src/core/snapshot'

const mount = (component: Parameters<typeof renderer.createApp>[0]) => {
  const root = createRoot()
  const app = renderer.createApp(component)
  app.mount(root)
  return { root, app }
}

describe('node operations', () => {
  it('inserts before an anchor and moves nodes between parents', () => {
    const a = createElementNode('TEXT')
    const b = createElementNode('TEXT')
    const parent = createElementNode('SLIDE')
    const other = createElementNode('SLIDE')
    nodeOps.insert(a, parent)
    nodeOps.insert(b, parent, a)
    expect(parent.children).toEqual([b, a])
    expect(nodeOps.nextSibling(b)).toBe(a)
    expect(nodeOps.nextSibling(a)).toBeNull()
    nodeOps.insert(b, other)
    expect(parent.children).toEqual([a])
    expect(nodeOps.parentNode(b)).toBe(other)
  })

  it('removes nodes and replaces element text', () => {
    const text = createElementNode('TEXT')
    nodeOps.insert(nodeOps.createText('old'), text)
    nodeOps.setElementText(text, 'new')
    expect(text.children).toMatchObject([{ kind: 'text', value: 'new' }])
    const child = text.children[0]!
    nodeOps.remove(child)
    nodeOps.remove(child)
    expect(text.children).toEqual([])
  })

  it('camelizes props and style keys and drops class', () => {
    const element = createElementNode('TEXT')
    nodeOps.patchProp(element, 'font-size', undefined, 12)
    nodeOps.patchProp(element, 'style', undefined, { 'flex-grow': 1 })
    nodeOps.patchProp(element, 'class', undefined, 'x')
    expect(element.props).toEqual({ fontSize: 12, style: { flexGrow: 1 } })
    nodeOps.patchProp(element, 'font-size', 12, undefined)
    expect(element.props).toEqual({ style: { flexGrow: 1 } })
  })
})

describe('snapshot', () => {
  it('copies the tree and casts bare boolean flags', () => {
    const { root } = mount(
      defineComponent(
        () => () =>
          h(Presentation, { title: 'T' }, () => h(Slide, null, () => h(Text, { bold: '', x: 1 }, () => 'Hi'))),
      ),
    )
    expect(snapshot(root)).toEqual({
      type: 'PRESENTATION',
      props: { title: 'T' },
      children: [
        { type: 'SLIDE', props: {}, children: [{ type: 'TEXT', props: { bold: true, x: 1 }, children: ['Hi'] }] },
      ],
    })
  })

  it('requires a Presentation at the root', () => {
    const { root } = mount(defineComponent(() => () => h(Slide)))
    expect(() => snapshot(root)).toThrow('must render a single <Presentation>')
  })

  it('drops text outside a text parent with a warning', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const { root } = mount(defineComponent(() => () => h(Presentation, null, () => h(Slide, null, () => 'loose text'))))
    expect(snapshot(root).children[0]).toMatchObject({ children: [] })
    expect(warn).toHaveBeenCalledWith(expect.stringContaining('loose text'))
    warn.mockRestore()
  })

  it('notifies the root on reactive changes', async () => {
    const label = ref('a')
    const { root } = mount(
      defineComponent(() => () => h(Presentation, null, () => h(Slide, null, () => h(Text, null, () => label.value)))),
    )
    const onChange = vi.fn()
    root.onChange = onChange
    label.value = 'b'
    await nextTick()
    expect(onChange).toHaveBeenCalled()
    expect(JSON.stringify(snapshot(root))).toContain('"b"')
  })
})
