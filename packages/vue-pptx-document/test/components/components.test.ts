import { mount } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createSSRApp, defineComponent, h, nextTick } from 'vue'
import { renderToString } from 'vue/server-renderer'

import {
  Chart,
  Notes,
  Preview,
  PptxDownloadButton,
  Shape,
  Slide,
  SlidePreview,
  Table,
  Text,
  renderToModel,
} from '../../src'
import { singleSlide } from '../fixtures'

const Deck = singleSlide(() => [
  h(Text, { x: 1, y: 1, w: 4, h: 1, fontSize: 24, color: '#f35815' }, () => 'Hello'),
  h(Shape, { type: 'ellipse', x: 5, y: 1, w: 2, h: 2, fill: '#111' }),
  h(Table, { x: 1, y: 3, w: 4, rows: [['A', 'B']] }),
  h(Chart, { type: 'pie', x: 6, y: 3, w: 3, h: 2, data: [{ name: 'S', labels: ['a', 'b'], values: [1, 3] }] }),
  h(Notes, null, () => 'Remember this'),
])
const wrappers: { unmount: () => void }[] = []

const until = (condition: () => boolean) =>
  vi.waitFor(() => {
    if (!condition()) throw new Error('not yet')
  })

afterEach(() => {
  wrappers.splice(0).forEach((wrapper) => wrapper.unmount())
})

describe('SlidePreview', () => {
  it('draws a slide with text, shapes, tables and charts', async () => {
    const model = await renderToModel(Deck)
    const wrapper = mount(SlidePreview, { props: { model, index: 0 } })
    wrappers.push(wrapper)
    const slide = wrapper.find('[data-pptx-slide]')
    expect(slide.attributes('role')).toBe('group')
    expect(slide.attributes('aria-roledescription')).toBe('slide')
    expect(slide.attributes('aria-label')).toBe('Slide 1')
    expect(slide.attributes('style')).toContain('aspect-ratio: 10 / 5.625')
    expect(wrapper.text()).toContain('Hello')
    expect(wrapper.find('table').text()).toContain('A')
    expect(wrapper.findAll('svg').length).toBeGreaterThanOrEqual(2)
    expect(wrapper.html()).toMatch(/color: (#F35815|rgb\(243, 88, 21\))/i)
  })

  it('renders nothing for a missing slide', async () => {
    const wrapper = mount(SlidePreview, { props: { model: await renderToModel(Deck), index: 4 } })
    wrappers.push(wrapper)
    expect(wrapper.find('[data-pptx-slide]').exists()).toBe(false)
  })
})

describe('Preview', () => {
  it('shows the loading slot, then every slide with notes', async () => {
    const wrapper = mount(Preview, { props: { document: Deck, notes: true }, slots: { loading: 'Rendering…' } })
    wrappers.push(wrapper)
    expect(wrapper.text()).toBe('Rendering…')
    await until(() => wrapper.find('[data-pptx-preview-slide]').exists())
    expect(wrapper.find('[data-pptx-preview-caption]').text()).toContain('1 / 1')
    expect(wrapper.find('[data-pptx-preview-notes]').text()).toBe('Remember this')
    expect(wrapper.attributes('aria-label')).toBe('Presentation preview')
  })

  it('follows document props and the caption slot', async () => {
    const Titled = defineComponent({
      props: { title: { type: String, required: true } },
      setup: (props) => () => h(singleSlide(() => h(Text, { w: 5, h: 1 }, () => props.title))),
    })
    const wrapper = mount(Preview, {
      props: { document: Titled, documentProps: { title: 'First' }, debounce: 0 },
      slots: { caption: ({ index, total }: { index: number; total: number }) => `Slide ${index + 1} of ${total}` },
    })
    wrappers.push(wrapper)
    await until(() => wrapper.text().includes('First'))
    expect(wrapper.text()).toContain('Slide 1 of 1')
    await wrapper.setProps({ documentProps: { title: 'Second' } })
    await until(() => wrapper.text().includes('Second'))
  })

  it('shows the error slot when the tree is wrong', async () => {
    const wrapper = mount(Preview, { props: { document: defineComponent(() => () => h(Slide)) } })
    wrappers.push(wrapper)
    await until(() => wrapper.find('[role="alert"]').exists())
    expect(wrapper.text()).toContain('<Presentation>')
  })

  it('renders on the server without PptxGenJS', async () => {
    const html = await renderToString(
      createSSRApp({ render: () => h(Preview, { document: Deck }, { loading: () => 'Loading' }) }),
    )
    expect(html).toContain('data-pptx-preview')
    expect(html).toContain('Loading')
  })
})

describe('PptxDownloadButton', () => {
  it('renders and saves the file on click', async () => {
    const click = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})
    const wrapper = mount(PptxDownloadButton, { props: { document: Deck, fileName: 'deck.pptx' } })
    wrappers.push(wrapper)
    expect(wrapper.text()).toBe('Download PPTX')
    await wrapper.trigger('click')
    await nextTick()
    expect(wrapper.attributes('aria-busy')).toBe('true')
    await until(() => wrapper.emitted('done') !== undefined)
    expect((wrapper.emitted('done')![0]![0] as Blob).size).toBeGreaterThan(1000)
    expect(click).toHaveBeenCalledOnce()
    click.mockRestore()
  })

  it('emits errors and passes state to the slot', async () => {
    const wrapper = mount(PptxDownloadButton, {
      props: { document: defineComponent(() => () => h(Slide)) },
      slots: { default: ({ error }: { error: unknown }) => (error ? 'Failed' : 'Get it') },
    })
    wrappers.push(wrapper)
    expect(wrapper.text()).toBe('Get it')
    await wrapper.trigger('click')
    await until(() => wrapper.emitted('error') !== undefined)
    expect(wrapper.text()).toBe('Failed')
  })
})
