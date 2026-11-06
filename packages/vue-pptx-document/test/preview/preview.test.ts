import { mount } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { defineComponent, h } from 'vue'

import {
  Chart,
  Image,
  Paragraph,
  Preview,
  Presentation,
  Shape,
  Slide,
  SlidePreview,
  Table,
  Text,
  renderToModel,
} from '../../src'
import type { ChartName } from '../../src'
import { formatNumber, niceScale } from '../../src/preview/chart'
import { isKnownShape, shapePath } from '../../src/preview/shapes'
import { singleSlide } from '../fixtures'

const draw = async (children: () => unknown) => {
  const model = await renderToModel(singleSlide(children as never))
  return mount(SlidePreview, { props: { model, index: 0 } })
}

const series = [
  { name: 'North', labels: ['Q1', 'Q2', 'Q3'], values: [3, 5, 4] },
  { name: 'South', labels: ['Q1', 'Q2', 'Q3'], values: [2, 1, 6] },
]

describe('chart helpers', () => {
  it('formats numbers with Excel codes', () => {
    expect(formatNumber(1234.5)).toBe('1234.5')
    expect(formatNumber(1234.5, '#,##0')).toBe('1,235')
    expect(formatNumber(0.256, '0.0%')).toBe('25.6%')
    expect(formatNumber(-3, '$0.00')).toBe('-$3.00')
    expect(formatNumber(Number.NaN)).toBe('')
  })

  it('picks round axis bounds', () => {
    expect(niceScale(0, 87)).toEqual({ min: 0, max: 100, step: 20 })
    expect(niceScale(5, 5)).toEqual({ min: 0, max: 5, step: 1 })
    expect(niceScale(0, 0)).toEqual({ min: 0, max: 1, step: 0.2 })
    expect(niceScale(0, 10, 4).step).toBe(4)
  })
})

describe('shapes', () => {
  it('knows the common presets', () => {
    for (const name of ['rect', 'roundRect', 'ellipse', 'triangle', 'rightArrow', 'star5', 'line'])
      expect(isKnownShape(name)).toBe(true)
    expect(isKnownShape('nope')).toBe(false)
    expect(shapePath('rect', 10, 5)).toMatch(/^M/)
  })

  it('draws unknown presets as rectangles', async () => {
    const wrapper = await draw(() => h(Shape, { type: 'cloudCallout', x: 1, y: 1, w: 2, h: 1, fill: '#f00' }))
    expect(wrapper.find('svg').exists()).toBe(true)
  })
})

describe('text', () => {
  it('draws bullets, numbers and links', async () => {
    const wrapper = await draw(() =>
      h(Text, { x: 0, y: 0, w: 5, h: 2 }, () => [
        h(Paragraph, { bullet: true }, () => 'Dot'),
        h(Paragraph, { bullet: { type: 'number' } }, () => 'First'),
        h(Paragraph, { bullet: { type: 'number', numberType: 'romanUcPeriod' } }, () => 'Second'),
        h(Paragraph, { href: 'https://vuejs.org' }, () => 'Link'),
      ]),
    )
    const paragraphs = wrapper.findAll('p')
    expect(paragraphs.map((p) => p.text())).toEqual(['•Dot', '1.First', 'I.Second', 'Link'])
    expect(wrapper.find('a').attributes()).toMatchObject({ href: 'https://vuejs.org', rel: 'noopener noreferrer' })
  })
})

describe('charts', () => {
  const types: ChartName[] = [
    'bar',
    'bar3D',
    'line',
    'area',
    'pie',
    'doughnut',
    'scatter',
    'bubble',
    'bubble3D',
    'radar',
  ]

  it.each(types)('draws a %s chart', async (type) => {
    const data =
      type === 'scatter'
        ? [
            { name: 'X', values: [1, 2, 3] },
            { name: 'Y', values: [3, 1, 2] },
          ]
        : type.startsWith('bubble')
          ? [
              { name: 'X', values: [1, 2, 3] },
              { name: 'Y', values: [3, 1, 2], sizes: [4, 6, 8] },
            ]
          : series
    const wrapper = await draw(() =>
      h(Chart, { type, x: 0, y: 0, w: 8, h: 4.5, data, showLegend: true, showTitle: true, title: 'Sales' }),
    )
    const svg = wrapper.find('svg')
    expect(svg.exists()).toBe(true)
    expect(svg.text()).toContain('Sales')
    expect(svg.findAll('path, rect, circle, polygon, polyline').length).toBeGreaterThan(2)
  })

  it('draws a combo with a secondary axis', async () => {
    const wrapper = await draw(() =>
      h(Chart, {
        type: [
          { type: 'bar', data: [series[0]!] },
          { type: 'line', data: [series[1]!], options: { secondaryValAxis: true, secondaryCatAxis: true } },
        ],
        x: 0,
        y: 0,
        w: 8,
        h: 4.5,
        valAxes: [
          { showValAxisTitle: true, valAxisTitle: 'Left' },
          { showValAxisTitle: true, valAxisTitle: 'Right' },
        ],
        catAxes: [{ catAxisTitle: 'Q' }, { catAxisHidden: true }],
      }),
    )
    const text = wrapper.find('svg').text()
    expect(text).toContain('Left')
    expect(text).toContain('Right')
  })

  it('runs horizontal bar categories from the bottom, as PowerPoint does', async () => {
    const wrapper = await draw(() =>
      h(Chart, { type: 'bar', barDir: 'bar', x: 0, y: 0, w: 8, h: 4.5, data: [series[0]!] }),
    )
    const labels = wrapper.findAll('svg text').filter((node) => ['Q1', 'Q3'].includes(node.text()))
    const y = (name: string) => Number(labels.find((node) => node.text() === name)!.attributes('y'))
    expect(y('Q1')).toBeGreaterThan(y('Q3'))
  })
})

describe('links', () => {
  it('only makes web, mail and phone links clickable', async () => {
    const { safeUrl } = await import('../../src/preview/render')
    expect(safeUrl('https://vuejs.org')).toBe('https://vuejs.org')
    expect(safeUrl('mailto:ada@example.com')).toBe('mailto:ada@example.com')
    expect(safeUrl('#slide-2')).toBe('#slide-2')
    expect(safeUrl('javascript:alert(1)')).toBeUndefined()
    expect(safeUrl(' java\tscript:alert(1)')).toBeUndefined()
    expect(safeUrl('data:text/html,<script>')).toBeUndefined()
    const wrapper = await draw(() => h(Text, { x: 0, y: 0, w: 5, h: 1, href: 'javascript:alert(1)' }, () => 'Click'))
    expect(wrapper.find('a').exists()).toBe(false)
    expect(wrapper.text()).toContain('Click')
  })
})

describe('large decks', () => {
  afterEach(() => vi.unstubAllGlobals())

  const Many = defineComponent({
    props: { second: { type: String, required: true } },
    setup: (props) => () =>
      h(Presentation, null, () =>
        ['One', props.second, 'Three'].map((text, index) =>
          h(Slide, { key: index }, () => h(Text, { x: 0, y: 0, w: 3, h: 1 }, () => text)),
        ),
      ),
  })

  const until = (condition: () => boolean) =>
    vi.waitFor(() => {
      if (!condition()) throw new Error('not yet')
    })

  it('keeps the DOM of slides that did not change', async () => {
    const wrapper = mount(Preview, {
      props: { document: Many, documentProps: { second: 'Two' }, debounce: 0, lazy: false },
    })
    await until(() => wrapper.findAll('[data-pptx-slide]').length === 3)
    const before = wrapper.findAll('[data-pptx-slide]').map((node) => node.element)
    await wrapper.setProps({ documentProps: { second: 'Changed' } })
    await until(() => wrapper.text().includes('Changed'))
    const after = wrapper.findAll('[data-pptx-slide]').map((node) => node.element)
    expect(after[0]).toBe(before[0])
    expect(after[2]).toBe(before[2])
    expect(wrapper.findAll('[data-pptx-slide]')[1]!.text()).toBe('Changed')
    wrapper.unmount()
  })

  it('draws only slides near the viewport until others scroll in', async () => {
    const watched: { callback: IntersectionObserverCallback; elements: Element[] }[] = []
    vi.stubGlobal(
      'IntersectionObserver',
      class {
        elements: Element[] = []
        constructor(callback: IntersectionObserverCallback) {
          watched.push({ callback, elements: this.elements })
        }
        observe(element: Element) {
          this.elements.push(element)
        }
        unobserve() {}
        disconnect() {}
      },
    )
    const wrapper = mount(Preview, { props: { document: Many, documentProps: { second: 'Two' }, debounce: 0 } })
    await until(() => wrapper.findAll('figure').length === 3)
    // Every figure already has its slide's shape, with nothing drawn in it.
    expect(wrapper.findAll('[data-pptx-slide]')).toHaveLength(0)
    expect(wrapper.findAll('[data-pptx-slide-placeholder]')).toHaveLength(3)
    expect(wrapper.find('figure > div').attributes('style')).toContain('aspect-ratio: 10 / 5.625')
    const [observer] = watched
    const [first] = observer!.elements
    observer!.callback([{ target: first!, isIntersecting: true } as IntersectionObserverEntry], {} as never)
    await until(() => wrapper.findAll('[data-pptx-slide]').length === 1)
    expect(wrapper.find('[data-pptx-slide]').text()).toBe('One')
    expect(wrapper.findAll('[data-pptx-slide-placeholder]')).toHaveLength(2)
    window.dispatchEvent(new Event('beforeprint'))
    await until(() => wrapper.findAll('[data-pptx-slide]').length === 3)
    wrapper.unmount()
  })

  it('draws every slide at once with lazy off', async () => {
    const wrapper = mount(Preview, {
      props: { document: Many, documentProps: { second: 'Two' }, debounce: 0, lazy: false },
    })
    await until(() => wrapper.findAll('[data-pptx-slide]').length === 3)
    expect(wrapper.findAll('[data-pptx-slide-placeholder]')).toHaveLength(0)
    wrapper.unmount()
  })
})

describe('accessible preview', () => {
  it('labels a slide with its title and a caption with its position', async () => {
    const Titled = defineComponent(
      () => () =>
        h(Presentation, null, () => [
          h(Slide, { title: 'Revenue' }, () => h(Text, { x: 0, y: 0, w: 3, h: 1 }, () => 'A')),
          h(Slide, null, () => h(Text, { x: 0, y: 0, w: 3, h: 1 }, () => 'B')),
        ]),
    )
    const wrapper = mount(Preview, { props: { document: Titled, debounce: 0, lazy: false } })
    await vi.waitFor(() => expect(wrapper.findAll('[data-pptx-slide]')).toHaveLength(2))
    const labels = wrapper.findAll('[data-pptx-slide]').map((node) => node.attributes('aria-label'))
    expect(labels).toEqual(['Slide 1: Revenue', 'Slide 2'])
    expect(wrapper.findAll('figcaption').map((node) => node.text())).toEqual(['Slide 1 of 2', 'Slide 2 of 2'])
    wrapper.unmount()
  })

  it('renders header rows as column headers', async () => {
    const wrapper = await draw(() =>
      h(Table, {
        x: 0,
        y: 0,
        w: 4,
        rows: [
          ['Name', 'Score'],
          ['Ada', '9'],
        ],
        headerRows: 1,
      }),
    )
    expect(wrapper.findAll('thead th').map((node) => node.attributes('scope'))).toEqual(['col', 'col'])
    expect(wrapper.findAll('tbody td').map((node) => node.text())).toEqual(['Ada', '9'])
    expect(wrapper.find('tbody th').exists()).toBe(false)
  })

  it('labels a chart and gives it a hidden data table', async () => {
    const wrapper = await draw(() => h(Chart, { type: 'bar', x: 0, y: 0, w: 6, h: 3, title: 'Sales', data: series }))
    const svg = wrapper.find('svg')
    expect(svg.attributes('role')).toBe('img')
    expect(svg.attributes('aria-label')).toBe('Sales, bar chart')
    const table = wrapper.find('table[data-pptx-chart-data]')
    expect(table.findAll('thead th').map((node) => node.text())).toEqual(['', 'Q1', 'Q2', 'Q3'])
    expect(table.findAll('tbody tr').map((row) => row.text())).toEqual(['North354', 'South216'])
  })

  it('prefers the described alt text and hides images without any', async () => {
    const PNG =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=='
    const wrapper = await draw(() => [
      h(Chart, { type: 'bar', x: 0, y: 0, w: 4, h: 2, altText: 'Sales by quarter', data: series }),
      h(Image, { src: PNG, x: 0, y: 3, w: 1, h: 1 }),
      h(Image, { src: PNG, x: 2, y: 3, w: 1, h: 1, altText: 'Logo' }),
    ])
    expect(wrapper.find('svg').attributes('aria-label')).toBe('Sales by quarter')
    const [bare, described] = wrapper.findAll('img')
    expect(bare!.attributes('aria-hidden')).toBe('true')
    expect(described!.attributes()).toMatchObject({ alt: 'Logo' })
    expect(described!.attributes('aria-hidden')).toBeUndefined()
  })
})
