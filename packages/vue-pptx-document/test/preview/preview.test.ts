import { mount } from '@vue/test-utils'
import { describe, expect, it } from 'vitest'
import { h } from 'vue'

import { Chart, Paragraph, Shape, SlidePreview, Text, renderToModel } from '../../src'
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
          { type: 'bar', data: [series[0]] },
          { type: 'line', data: [series[1]], options: { secondaryValAxis: true, secondaryCatAxis: true } },
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
      h(Chart, { type: 'bar', barDir: 'bar', x: 0, y: 0, w: 8, h: 4.5, data: [series[0]] }),
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
