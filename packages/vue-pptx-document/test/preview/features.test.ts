import { mount } from '@vue/test-utils'
import { describe, expect, it, vi } from 'vitest'
import { defineComponent, h } from 'vue'

import {
  Chart,
  Image,
  Master,
  Media,
  Placeholder,
  Presentation,
  Run,
  Shape,
  Slide,
  SlidePreview,
  Table,
  TableCell,
  TableRow,
  Text,
  View,
  renderToBuffer,
  renderToModel,
} from '../../src'
import { parsePptx, pptxToVue } from '../../src/import'
import { shapePath } from '../../src/preview/shapes'
import { unzip } from '../fixtures'

const PIXEL =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

const SHAPES = (
  'rect roundRect round2SameRect snip1Rect ellipse flowChartConnector triangle rtTriangle diamond flowChartDecision parallelogram ' +
  'flowChartInputOutput trapezoid pentagon hexagon heptagon octagon decagon dodecagon star4 star5 star6 star7 star8 star10 star12 ' +
  'rightArrow leftArrow upArrow downArrow leftRightArrow upDownArrow chevron homePlate plus mathPlus mathMinus donut heart cloud ' +
  'flowChartProcess flowChartAlternateProcess flowChartTerminator flowChartDocument can pie arc lightningBolt wedgeRectCallout ' +
  'wedgeRoundRectCallout frame line'
).split(' ')

// One deck with every kind of object and most options, for the preview and the writer.
const Everything = defineComponent(
  () => () =>
    h(
      Presentation,
      { layout: { width: 12, height: 6 }, title: 'Everything', subject: 'Tests', company: 'Acme' },
      () => [
        h(Master, { name: 'MAIN', background: { color: '#f5f5f5', transparency: 10 } }, () => [
          h(
            Placeholder,
            { name: 'heading', type: 'title', x: 0.5, y: 0.2, w: 11, h: 0.8, fontSize: 28 },
            () => 'Title here',
          ),
          h(Shape, { type: 'line', x: 0.5, y: 5.6, w: 11, h: 0, line: { color: '#ccc', width: 1, dashType: 'dash' } }),
          h(Image, { src: PIXEL, x: 11, y: 5.7, w: 0.2, h: 0.2 }),
          h(Text, { x: 0.5, y: 5.7, w: 5, h: 0.2, fontSize: 8 }, () => [
            'Footer ',
            h(Run, { bold: true }, () => 'bold'),
          ]),
        ]),
        h(Slide, { master: 'MAIN', slideNumber: { x: '90%', y: '95%', color: '#999' } }, () => [
          h(Placeholder, { name: 'heading' }, () => 'Filled title'),
          ...SHAPES.map((type, index) =>
            h(Shape, {
              type,
              x: (index % 12) * 0.9 + 0.3,
              y: Math.floor(index / 12) * 0.9 + 1.1,
              w: 0.8,
              h: 0.8,
              fill: { color: '#f35815', transparency: 20 },
              line: { color: '#111', width: 1 },
              rotate: index % 3 === 0 ? 15 : undefined,
              flipH: index % 5 === 0,
              shadow:
                index % 4 === 0
                  ? { type: 'outer', blur: 3, offset: 2, angle: 45, color: '000000', opacity: 0.4 }
                  : undefined,
              rectRadius: 0.1,
            }),
          ),
        ]),
        h(Slide, { background: { src: PIXEL } }, () => [
          h(
            Text,
            {
              x: 0.5,
              y: 0.5,
              w: 5,
              h: 1,
              shape: 'roundRect',
              fill: '#eee',
              line: '#333',
              valign: 'bottom',
              align: 'center',
              italic: true,
              underline: true,
              strike: 'sngStrike',
              highlight: '#ff0',
              charSpacing: 2,
              lineSpacing: 20,
              rotate: 5,
              vert: 'vert',
            },
            () => 'Styled text',
          ),
          h(
            Shape,
            { type: 'ellipse', x: 6, y: 0.5, w: 2, h: 1, fill: '#111', color: '#fff', fontSize: 14 },
            () => 'In a shape',
          ),
          h(
            Text,
            {
              x: 0.5,
              y: 2,
              w: 5,
              h: 1,
              superscript: true,
              fontFace: 'Georgia',
              glow: { size: 4, opacity: 0.3, color: 'FF0000' },
              outline: { size: 1, color: '000000' },
            },
            () => 'Sup',
          ),
          h(
            Text,
            {
              x: 0.5,
              y: 3,
              w: 5,
              h: 1,
              subscript: true,
              paraSpaceBefore: 6,
              lineSpacingMultiple: 1.5,
              bullet: { code: '25BA' },
              indentLevel: 1,
            },
            () => 'Sub',
          ),
          h(Image, {
            src: PIXEL,
            x: 6,
            y: 2,
            w: 2,
            h: 1,
            rounding: true,
            transparency: 50,
            sizing: { type: 'cover', w: 2, h: 1 },
            altText: 'Pixel',
            hyperlink: { url: 'https://vuejs.org' },
          }),
          h(Image, {
            src: PIXEL,
            x: 8.5,
            y: 2,
            w: 2,
            h: 1,
            sizing: { type: 'crop', x: 0, y: 0, w: 1, h: 1 },
            flipV: true,
          }),
          h(Media, { type: 'online', src: 'https://www.youtube.com/embed/xyz', x: 6, y: 3.5, w: 2, h: 1.2 }),
          h(Media, { src: 'data:video/mp4;base64,AAAA', cover: PIXEL, x: 8.5, y: 3.5, w: 2, h: 1.2 }),
          h(
            Table,
            {
              x: 0.5,
              y: 4.2,
              w: 5,
              colW: ['40%', '30%', '30%'],
              rowH: 0.3,
              border: { type: 'solid', color: '#999', pt: 1 },
              fill: '#fafafa',
              margin: 4,
              align: 'center',
            },
            () => [
              h(TableRow, { bold: true, fill: '#111', color: '#fff' }, () => [
                h(TableCell, { rowspan: 2 }, () => 'Tall'),
                h(TableCell, { colspan: 2 }, () => 'Wide'),
              ]),
              h(TableRow, null, () => [
                h(
                  TableCell,
                  {
                    border: [
                      { type: 'none' },
                      { type: 'dash', color: 'f00', pt: 2 },
                      { type: 'solid' },
                      { type: 'none' },
                    ],
                  },
                  () => 'B',
                ),
                h(TableCell, { valign: 'bottom' }, () => 'C'),
              ]),
            ],
          ),
          h(
            View,
            {
              x: 10.6,
              y: 0.5,
              w: 1.2,
              h: 3,
              fill: '#fff',
              line: { color: '#ccc' },
              rectRadius: 0.1,
              style: { padding: 0.1, gap: 0.1, justifyContent: 'space-between', alignItems: 'stretch' },
            },
            () => [
              h(Text, { fontSize: 10, style: { alignSelf: 'center' } }, () => 'Top'),
              h(View, { style: { display: 'none' } }, () => h(Text, null, () => 'Hidden')),
              h(Shape, { type: 'rect', fill: '#f35815', style: { height: '20%', marginTop: 0.05 } }),
              h(Text, { fontSize: 10, style: { flexBasis: 0.4, flexShrink: 0, aspectRatio: 2 } }, () => 'Bottom'),
            ],
          ),
        ]),
      ],
    ),
)

describe('every feature', () => {
  it('draws every preset shape path', () => {
    for (const type of SHAPES.filter((name) => name !== 'line' && name !== 'ellipse')) {
      expect(shapePath(type, 100, 50, 10), type).toMatch(/^M/)
    }
  })

  it('previews each slide', async () => {
    const model = await renderToModel(Everything)
    expect(model.layout).toEqual({ name: 'CUSTOM', width: 12, height: 6 })
    for (const index of model.slides.keys()) {
      const wrapper = mount(SlidePreview, { props: { model, index, label: `Slide ${index + 1} of Everything` } })
      expect(wrapper.find('[data-pptx-slide]').attributes('aria-label')).toBe(`Slide ${index + 1} of Everything`)
      if (index === 0) {
        expect(wrapper.text()).toContain('Filled title')
        expect(wrapper.text()).not.toContain('Title here')
        expect(wrapper.text()).toContain('Footer bold')
        expect(wrapper.findAll('svg').length).toBeGreaterThanOrEqual(SHAPES.length)
      } else {
        expect(wrapper.text()).toContain('Styled text')
        expect(wrapper.text()).toContain('In a shape')
        expect(wrapper.find('img[alt="Pixel"]').exists()).toBe(true)
        expect(wrapper.find('table').text()).toContain('Tall')
        expect(wrapper.find('td[rowspan="2"]').exists()).toBe(true)
        expect(wrapper.text()).not.toContain('Hidden')
      }
      wrapper.unmount()
    }
  })

  it('writes every object to the file', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const parts = await unzip(await renderToBuffer(Everything))
    warn.mockRestore()
    const layout = Object.entries(parts).find(
      ([name, xml]) => name.startsWith('ppt/slideLayouts/') && xml.includes('Title here'),
    )
    expect(layout).toBeDefined()
    expect(parts['ppt/slides/slide1.xml']).toContain('Filled title')
    expect(parts['ppt/slides/slide1.xml']).toContain('prst="star5"')
    expect(parts['ppt/slides/slide2.xml']).toContain('Styled text')
    expect(parts['ppt/slides/slide2.xml']).toContain('gridSpan="2"')
    expect(parts['docProps/app.xml']).toContain('Acme')
    expect(parts['ppt/presentation.xml']).toContain('cx="10972800"')
  })
})

const series = [
  { name: 'North', labels: ['Jan', 'Feb', 'Mar'], values: [3, -2, 4] },
  { name: 'South', labels: ['Jan', 'Feb', 'Mar'], values: [2, 5, 1] },
]
const xy = [
  { name: 'X', values: [1, 2, 3] },
  { name: 'Y', values: [3, 1, 2], sizes: [2, 4, 6] },
]

const Charts = defineComponent(
  () => () =>
    h(Presentation, null, () => [
      ...(
        [
          [
            'bar',
            { barGrouping: 'stacked', showValue: true, dataLabelFormatCode: '0.0', legendPos: 't', showLegend: true },
          ],
          [
            'bar',
            {
              barDir: 'bar',
              barGrouping: 'percentStacked',
              valAxisHidden: true,
              catAxisLabelRotate: -45,
              legendPos: 'l',
              showLegend: true,
            },
          ],
          [
            'bar3D',
            {
              barGrouping: 'clustered',
              valAxisMinVal: -5,
              valAxisMaxVal: 10,
              valAxisMajorUnit: 5,
              valGridLine: { style: 'none' },
            },
          ],
          [
            'line',
            {
              lineSmooth: true,
              lineDataSymbol: 'none',
              showValue: true,
              legendPos: 'r',
              showLegend: true,
              valAxisLabelFormatCode: '$0',
            },
          ],
          [
            'area',
            {
              barGrouping: 'stacked',
              chartColorsOpacity: 50,
              catAxisHidden: true,
              showCatAxisTitle: true,
              catAxisTitle: 'Month',
            },
          ],
          ['pie', { showPercent: true, showLabel: true, showLegend: true, legendPos: 'b', firstSliceAng: 90 }],
          ['doughnut', { holeSize: 40, showValue: true, dataBorder: { pt: 1, color: 'FFFFFF' } }],
          ['radar', { radarStyle: 'filled', showLegend: true }],
          ['scatter', { lineSize: 0, lineDataSymbolSize: 8, showLegend: true, legendPos: 'tr' }],
          ['bubble', { showValue: true }],
        ] as const
      ).map(([type, options]) =>
        h(Slide, null, () =>
          h(Chart, {
            type,
            x: 0.5,
            y: 0.5,
            w: 9,
            h: 4.5,
            data:
              type === 'scatter' || type === 'bubble'
                ? xy
                : type === 'pie' || type === 'doughnut'
                  ? [{ ...series[1]!, values: [2, 5, 1] }]
                  : series,
            showTitle: true,
            title: type,
            chartColors: ['#f35815', '#111111'],
            ...options,
          }),
        ),
      ),
    ]),
)

describe('charts end to end', () => {
  it('previews every chart type with its options', async () => {
    const model = await renderToModel(Charts)
    for (const index of model.slides.keys()) {
      const wrapper = mount(SlidePreview, { props: { model, index } })
      const svg = wrapper.find('svg')
      expect(svg.text()).toContain((model.slides[index]!.objects[0] as { type: string }).type)
      wrapper.unmount()
    }
  })

  it('reads every chart type back from the file', async () => {
    const { model, warnings } = await parsePptx(await renderToBuffer(Charts))
    expect(warnings).toEqual([])
    expect(model.slides.map((slide) => (slide.objects[0] as { type: string }).type)).toEqual([
      'bar',
      'bar',
      'bar3D',
      'line',
      'area',
      'pie',
      'doughnut',
      'radar',
      'scatter',
      'bubble',
    ])
    expect(model.slides[1]!.objects[0]).toMatchObject({ options: { barDir: 'bar', barGrouping: 'percentStacked' } })
    expect(model.slides[8]!.objects[0]).toMatchObject({ data: [{ values: [1, 2, 3] }, { values: [3, 1, 2] }] })
  })

  it('reads the feature deck back and writes it as code', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const bytes = await renderToBuffer(Everything)
    warn.mockRestore()
    const { model, code, warnings } = await pptxToVue(bytes, { embedImages: false })
    expect(model.layout).toMatchObject({ width: 12, height: 6 })
    expect(model.meta).toMatchObject({ title: 'Everything', subject: 'Tests', company: 'Acme' })
    expect(model.slides[1]!.background?.src).toBeDefined()
    expect(code).toContain('<Table')
    expect(code).toContain('<Image')
    expect(code).toMatch(/type="star5"/)
    expect(Array.isArray(warnings)).toBe(true)
  })
})
