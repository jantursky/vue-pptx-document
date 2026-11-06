import { describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { defineComponent, h, nextTick, ref } from 'vue'

import {
  Chart,
  Image,
  Master,
  Notes,
  Presentation,
  Shape,
  Slide,
  Table,
  Text,
  PPTX_MIME,
  pptx,
  renderToBlob,
  renderToBuffer,
} from '../../src'
import Deck from '../fixtures/Deck.vue'
import { slideText, unzip } from '../fixtures'

const isZip = (bytes: Uint8Array) => bytes[0] === 0x50 && bytes[1] === 0x4b

describe('writing', () => {
  it('writes a .pptx from an SFC with props', async () => {
    const bytes = await renderToBuffer(Deck, { title: 'Quarterly', points: ['One', 'Two'] })
    expect(isZip(bytes)).toBe(true)
    const parts = await unzip(bytes)
    expect(Object.keys(parts).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))).toHaveLength(3)
    expect(slideText(parts['ppt/slides/slide1.xml']!)).toEqual(['Quarterly'])
    expect(slideText(parts['ppt/slides/slide3.xml']!)).toEqual(['Two'])
    expect(parts['docProps/core.xml']).toContain('<dc:title>Quarterly</dc:title>')
    // PptxGenJS's own defaults stay out of the file.
    expect(parts['docProps/core.xml']).not.toContain('PptxGenJS')
  })

  it('writes masters, notes, sections, charts and tables', async () => {
    const bytes = await renderToBuffer(
      defineComponent(
        () => () =>
          h(Presentation, null, () => [
            h(Master, { name: 'BRAND', slideNumber: true }, () => [
              h(Text, { x: 0.5, y: 5.2, w: 6, h: 0.3, fontSize: 9, margin: 0 }, () => 'Footer'),
              h(Shape, { type: 'rect', x: 0, y: 0, w: 10, h: 0.1, fill: '#f35815' }),
            ]),
            h(Slide, { master: 'BRAND', section: 'Results' }, () => [
              h(Chart, {
                type: 'bar',
                x: 0.5,
                y: 0.5,
                w: 5,
                h: 3,
                data: [{ name: 'A', labels: ['a', 'b'], values: [1, 2] }],
              }),
              h(Table, {
                x: 6,
                y: 0.5,
                w: 3,
                rows: [
                  ['Region', 'Revenue'],
                  ['North', '12'],
                ],
              }),
              h(Notes, null, () => 'Speak slowly'),
            ]),
          ]),
      ),
    )
    const parts = await unzip(bytes)
    const layouts = Object.entries(parts).filter(
      ([name]) => name.startsWith('ppt/slideLayouts/slideLayout') && name.endsWith('.xml'),
    )
    expect(layouts.some(([, xml]) => xml.includes('<a:t>Footer</a:t>') && xml.includes('F35815'))).toBe(true)
    expect(Object.keys(parts).some((name) => name.startsWith('ppt/charts/chart'))).toBe(true)
    expect(slideText(parts['ppt/slides/slide1.xml']!)).toEqual(expect.arrayContaining(['Region', 'North', '12']))
    expect(Object.values(parts).some((xml) => xml.includes('Speak slowly'))).toBe(true)
    expect(parts['ppt/presentation.xml']).toContain('name="Results"')
  })

  it('writes combo charts with their axis options', async () => {
    const labels = ['Jan', 'Feb']
    const bytes = await renderToBuffer(
      defineComponent(
        () => () =>
          h(Presentation, null, () =>
            h(Slide, null, () =>
              h(Chart, {
                type: [
                  {
                    type: 'bar',
                    data: [{ name: 'Revenue', labels, values: [1, 2] }],
                    options: { chartColors: ['#C1C1C1'] },
                  },
                  {
                    type: 'line',
                    data: [{ name: 'Margin', labels, values: [30, 40] }],
                    options: { secondaryValAxis: true, secondaryCatAxis: true },
                  },
                ],
                x: 0.5,
                y: 0.5,
                w: 8,
                h: 4,
                showLegend: true,
                legendPos: 'b',
                valAxes: [
                  { showValAxisTitle: true, valAxisTitle: 'Revenue' },
                  { showValAxisTitle: true, valAxisTitle: 'Margin' },
                ],
                catAxes: [{ catAxisTitle: 'Month' }, { catAxisHidden: true }],
              }),
            ),
          ),
      ),
    )
    const parts = await unzip(bytes)
    const chart = Object.entries(parts).find(([name]) => /^ppt\/charts\/chart\d+\.xml$/.test(name))![1]
    expect(chart).toContain('<c:barChart>')
    expect(chart).toContain('<c:lineChart>')
    expect(chart).toContain('<c:legendPos val="b"/>')
    expect(chart).toContain('Margin')
    const { parsePptx } = await import('../../src/import')
    const { model } = await parsePptx(bytes)
    expect(model.slides[0]!.objects[0]).toMatchObject({
      type: [
        { type: 'bar', options: { chartColors: ['C1C1C1'] } },
        { type: 'line', options: { secondaryValAxis: true, secondaryCatAxis: true } },
      ],
    })
  })

  it('makes a Blob with the PowerPoint type', async () => {
    const blob = await renderToBlob(Deck, { title: 'T', points: [] })
    expect(blob.type).toBe(PPTX_MIME)
    expect(isZip(new Uint8Array(await blob.arrayBuffer()))).toBe(true)
  })
})

describe('pptx()', () => {
  it('follows reactive props and reports changes', async () => {
    const title = ref('First')
    const instance = pptx(Deck, () => ({ title: title.value, points: [] }))
    const onChange = vi.fn()
    instance.onChange(onChange)
    expect((await instance.toModel()).meta.title).toBe('First')
    title.value = 'Second'
    await nextTick()
    expect(onChange).toHaveBeenCalled()
    expect((await instance.toModel()).meta.title).toBe('Second')
    expect(typeof (await instance.toBase64())).toBe('string')
    const deck = await instance.toPptxGenJS()
    expect(deck.title).toBe('Second')
    instance.unmount()
    await expect(instance.toModel()).rejects.toThrow('unmounted')
  })

  it('calls onRender on the Presentation with the blob', async () => {
    const onRender = vi.fn()
    const instance = pptx(defineComponent(() => () => h(Presentation, { onRender }, () => h(Slide))))
    const blob = await instance.toBlob()
    expect(onRender).toHaveBeenCalledWith({ blob })
    instance.unmount()
  })

  it('rejects a tree without a Presentation', async () => {
    const instance = pptx(defineComponent(() => () => h(Slide)))
    await expect(instance.toModel()).rejects.toThrow('<Presentation>')
    instance.unmount()
  })
})

describe('accessible output', () => {
  const PNG =
    'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=='
  const write = async (children: () => unknown, props: Record<string, unknown> = {}) => {
    const bytes = await renderToBuffer(
      defineComponent(() => () => h(Presentation, props, () => h(Slide, null, children as never))),
    )
    return (await unzip(bytes))['ppt/slides/slide1.xml']!
  }

  it('sets the deck language on runs that have none', async () => {
    const xml = await write(
      () => [
        h(Text, { x: 0, y: 0, w: 3, h: 1 }, () => 'Plain'),
        h(Text, { x: 0, y: 1, w: 3, h: 1, lang: 'en-US' }, () => 'Own'),
        h(Table, { x: 0, y: 2, w: 3, rows: [['Cell']] }),
      ],
      { lang: 'sk-SK' },
    )
    expect(xml).toMatch(/<a:rPr lang="sk-SK"[^>]*>(?:(?!<\/a:r>).)*<a:t>Plain<\/a:t>/)
    expect(xml).toMatch(/<a:rPr lang="en-US"[^>]*>(?:(?!<\/a:r>).)*<a:t>Own<\/a:t>/)
    expect(xml).toMatch(/<a:rPr lang="sk-SK"[^>]*>(?:(?!<\/a:r>).)*<a:t>Cell<\/a:t>/)
  })

  it('flags the header row of a table', async () => {
    const xml = await write(() => [
      h(Table, { x: 0, y: 0, w: 3, rows: [['A'], ['1']], headerRows: 1, objectName: 'Results' }),
      h(Table, { x: 0, y: 3, w: 3, rows: [['B']] }),
    ])
    expect(xml.match(/<a:tblPr firstRow="1"/g)).toHaveLength(1)
    expect(xml).toContain('name="Results"')
    expect(xml).not.toContain('pptx-a11y')
  })

  it('describes an image without alt text with nothing, never its data', async () => {
    const xml = await write(() => [
      h(Image, { src: PNG, x: 0, y: 0, w: 1, h: 1 }),
      h(Image, { src: PNG, x: 1, y: 0, w: 1, h: 1, altText: 'A logo' }),
      h(Image, { src: PNG, x: 2, y: 0, w: 1, h: 1, altText: 'Ignored', decorative: true }),
    ])
    const descriptions = [...xml.matchAll(/<p:cNvPr id="\d+" name="[^"]+" descr="([^"]*)"/g)].map(([, text]) => text)
    expect(descriptions).toEqual(['', 'A logo', ''])
    expect(xml).not.toContain('pptx-a11y')
    expect(xml).not.toContain('base64')
  })

  it('writes the description of a chart', async () => {
    const xml = await write(() =>
      h(Chart, {
        type: 'bar',
        x: 0,
        y: 0,
        w: 4,
        h: 3,
        altText: 'Sales by quarter',
        data: [{ name: 'A', labels: ['Q1'], values: [1] }],
      }),
    )
    expect(xml).toContain('descr="Sales by quarter"')
  })

  it('opens the zip only when something needs a fix', async () => {
    const render = (children: () => unknown) =>
      renderToBuffer(defineComponent(() => () => h(Presentation, null, () => h(Slide, null, children as never))))
    const spy = vi.spyOn(JSZip, 'loadAsync')
    await render(() => h(Text, { x: 0, y: 0, w: 3, h: 1 }, () => 'Nothing to fix'))
    expect(spy).not.toHaveBeenCalled()
    await render(() => h(Image, { src: PNG, x: 0, y: 0, w: 1, h: 1 }))
    expect(spy).toHaveBeenCalledTimes(1)
    spy.mockRestore()
  })
})
