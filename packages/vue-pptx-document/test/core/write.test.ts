import { describe, expect, it, vi } from 'vitest'
import { defineComponent, h, nextTick, ref } from 'vue'

import {
  Chart,
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
