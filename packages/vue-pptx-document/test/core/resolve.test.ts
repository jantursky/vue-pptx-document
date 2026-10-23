import { describe, expect, it, vi } from 'vitest'
import { defineComponent, h } from 'vue'

import {
  Br,
  Chart,
  Image,
  Master,
  Notes,
  Paragraph,
  Presentation,
  Run,
  Shape,
  Slide,
  Table,
  TableCell,
  TableRow,
  Text,
  View,
  renderToModel,
} from '../../src'
import { resolveLayout } from '../../src/core/resolve'
import { singleSlide } from '../fixtures'

// A bare attribute, as a template writes `<Paragraph bullet>`; the component casts it to true.
const bare = '' as unknown as true

const objectsOf = async (children: () => unknown, slideProps: Record<string, unknown> = {}) =>
  (await renderToModel(singleSlide(children as never, slideProps))).slides[0]!.objects

describe('layouts', () => {
  it('names the built-in sizes and takes custom ones', () => {
    expect(resolveLayout(undefined)).toEqual({ name: 'LAYOUT_16x9', width: 10, height: 5.625 })
    expect(resolveLayout('wide')).toMatchObject({ name: 'LAYOUT_WIDE', width: 13.333 })
    expect(resolveLayout('LAYOUT_4x3')).toMatchObject({ width: 10, height: 7.5 })
    expect(resolveLayout({ width: '25.4cm', height: 5 })).toEqual({ name: 'CUSTOM', width: 10, height: 5 })
    expect(() => resolveLayout('A4')).toThrow('Unknown layout')
    expect(() => resolveLayout({ width: 1 })).toThrow('width and a height')
  })
})

describe('absolute frames', () => {
  it('reads inches, units and percentages of the slide', async () => {
    const [shape] = await objectsOf(() =>
      h(Shape, { type: 'rect', x: '10%', y: '72pt', w: '50%', h: '2.54cm', fill: '#111' }),
    )
    expect(shape).toEqual({
      kind: 'shape',
      shape: 'rect',
      frame: { x: 1, y: 1, w: 5, h: 1 },
      options: { fill: { color: '111111' } },
    })
  })

  it('fills the remaining width and measures text height', async () => {
    const [text] = await objectsOf(() => h(Text, { x: 1, y: 1, fontSize: 18 }, () => 'One line'))
    expect(text!.frame.w).toBe(9)
    expect(text!.frame.h).toBeGreaterThan(0.25)
    expect(text!.frame.h).toBeLessThan(0.6)
  })
})

describe('text', () => {
  it('collects runs, paragraphs and breaks and trims template whitespace', async () => {
    const [text] = await objectsOf(() =>
      h(Text, { x: 0, y: 0, w: 5, h: 1, bold: true }, () => [
        '  Hello ',
        h(Run, { color: '#f35815', href: 'https://vuejs.org' }, () => 'world'),
        h(Br),
        'next ',
      ]),
    )
    expect(text).toMatchObject({
      kind: 'text',
      runs: [
        { text: 'Hello ' },
        // <Br> ends the run before it.
        { text: 'world', options: { color: 'F35815', hyperlink: { url: 'https://vuejs.org' }, breakLine: true } },
        { text: 'next' },
      ],
      options: { bold: true },
    })
  })

  it('puts bullets on paragraphs', async () => {
    const [text] = await objectsOf(() =>
      h(Text, { w: 5, h: 1 }, () => [
        h(Paragraph, { bullet: bare }, () => 'One'),
        h(Paragraph, { bullet: { type: 'number' } }, () => 'Two'),
      ]),
    )
    expect(text).toMatchObject({
      runs: [
        { text: 'One', options: { bullet: true, breakLine: true } },
        { text: 'Two', options: { bullet: { type: 'number' } } },
      ],
    })
  })

  it('converts the CSS margin order to PptxGenJS', async () => {
    const [text] = await objectsOf(() => h(Text, { w: 5, h: 1, margin: [1, 2, 3, 4] }, () => 'x'))
    expect(text!.options).toMatchObject({ margin: [4, 2, 3, 1] })
  })
})

describe('flex layout', () => {
  it('lays out a slide with padding, gap and grow', async () => {
    const objects = await objectsOf(
      () => [
        h(Text, { fontSize: 24, h: 1 }, () => 'Title'),
        h(Shape, { type: 'rect', fill: '#eee', style: { flexGrow: 1 } }),
      ],
      { layout: 'flex', style: { padding: 0.5, gap: 0.25 } },
    )
    expect(objects.map((object) => object.frame)).toEqual([
      { x: 0.5, y: 0.5, w: 9, h: 1 },
      { x: 0.5, y: 1.75, w: 9, h: 3.375 },
    ])
  })

  it('splits a row by flex and draws a View background', async () => {
    const objects = await objectsOf(() =>
      h(View, { x: 0, y: 0, w: 10, h: 2, fill: '#fafafa', style: { flexDirection: 'row', gap: 1 } }, () => [
        h(Shape, { type: 'rect', fill: '#111', style: { flex: 1 } }),
        h(Shape, { type: 'rect', fill: '#222', style: { flex: 2 } }),
      ]),
    )
    expect(objects.map((object) => [object.kind, object.frame])).toEqual([
      ['shape', { x: 0, y: 0, w: 10, h: 2 }],
      ['shape', { x: 0, y: 0, w: 3, h: 2 }],
      ['shape', { x: 4, y: 0, w: 6, h: 2 }],
    ])
  })

  it('places absolute children inside a flex container', async () => {
    const objects = await objectsOf(
      () => h(Shape, { type: 'rect', fill: '#111', x: 1, y: 2, w: 3, h: 1, style: { position: 'absolute' } }),
      {
        layout: 'flex',
        style: { padding: 0.5 },
      },
    )
    expect(objects[0]!.frame).toEqual({ x: 1, y: 2, w: 3, h: 1 })
  })
})

describe('objects', () => {
  it('resolves images, tables and charts', async () => {
    const objects = await objectsOf(() => [
      h(Image, { src: 'data:image/png;base64,AAAA', x: 0, y: 0, w: 1, h: 1 }),
      h(Table, { x: 0, y: 1, w: 4 }, () => [
        h(TableRow, { bold: bare }, () => [h(TableCell, null, () => 'A'), h(TableCell, { colspan: 2 }, () => 'B')]),
      ]),
      h(Chart, { type: 'bar', x: 5, y: 0, w: 4, h: 3, data: [{ name: 'S', labels: ['a'], values: [1] }] }),
      h(Chart, {
        type: [
          { type: 'bar', data: [{ name: 'S', labels: ['a'], values: [1] }] },
          { type: 'line', data: [{ name: 'T', labels: ['a'], values: [2] }], options: { secondaryValAxis: true } },
        ],
        x: 5,
        y: 3,
        w: 4,
        h: 2,
      }),
    ])
    expect(objects.map((object) => object.kind)).toEqual(['image', 'table', 'chart', 'chart'])
    expect(objects[0]).toMatchObject({ src: 'data:image/png;base64,AAAA' })
    expect(objects[1]).toMatchObject({
      rows: [
        [
          { text: 'A', options: { bold: true } },
          { text: 'B', options: { bold: true, colspan: 2 } },
        ],
      ],
    })
    expect(objects[2]).toMatchObject({ type: 'bar', data: [{ name: 'S', values: [1] }] })
    expect(Array.isArray((objects[3] as { type: unknown }).type)).toBe(true)
  })

  it('skips what cannot sit on a slide', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const objects = await objectsOf(() => [h(Run, null, () => 'x')])
    expect(objects).toEqual([])
    expect(warn).toHaveBeenCalled()
    warn.mockRestore()
  })
})

describe('deck', () => {
  it('reads metadata, masters, sections, notes and slide numbers', async () => {
    const model = await renderToModel(
      defineComponent(
        () => () =>
          h(Presentation, { layout: 'wide', title: 'Deck', author: 'Ada', rtl: bare }, () => [
            h(Master, { name: 'BRAND', background: '#fafafa', slideNumber: bare }, () =>
              h(Shape, { type: 'rect', x: 0, y: 0, w: '100%', h: 0.1, fill: '#f35815' }),
            ),
            h(Slide, { master: 'BRAND', section: 'Intro', notes: 'First' }, () => [h(Notes, null, () => 'Second')]),
            h(Slide, { hidden: bare, background: { src: 'data:image/png;base64,AAAA' } }),
          ]),
      ),
    )
    expect(model.layout.name).toBe('LAYOUT_WIDE')
    expect(model.meta).toMatchObject({ title: 'Deck', author: 'Ada' })
    expect(model.rtl).toBe(true)
    expect(model.masters[0]).toMatchObject({
      name: 'BRAND',
      background: { color: 'FAFAFA' },
      slideNumber: { x: 12.333, y: 7.05 },
    })
    expect(model.masters[0]!.objects[0]!.frame.w).toBe(13.333)
    expect(model.slides[0]).toMatchObject({ master: 'BRAND', section: 'Intro', notes: 'First\nSecond' })
    expect(model.slides[1]).toMatchObject({ hidden: true, background: { src: 'data:image/png;base64,AAAA' } })
  })
})

describe('table margins', () => {
  it('keeps the CSS order PptxGenJS uses for tables', async () => {
    const [table] = await objectsOf(() => h(Table, { x: 0, y: 0, w: 4, rows: [['A']], margin: [1, 2, 3, 4] }))
    expect(table!.options).toMatchObject({ margin: [1, 2, 3, 4] })
  })

  it('takes strike styles without a prop warning', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const [text] = await objectsOf(() => h(Text, { w: 2, h: 1, strike: 'dblStrike' }, () => 'x'))
    expect(text!.options).toMatchObject({ strike: 'dblStrike' })
    expect(warn).not.toHaveBeenCalled()
    warn.mockRestore()
  })
})

describe('notes and hidden objects', () => {
  it('reads <Notes> inside views, in order', async () => {
    const model = await renderToModel(
      singleSlide(() => [
        h(Notes, null, () => 'First'),
        h(View, { x: 0, y: 0, w: 5, h: 2 }, () => [
          h(Text, null, () => 'Body'),
          h(View, null, () => h(Notes, null, () => 'Nested')),
        ]),
      ]),
    )
    expect(model.slides[0]!.notes).toBe('First\nNested')
    expect(model.slides[0]!.objects.map((object) => object.kind)).toEqual(['text'])
  })

  it('hides display: none objects and views on absolute slides', async () => {
    const objects = await objectsOf(() => [
      h(Shape, { type: 'rect', x: 0, y: 0, w: 1, h: 1, fill: '#111', style: { display: 'none' } }),
      h(View, { x: 0, y: 0, w: 5, h: 2, fill: '#eee', style: { display: 'none' } }, () =>
        h(Text, null, () => 'Hidden'),
      ),
      h(Text, { x: 0, y: 3, w: 5, h: 1 }, () => 'Shown'),
    ])
    expect(objects).toHaveLength(1)
    expect(objects[0]).toMatchObject({ kind: 'text', runs: [{ text: 'Shown' }] })
  })
})

describe('images', () => {
  it('skips an <Image> without a src instead of writing an empty path', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined)
    const model = await renderToModel(
      singleSlide(() => [h(Image, { x: 0, y: 0, w: 1, h: 1 }), h(Text, { x: 0, y: 1, w: 1, h: 1 }, () => 'Kept')]),
    )
    expect(model.slides[0]!.objects.map((object) => object.kind)).toEqual(['text'])
    expect(warn).toHaveBeenCalledWith(expect.stringContaining('<Image> without a `src` is skipped'))
    warn.mockRestore()
  })
})
