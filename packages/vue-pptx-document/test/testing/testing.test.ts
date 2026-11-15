// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { defineComponent, h } from 'vue'

import { Chart, Image, Notes, Paragraph, Presentation, Shape, Slide, Table, Text, renderToBuffer } from '../../src'
import { findText, readPptx, toOutline } from '../../src/testing'

const PIXEL =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

const Deck = defineComponent(
  () => () =>
    h(Presentation, { title: 'Deck' }, () => [
      h(Slide, { title: 'First' }, () => [
        h(Text, { x: 1, y: 1, w: 8, h: 1 }, () => 'Hello'),
        h(Image, { src: PIXEL, x: 1, y: 2, w: 1, h: 1 }),
      ]),
      h(Slide, null, () => [
        h(Text, { x: 1, y: 1, w: 8, h: 2 }, () => [
          h(Paragraph, null, () => 'Revenue'),
          h(Paragraph, null, () => 'Costs'),
        ]),
        h(Table, { x: 1, y: 4, w: 4, rows: [['Cell A', 'Cell B']] }),
        h(Notes, null, () => 'Say hi'),
      ]),
    ]),
)

describe('readPptx', () => {
  it('lists the texts of a slide in order and its XML part', async () => {
    const deck = await readPptx(await renderToBuffer(Deck))
    expect(deck.slides).toHaveLength(2)
    expect(deck.slides[0]!.texts).toEqual(['Hello'])
    expect(deck.slides[1]!.texts[0]).toBe('Revenue\nCosts')
    expect(deck.slides[1]!.xml).toContain('<p:sld')
    expect(deck.slides[1]!.xml).toContain('Revenue')
    expect(deck.slides[1]!.notes).toBe('Say hi')
    expect(deck.model.slides).toHaveLength(2)
  })

  it('accepts a Blob and an ArrayBuffer', async () => {
    const bytes = await renderToBuffer(Deck)
    const copy = new Uint8Array(bytes)
    const fromBlob = await readPptx(new Blob([copy]))
    const fromBuffer = await readPptx(copy.buffer)
    expect(fromBlob.slides[0]!.texts).toEqual(['Hello'])
    expect(fromBuffer.slides[1]!.notes).toBe('Say hi')
  })
})

describe('findText', () => {
  it('returns the 1-based slide numbers for a string or a RegExp', async () => {
    const deck = await readPptx(await renderToBuffer(Deck))
    expect(findText(deck, 'Revenue')).toEqual([2])
    expect(findText(deck, /^Hel+o$/g)).toEqual([1])
    expect(findText(deck, 'nothing')).toEqual([])
  })
})

describe('toOutline', () => {
  it('is identical for two renders of the same deck, without image data', async () => {
    const first = toOutline(await readPptx(await renderToBuffer(Deck)))
    const second = toOutline(await readPptx(await renderToBuffer(Deck)))
    expect(first).toBe(second)
    expect(first).toContain('Slide 1\n  text')
    expect(first).toContain('text @ 1,1 8×1: Hello')
    expect(first).toMatch(/image @ 1,2 1×1: \d+ bytes/)
    expect(first).not.toContain('base64')
    expect(first).toContain('notes: Say hi')
  })
})

describe('toOutline of other objects', () => {
  const Mixed = defineComponent(
    () => () =>
      h(Presentation, null, () => [
        h(Slide, { hidden: true }, () => [
          h(Shape, { type: 'rect', x: 1, y: 1, w: 2, h: 1, fill: '#000000' }),
          h(Chart, {
            type: 'bar',
            x: 1,
            y: 3,
            w: 4,
            h: 2,
            data: [{ name: 'North', labels: ['Q1', 'Q2'], values: [1, 2] }],
          }),
          h(Table, { x: 5, y: 1, w: 3, rows: [['A', 'B']] }),
        ]),
        h(Slide),
      ]),
  )

  it('names each kind and marks a hidden slide', async () => {
    const outline = toOutline(await readPptx(await renderToBuffer(Mixed)))
    expect(outline).toContain('Slide 1 (hidden)')
    expect(outline).toMatch(/shape @ 1,1 2×1: rect/)
    expect(outline).toMatch(/chart @ 1,3 4×2: bar/)
    expect(outline).toMatch(/table @ 5,1 3×[\d.]+: A \| B/)
    expect(outline).toContain('Slide 2')
  })

  it('does not match a sticky RegExp from a previous search', async () => {
    const deck = await readPptx(await renderToBuffer(Deck))
    const pattern = /Hello/y
    expect(findText(deck, pattern)).toEqual([1])
    expect(findText(deck, pattern)).toEqual([1])
  })
})
