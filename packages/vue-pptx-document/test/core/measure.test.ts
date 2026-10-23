import { describe, expect, it } from 'vitest'

import { BULLET_INDENT, measureText, paragraphs, textInsets } from '../../src/core/measure'

describe('measureText', () => {
  it('splits paragraphs at breaks and newlines', () => {
    expect(
      paragraphs([{ text: 'a\nb' }, { text: 'c', options: { breakLine: true } }, { text: 'd' }]).map((p) =>
        p.map((run) => run.text).join(''),
      ),
    ).toEqual(['a', 'bc', 'd'])
  })

  it('reads PptxGenJS margins in points', () => {
    expect(textInsets(undefined)).toEqual({ left: 0.1, right: 0.1, top: 0.05, bottom: 0.05 })
    expect(textInsets(0)).toEqual({ left: 0, right: 0, top: 0, bottom: 0 })
    // [left, right, bottom, top]
    expect(textInsets([72, 36, 18, 9])).toEqual({ left: 1, right: 0.5, bottom: 0.25, top: 0.125 })
  })

  it('wraps long text and grows with font size', () => {
    const text = [{ text: 'The quick brown fox jumps over the lazy dog '.repeat(4) }]
    const one = measureText(text, { fontSize: 18 })
    const wrapped = measureText(text, { fontSize: 18 }, 3)
    expect(wrapped.height).toBeGreaterThan(one.height * 3)
    expect(wrapped.width).toBeLessThanOrEqual(3)
    expect(measureText(text, { fontSize: 36 }, 3).height).toBeGreaterThan(wrapped.height)
    expect(measureText(text, { fontSize: 18, wrap: false }, 3).height).toBeCloseTo(one.height)
  })

  it('leaves room for bullets', () => {
    const plain = measureText([{ text: 'Point' }], { margin: 0 })
    const bulleted = measureText([{ text: 'Point', options: { bullet: true } }], { margin: 0 })
    expect(bulleted.width - plain.width).toBeCloseTo(BULLET_INDENT)
  })
})
