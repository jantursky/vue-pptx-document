import { describe, expect, it } from 'vitest'
import { h } from 'vue'

import { Image, renderToBuffer, Shape, Table, TableCell, TableRow } from '../../src'
import { singleSlide, unzip } from '../fixtures'

// A 1x1 PNG.
const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

const links = (xml: string) => xml.match(/<a:hlinkClick\b[^>]*\/?>/g) ?? []

/** The URL the first `hlinkClick` in `xml` points to, read through the slide's relationships. */
const targetOf = (xml: string, rels: string) => {
  const id = /<a:hlinkClick[^>]*r:id="([^"]+)"/.exec(xml)?.[1]
  const relationship = (rels.match(/<Relationship\b[^>]*>/g) ?? []).find((tag) => tag.includes(`Id="${id}"`))
  return /Target="([^"]+)"/.exec(relationship ?? '')?.[1]
}

describe('link shorthand on objects', () => {
  it('clickable image: href becomes a hyperlink on the picture', async () => {
    const bytes = await renderToBuffer(
      singleSlide(() => h(Image, { src: PNG, x: 1, y: 1, w: 1, h: 1, href: 'https://example.com' })),
    )
    const parts = await unzip(bytes)
    const slide = parts['ppt/slides/slide1.xml']!
    expect(links(slide)).toHaveLength(1)
    expect(parts['ppt/slides/_rels/slide1.xml.rels']).toContain('Target="https://example.com"')
  })

  it('shape: href is a hyperlink, and an explicit hyperlink wins', async () => {
    const bytes = await renderToBuffer(
      singleSlide(() => [
        h(Shape, { type: 'rect', x: 1, y: 1, w: 1, h: 1, href: 'https://a.example' }),
        h(Shape, {
          type: 'rect',
          x: 3,
          y: 1,
          w: 1,
          h: 1,
          href: 'https://a.example',
          hyperlink: { url: 'https://b.example' },
        }),
      ]),
    )
    const parts = await unzip(bytes)
    const rels = parts['ppt/slides/_rels/slide1.xml.rels']!
    expect(rels).toContain('Target="https://a.example"')
    expect(rels).toContain('Target="https://b.example"')
    expect(links(parts['ppt/slides/slide1.xml']!)).toHaveLength(2)
  })

  it('row link reaches its cells; a cell with its own link keeps it', async () => {
    const bytes = await renderToBuffer(
      singleSlide(() =>
        h(Table, { x: 1, y: 1, w: 6 }, () =>
          h(TableRow, { href: 'https://row.example' }, () => [
            h(TableCell, null, () => 'Plain'),
            h(TableCell, { href: 'https://cell.example' }, () => 'Own'),
          ]),
        ),
      ),
    )
    const parts = await unzip(bytes)
    const rels = parts['ppt/slides/_rels/slide1.xml.rels']!
    expect(rels).toContain('Target="https://row.example"')
    expect(rels).toContain('Target="https://cell.example"')
    const xml = parts['ppt/slides/slide1.xml']!
    expect(links(xml)).toHaveLength(2)
    // The plain cell links to the row; the cell with its own href to its own URL.
    const [plain, own] = xml.split(/<a:tc[ >]/).slice(1)
    expect(targetOf(plain!, rels)).toBe('https://row.example')
    expect(targetOf(own!, rels)).toBe('https://cell.example')
  })
})
