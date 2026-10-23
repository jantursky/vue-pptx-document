import { spawnSync } from 'node:child_process'
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { defineComponent, h } from 'vue'

import {
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
  Text,
  renderToBuffer,
  renderToModel,
} from '../../src'
import { literal, modelToVue, parsePptx, pptxToVue } from '../../src/import'

const PIXEL =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

const Source = defineComponent(
  () => () =>
    h(Presentation, { title: 'Round trip', author: 'Ada', layout: '4x3' }, () => [
      h(Master, { name: 'BRAND', slideNumber: { x: 9, y: 7, w: 0.6, h: 0.3, fontSize: 9, color: '#737373' } }, () =>
        h(Text, { x: 0.5, y: 7, w: 6, h: 0.3, fontSize: 9, color: '#737373', margin: 0 }, () => 'Footer'),
      ),
      h(Slide, { master: 'BRAND', background: '#fafafa', section: 'Intro' }, () => [
        h(Text, { x: 1, y: 1, w: 6, h: 1, fontSize: 28, bold: true }, () => [
          'Hello ',
          h(Run, { color: '#f35815', href: 'https://vuejs.org' }, () => 'world'),
        ]),
        h(Text, { x: 1, y: 2.5, w: 6, h: 2 }, () => [
          h(Paragraph, { bullet: true }, () => 'One'),
          h(Paragraph, { bullet: { type: 'number' } }, () => 'Two'),
        ]),
        h(Shape, { type: 'roundRect', x: 7.5, y: 1, w: 2, h: 1, fill: '#111111', rectRadius: 0.2 }),
        h(Image, { src: PIXEL, x: 7.5, y: 3, w: 1, h: 1 }),
        h(Notes, null, () => 'Say hi'),
      ]),
      h(Slide, { hidden: true }, () => [
        h(Chart, {
          type: 'bar',
          x: 0.5,
          y: 0.5,
          w: 5,
          h: 3,
          showLegend: true,
          data: [
            { name: 'North', labels: ['Q1', 'Q2'], values: [1, 2] },
            { name: 'South', labels: ['Q1', 'Q2'], values: [3, 1] },
          ],
        }),
        h(Table, {
          x: 6,
          y: 0.5,
          w: 3,
          rows: [
            ['A', 'B'],
            ['1', { text: '2', options: { bold: true, fill: { color: 'EEEEEE' } } }],
          ],
        }),
      ]),
    ]),
)

describe('parsePptx', () => {
  it('reads back what the library writes', async () => {
    const { model, warnings } = await parsePptx(await renderToBuffer(Source))
    expect(warnings).toEqual([])
    expect(model.layout).toMatchObject({ width: 10, height: 7.5 })
    expect(model.meta).toMatchObject({ title: 'Round trip', author: 'Ada' })
    const [first, second] = model.slides
    expect(first).toMatchObject({ section: 'Intro', notes: 'Say hi', background: { color: 'FAFAFA' } })
    expect(first!.slideNumber).toMatchObject({ x: 9, y: 7, fontSize: 9, color: '737373' })
    const kinds = first!.objects.map((object) => object.kind)
    expect(kinds).toEqual(['text', 'text', 'text', 'shape', 'image'])
    expect(first!.objects[1]).toMatchObject({
      runs: [
        { text: 'Hello ', options: { bold: true, fontSize: 28 } },
        { text: 'world', options: { color: 'F35815', hyperlink: { url: 'https://vuejs.org' } } },
      ],
    })
    expect(first!.objects[2]).toMatchObject({
      runs: [
        { text: 'One', options: { bullet: true } },
        { text: 'Two', options: { bullet: { type: 'number' } } },
      ],
    })
    expect(first!.objects[3]).toMatchObject({
      shape: 'roundRect',
      frame: { x: 7.5, y: 1, w: 2, h: 1 },
      options: { fill: { color: '111111' } },
    })
    expect((first!.objects[4] as { src: string }).src).toMatch(/^data:image\/png;base64,/)
    expect(second!.hidden).toBe(true)
    expect(second!.objects[0]).toMatchObject({
      kind: 'chart',
      type: 'bar',
      data: [{ name: 'North', labels: ['Q1', 'Q2'], values: [1, 2] }, { name: 'South' }],
    })
    expect(second!.objects[1]).toMatchObject({
      kind: 'table',
      rows: [
        [{ text: 'A' }, { text: 'B' }],
        [{ text: '1' }, { text: '2', options: { bold: true, fill: { color: 'EEEEEE' } } }],
      ],
    })
  })

  it('skips layout and master shapes when asked', async () => {
    const { model } = await parsePptx(await renderToBuffer(Source), { masterShapes: false })
    expect(model.slides[0]!.objects.some((object) => object.kind === 'text' && object.runs[0]?.text === 'Footer')).toBe(
      false,
    )
  })

  it('reads theme colors, groups and connectors from files PowerPoint writes', async () => {
    const zip = await JSZip.loadAsync(
      await renderToBuffer(
        defineComponent(
          () => () =>
            h(Presentation, null, () =>
              h(Slide, null, () => h(Shape, { type: 'rect', x: 1, y: 1, w: 1, h: 1, fill: '#000000' })),
            ),
        ),
      ),
    )
    const part = 'ppt/slides/slide1.xml'
    const xml = await zip.file(part)!.async('string')
    const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="90" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="3657600" cy="1828800"/></a:xfrm></p:grpSpPr>
      <p:sp><p:nvSpPr><p:cNvPr id="91" name="Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
      <a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></p:spPr></p:sp></p:grpSp>
      <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="92" name="Line"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="0" y="4572000"/><a:ext cx="9144000" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>
      <a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`
    zip.file(part, xml.replace('</p:spTree>', `${group}</p:spTree>`))
    const { model } = await parsePptx(await zip.generateAsync({ type: 'uint8array' }))
    const objects = model.slides[0]!.objects
    // The child sits at half scale inside the group.
    expect(objects[1]).toMatchObject({ kind: 'shape', shape: 'ellipse', frame: { x: 1, y: 1, w: 1, h: 0.5 } })
    expect((objects[1] as { options: { fill: { color: string } } }).options.fill.color).toMatch(/^[0-9A-F]{6}$/)
    expect(objects[2]).toMatchObject({
      kind: 'shape',
      shape: 'line',
      frame: { y: 5, w: 10, h: 0 },
      options: { line: { color: 'FF0000', width: 1 } },
    })
  })

  it('rejects files that are not presentations', async () => {
    const zip = new JSZip()
    zip.file('hello.txt', 'hi')
    await expect(parsePptx(await zip.generateAsync({ type: 'uint8array' }))).rejects.toThrow()
  })
})

describe('codegen', () => {
  it('writes literals the way a person would', () => {
    expect(literal({ a: 1, 'b-c': 'x', list: [true, null] })).toBe("{ a: 1, 'b-c': 'x', list: [true, null] }")
    // Single quotes, because the literal sits inside a double-quoted attribute.
    expect(literal("it's")).toBe("'it\\'s'")
  })

  it('writes a component that renders the same model', async () => {
    const { code } = await pptxToVue(await renderToBuffer(Source))
    expect(code).toContain("from 'vue-pptx-document'")
    expect(code).toContain('<Presentation layout="4x3" title="Round trip" author="Ada">')
    expect(code).toContain('<Slide background="#FAFAFA" section="Intro"')
    expect(code).toMatch(/<Paragraph bullet>One<\/Paragraph>/)
    expect(code).toContain('<Notes>Say hi</Notes>')
    expect(code).toContain('<Chart')
  })

  it('round-trips a model through code', async () => {
    const model = await renderToModel(Source)
    const code = modelToVue(model, { importFrom: '../src' })
    expect(code).toContain("from '../src'")
    expect(code).toContain('<Master name="BRAND"')
  })
})

describe('CLI', () => {
  it('converts a file and prints help', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'vue-pptx-'))
    const input = join(dir, 'deck.pptx')
    const output = join(dir, 'Deck.vue')
    writeFileSync(input, await renderToBuffer(Source))
    const cli = join(import.meta.dirname, '../../src/cli.ts')
    const run = spawnSync('bun', [cli, 'import', input, '-o', output, '--no-images'], { encoding: 'utf8' })
    expect(run.status).toBe(0)
    expect(run.stderr).toContain('2 slides written')
    expect(readFileSync(output, 'utf8')).toContain('<Presentation')
    const help = spawnSync('bun', [cli, '--help'], { encoding: 'utf8' })
    expect(help.status).toBe(0)
    expect(help.stdout).toContain('Usage: vue-pptx-document import')
    expect(spawnSync('bun', [cli], { encoding: 'utf8' }).status).toBe(1)
  })
})
