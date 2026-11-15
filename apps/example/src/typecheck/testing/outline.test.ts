/*
  Generated from apps/docs/docs/guides/testing.mdx by scripts/samples.js; edit the source, then run
  `bun run sync:samples`. `bun run typecheck` compiles it against the library.
*/
import { expect, it } from 'vitest'
import { h } from 'vue'
import { Presentation, Slide, Text, renderToBuffer } from 'vue-pptx-document'
import { readPptx, toOutline } from 'vue-pptx-document/testing'

const Hello = {
  render: () => h(Presentation, null, () => h(Slide, null, () => h(Text, { x: 0.5, y: 2, w: 9, h: 1 }, () => 'Hello'))),
}

it('keeps its layout', async () => {
  const outline = toOutline(await readPptx(await renderToBuffer(Hello, {})))

  expect(outline).toMatchInlineSnapshot(`
    "Slide 1
      text @ 0.5,2 9×1: Hello"
  `)
})
