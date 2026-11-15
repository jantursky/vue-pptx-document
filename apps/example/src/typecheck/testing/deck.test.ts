/*
  Generated from apps/docs/docs/guides/testing.mdx by scripts/samples.js; edit the source, then run
  `bun run sync:samples`. `bun run typecheck` compiles it against the library.
*/
import { expect, it } from 'vitest'
import { h } from 'vue'
import { Presentation, Slide, Text, renderToBuffer } from 'vue-pptx-document'
import { readPptx } from 'vue-pptx-document/testing'

const Hello = {
  render: () =>
    h(Presentation, { title: 'Hello' }, () =>
      h(Slide, null, () => h(Text, { x: 0.5, y: 2, w: 9, h: 1 }, () => 'Hello from a test')),
    ),
}

it('writes the greeting', async () => {
  const deck = await readPptx(await renderToBuffer(Hello, {}))

  expect(deck.slides).toHaveLength(1)
  expect(deck.slides[0]?.texts).toEqual(['Hello from a test'])
})
