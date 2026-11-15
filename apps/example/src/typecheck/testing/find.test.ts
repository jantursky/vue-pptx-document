/*
  Generated from apps/docs/docs/guides/testing.mdx by scripts/samples.js; edit the source, then run
  `bun run sync:samples`. `bun run typecheck` compiles it against the library.
*/
import { readFile } from 'node:fs/promises'

import { expect, it } from 'vitest'
import { findText, readPptx } from 'vue-pptx-document/testing'

it('puts the revenue table on slide 2', async () => {
  const deck = await readPptx(await readFile('report.pptx'))

  expect(findText(deck, 'Revenue')).toEqual([2])
  expect(findText(deck, /^Q[1-4] 2026$/)).toEqual([1])
})
