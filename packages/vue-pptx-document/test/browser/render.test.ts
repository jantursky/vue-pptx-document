import { describe, expect, it } from 'vitest'
import { defineComponent, effectScope, h } from 'vue'

import { Chart, Presentation, PPTX_MIME, Slide, Text, View, renderToBlob, renderToModel, usePptx } from '../../src'

const Deck = defineComponent(
  () => () =>
    h(Presentation, { title: 'Browser' }, () =>
      h(Slide, { layout: 'flex', style: { padding: 0.5, gap: 0.25 } }, () => [
        h(
          Text,
          { fontSize: 28, bold: true },
          () => 'A title that is long enough to wrap onto a second line in a narrow box',
        ),
        h(View, { style: { flexDirection: 'row', flexGrow: 1, gap: 0.25 } }, () => [
          h(Text, { style: { flex: 1 } }, () => 'Left'),
          h(Chart, { type: 'bar', style: { flex: 1 }, data: [{ name: 'A', labels: ['a', 'b'], values: [1, 2] }] }),
        ]),
      ]),
    ),
)

const header = async (blob: Blob) => new Uint8Array(await blob.slice(0, 2).arrayBuffer())

describe('in a real browser', () => {
  it('writes a .pptx Blob', async () => {
    const blob = await renderToBlob(Deck)
    expect(blob.type).toBe(PPTX_MIME)
    expect(blob.size).toBeGreaterThan(5000)
    expect([...(await header(blob))]).toEqual([0x50, 0x4b])
  })

  it('measures text with the canvas', async () => {
    const model = await renderToModel(Deck)
    const title = model.slides[0]!.objects[0]!
    // 28pt bold over 9 inches: two lines.
    expect(title.frame.h).toBeGreaterThan(0.8)
    expect(title.frame.h).toBeLessThan(1.3)
  })

  it('gives usePptx an object URL the browser can fetch', async () => {
    const scope = effectScope()
    const { url, loading } = scope.run(() => usePptx(Deck, undefined, { debounce: 0 }))!
    await expect.poll(() => loading.value, { timeout: 15_000 }).toBe(false)
    const response = await fetch(url.value!)
    expect([...(await header(await response.blob()))]).toEqual([0x50, 0x4b])
    scope.stop()
  })
})
