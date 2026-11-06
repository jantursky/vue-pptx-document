// Times resolving, laying out and previewing a 100-slide report deck. Run after `bun run build`.
// The numbers vary by machine and runner, so this prints them and never fails.
import { performance } from 'node:perf_hooks'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { createSSRApp, defineComponent, h } from 'vue'
import { renderToString } from 'vue/server-renderer'

const lib = await import(pathToFileURL(resolve(import.meta.dirname, '../dist/index.js')).href)
const { Chart, Presentation, Slide, SlidePreview, Table, Text, View, renderToModel } = lib

const SLIDES = 100
const rows = Array.from({ length: 6 }, (_, row) => [
  `Team ${row + 1}`,
  `Owner ${row + 1}`,
  row % 2 ? 'At risk' : 'On track',
])
const series = [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [4, 7, 6, 9] }]

/** Slides placed by x/y: a title, a paragraph, a table and a chart. */
const Absolute = defineComponent(
  () => () =>
    h(Presentation, { title: 'Bench' }, () =>
      Array.from({ length: SLIDES }, (_, index) =>
        h(Slide, { key: index }, () => [
          h(Text, { x: 0.5, y: 0.3, w: 9, fontSize: 28, bold: true }, () => `Quarterly report, part ${index + 1}`),
          h(
            Text,
            { x: 0.5, y: 1.1, w: 4.2, fontSize: 14 },
            () =>
              'Revenue grew in every region while costs stayed flat, so the margin widened for the third quarter in a row.',
          ),
          h(Table, { x: 5, y: 1.1, w: 4.5, rows }),
          h(Chart, { type: 'bar', x: 0.5, y: 3, w: 9, h: 2.4, data: series }),
        ]),
      ),
    ),
)

/** The same content laid out with Yoga, where text is measured several times per node. */
const Flex = defineComponent(
  () => () =>
    h(Presentation, { title: 'Bench' }, () =>
      Array.from({ length: SLIDES }, (_, index) =>
        h(Slide, { key: index, layout: 'flex', style: { padding: 0.5, gap: 0.2 } }, () => [
          h(Text, { fontSize: 28, bold: true }, () => `Quarterly report, part ${index + 1}`),
          h(View, { style: { flexDirection: 'row', gap: 0.3, flexGrow: 1 } }, () => [
            h(
              Text,
              { fontSize: 14, style: { flex: 1 } },
              () =>
                'Revenue grew in every region while costs stayed flat, so the margin widened for the third quarter in a row.',
            ),
            h(Table, { rows, style: { flex: 1 } }),
          ]),
        ]),
      ),
    ),
)

const time = async (label, run) => {
  const start = performance.now()
  const result = await run()
  const ms = performance.now() - start
  console.log(`${label.padEnd(44)} ${ms.toFixed(1).padStart(8)} ms`)
  return result
}

console.log(`${SLIDES}-slide deck, Node ${process.version}`)
await time('resolve, absolute layout (cold)', () => renderToModel(Absolute))
await time('resolve, absolute layout (warm cache)', () => renderToModel(Absolute))
await time('layout, flex with Yoga (cold)', () => renderToModel(Flex))
await time('layout, flex with Yoga (warm cache)', () => renderToModel(Flex))
const model = await renderToModel(Absolute)
await time('preview, 100 slides as HTML (server-rendered)', async () => {
  for (let index = 0; index < SLIDES; index++) {
    await renderToString(createSSRApp({ render: () => h(SlidePreview, { model, index }) }))
  }
})
