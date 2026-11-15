import { describe, expect, it } from 'vitest'
import type { Component } from 'vue'

import { renderToBuffer, renderToModel } from '../src'
import { parsePptx } from '../src/import'
import ChartsDeck from '../../../apps/example/src/examples/ChartsDeck.vue'
import HelloDeck from '../../../apps/example/src/examples/HelloDeck.vue'
import LayoutDeck from '../../../apps/example/src/examples/LayoutDeck.vue'
import ReportDeck from '../../../apps/example/src/examples/ReportDeck.vue'
import { sampleReport } from '../../../apps/example/src/generator/model'
import ReportTemplate from '../../../apps/example/src/generator/ReportTemplate.vue'
import { compareModels } from './roundtrip'

interface Case {
  name: string
  deck: Component
  props: Record<string, unknown>
}

const rows = [
  { team: 'Checkout', owner: 'Mira', status: 'Done' },
  { team: 'Search', owner: 'Tomas', status: 'At risk' },
  { team: 'Mobile', owner: 'Eva', status: 'On track' },
]

const cases: Case[] = [
  { name: 'the Hello example', deck: HelloDeck, props: { name: 'Jan' } },
  {
    name: 'the Layout example, in a row',
    deck: LayoutDeck,
    props: { cards: ['One', 'Two', 'Three'], direction: 'row' },
  },
  {
    name: 'the Layout example, in a column',
    deck: LayoutDeck,
    props: { cards: ['One', 'Two', 'Three', 'Four'], direction: 'column' },
  },
  { name: 'the Charts example', deck: ChartsDeck, props: { seed: 7 } },
  { name: 'the Report example', deck: ReportDeck, props: { rows, quarter: 'Q3 2026' } },
  { name: 'the Report generator template', deck: ReportTemplate, props: { report: sampleReport() } },
]

describe('round trip', () => {
  it.each(cases)('imports $name as the model it was written from', async ({ deck, props }) => {
    const written = await renderToModel(deck, props as never)
    const bytes = await renderToBuffer(deck, props as never)
    const { model } = await parsePptx(bytes, { masterShapes: false })
    expect(written.slides.length).toBeGreaterThan(0)
    expect(written.slides.every((slide) => slide.objects.length > 0)).toBe(true)
    expect(compareModels(written, model)).toEqual([])
  })

  it('reports a text and a frame that differ', async () => {
    const written = await renderToModel(HelloDeck, { name: 'Jan' })
    const { model } = await parsePptx(await renderToBuffer(HelloDeck, { name: 'Jan' }), { masterShapes: false })
    const changed = structuredClone(model)
    const text = changed.slides[0]!.objects.find((object) => object.kind === 'text')!
    text.runs = [{ text: 'Other' }]
    text.frame.x += 0.5
    const differences = compareModels(written, changed)
    expect(differences.some((line) => line.includes('text written'))).toBe(true)
    expect(differences.some((line) => line.includes('x written'))).toBe(true)
  })
})
