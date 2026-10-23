import { readFileSync } from 'node:fs'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'
import { camelize, h } from 'vue'

import { Text, renderToModel, type PptxWarning } from '../../src'
import { KNOWN_PROPS } from '../../src/core/known-props'
import { TYPES } from '../../src/core/nodes'
import { generate } from '../../scripts/known-props.js'
import { singleSlide } from '../fixtures'

const warningsOf = async (children: Parameters<typeof singleSlide>[0]): Promise<PptxWarning[]> => {
  const warnings: PptxWarning[] = []
  await renderToModel(singleSlide(children), undefined, { onWarn: (warning) => warnings.push(warning) })
  return warnings
}

describe('unknown props', () => {
  it('warns about a misspelled prop and suggests the right one', async () => {
    const warnings = await warningsOf(() => h(Text, { x: 0, y: 0, w: 2, h: 1, fontsize: 20 }, () => 'Hi'))
    expect(warnings).toEqual([expect.objectContaining({ code: 'unknown-prop', slide: 1, object: '<Text>' })])
    expect(warnings[0]!.message).toContain('`fontsize`')
    expect(warnings[0]!.message).toContain('Did you mean `fontSize`?')
  })

  it('stays quiet for known props, listeners and Vue keys', async () => {
    const warnings = await warningsOf(() =>
      h(Text, { key: 'a', x: 0, y: 0, w: 2, h: 1, fontSize: 20, bold: true, onClick: () => undefined }, () => 'Hi'),
    )
    expect(warnings).toEqual([])
  })

  it('knows every prop the components reference documents', () => {
    const doc = readFileSync(join(import.meta.dirname, '../../../../apps/docs/docs/reference/components.mdx'), 'utf8')
    const types: Record<string, string> = Object.fromEntries(
      Object.entries(TYPES).map(([key, type]) => [
        key === 'row' ? 'TableRow' : key === 'cell' ? 'TableCell' : key,
        type,
      ]),
    )
    const missing: string[] = []
    let checked = 0
    for (const section of doc.split(/^## /m).slice(1)) {
      const name = section.slice(0, section.indexOf('\n')).trim()
      const type = types[name.charAt(0).toLowerCase() + name.slice(1)] ?? types[name]
      const known = type ? KNOWN_PROPS[type] : undefined
      if (known === undefined) continue
      for (const row of section.split('\n').filter((line) => line.startsWith('| `'))) {
        const cell = row.split('|')[1] ?? ''
        for (const [, prop] of cell.matchAll(/`([a-z][\w-]*)`/g)) {
          checked += 1
          if (!known.split(' ').includes(camelize(prop!))) missing.push(`${name}.${prop!}`)
        }
      }
    }
    expect(missing).toEqual([])
    expect(checked).toBeGreaterThan(80)
  })

  it('has a known-props list generated from the current types', async () => {
    const current = readFileSync(join(import.meta.dirname, '../../src/core/known-props.ts'), 'utf8')
    expect(current).toBe(await generate())
  }, 20_000)
})
