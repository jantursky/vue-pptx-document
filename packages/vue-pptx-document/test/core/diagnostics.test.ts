import { describe, expect, it, vi } from 'vitest'
import { defineComponent, h } from 'vue'

import { Image, Presentation, Slide, Text, renderToBuffer, renderToModel, type PptxWarning } from '../../src'
import { at, createContext, inContext, locate, warn } from '../../src/core/diagnostics'

// Text straight inside a <Slide> is dropped: the simplest deck that warns.
const DroppedText = defineComponent(
  () => () =>
    h(Presentation, null, () => [
      h(Slide, null, () => [h(Text, { x: 0, y: 0, w: 2, h: 1 }, () => 'Fine')]),
      h(Slide, null, () => 'Loose text'),
    ]),
)

describe('the warning channel', () => {
  it('reports the slide and object of the active context', () => {
    const warnings: PptxWarning[] = []
    inContext(createContext({ onWarn: (warning) => warnings.push(warning) }), () =>
      at({ slide: 2 }, () => at({ object: '<Image>' }, () => warn('missing-src', 'No src.'))),
    )
    expect(warnings).toEqual([{ code: 'missing-src', message: 'No src.', slide: 2, object: '<Image>' }])
  })

  it('logs once per message and context without onWarn', () => {
    const log = vi.spyOn(console, 'warn').mockImplementation(() => undefined)
    const context = createContext()
    inContext(context, () => {
      warn('unknown-color', 'Same.')
      warn('unknown-color', 'Same.')
    })
    inContext(createContext(), () => warn('unknown-color', 'Same.'))
    expect(log).toHaveBeenCalledTimes(2)
    log.mockRestore()
  })

  it('throws in strict mode, naming the code', () => {
    expect(() =>
      inContext(createContext({ strict: true }), () => at({ slide: 1 }, () => warn('dropped-text', 'Gone.'))),
    ).toThrow('[vue-pptx-document] dropped-text: Gone. (slide 1)')
  })

  it('locates an error once and keeps the cause', () => {
    const original = new Error('[vue-pptx-document] Broken.')
    const error = locate(original, { slide: 3, object: '<Chart>' })
    expect(error.message).toBe('[vue-pptx-document] Broken. (slide 3, <Chart>)')
    expect(error.cause).toBe(original)
    expect(locate(error, { slide: 4 })).toBe(error)
    expect(locate(original, {})).toBe(original)
  })

  it('restores the outer location after a nested one', () => {
    const warnings: PptxWarning[] = []
    inContext(createContext({ onWarn: (warning) => warnings.push(warning) }), () =>
      at({ slide: 1 }, () => {
        at({ object: '<View>' }, () => undefined)
        warn('misplaced-object', 'Outer.')
      }),
    )
    expect(warnings[0]).toEqual({ code: 'misplaced-object', message: 'Outer.', slide: 1 })
  })
})

describe('diagnostics in renders', () => {
  it('collects warnings with onWarn instead of the console', async () => {
    const log = vi.spyOn(console, 'warn').mockImplementation(() => undefined)
    const warnings: PptxWarning[] = []
    await renderToBuffer(DroppedText, undefined, { onWarn: (warning) => warnings.push(warning) })
    expect(warnings).toEqual([expect.objectContaining({ code: 'dropped-text', slide: 2, object: '<Slide>' })])
    expect(log).not.toHaveBeenCalled()
    log.mockRestore()
  })

  it('rejects in strict mode', async () => {
    await expect(renderToModel(DroppedText, undefined, { strict: true })).rejects.toThrow(/dropped-text: .*slide 2/)
  })

  it('warns again on every render', async () => {
    const log = vi.spyOn(console, 'warn').mockImplementation(() => undefined)
    await renderToModel(DroppedText)
    await renderToModel(DroppedText)
    expect(log).toHaveBeenCalledTimes(2)
    log.mockRestore()
  })

  it('names the slide and object of an image that fails to load', async () => {
    const Broken = defineComponent(
      () => () =>
        h(Presentation, null, () => [
          h(Slide),
          h(Slide),
          h(Slide, null, () => [h(Image, { x: 0, y: 0, w: 1, h: 1, src: '/no/such/picture.png' })]),
        ]),
    )
    vi.spyOn(console, 'error').mockImplementation(() => undefined)
    const error: unknown = await renderToBuffer(Broken).catch((caught: unknown) => caught)
    expect(error).toBeInstanceOf(Error)
    expect((error as Error).message).toMatch(/slide 3, <Image>/)
    vi.restoreAllMocks()
  })
})
