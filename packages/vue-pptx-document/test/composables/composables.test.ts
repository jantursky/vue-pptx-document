import { afterEach, describe, expect, it, vi } from 'vitest'
import { defineComponent, effectScope, h, nextTick, ref } from 'vue'

import { Presentation, Slide, Text, usePptx, useSlides, type PptxWarning } from '../../src'
import Deck from '../fixtures/Deck.vue'

const until = (condition: () => boolean) =>
  vi.waitFor(() => {
    if (!condition()) throw new Error('not yet')
  })

describe('usePptx', () => {
  it('renders a blob and an object URL and re-renders after changes', async () => {
    const title = ref('First')
    const scope = effectScope()
    const result = scope.run(() => usePptx(Deck, () => ({ title: title.value, points: [] }), { debounce: 0 }))!
    expect(result.loading.value).toBe(true)
    await until(() => result.url.value !== null)
    const first = result.blob.value
    expect(first?.size).toBeGreaterThan(1000)
    expect(result.url.value).toMatch(/^blob:/)
    title.value = 'Second'
    await nextTick()
    await until(() => result.blob.value !== first)
    expect(result.error.value).toBeNull()
    scope.stop()
    expect(result.url.value).toBeNull()
  })

  it('waits for update() when auto is off', async () => {
    const scope = effectScope()
    const result = scope.run(() => usePptx(Deck, { title: 'T', points: [] }, { auto: false }))!
    expect(result.loading.value).toBe(false)
    expect(result.blob.value).toBeNull()
    const blob = await result.update()
    expect(blob).toBe(result.blob.value)
    scope.stop()
    expect(await result.update()).toBeNull()
  })
})

describe('useSlides', () => {
  it('resolves the model and follows props', async () => {
    const points = ref(['One'])
    const scope = effectScope()
    const result = scope.run(() => useSlides(Deck, () => ({ title: 'T', points: points.value }), { debounce: 0 }))!
    await until(() => result.model.value !== null)
    expect(result.model.value!.slides).toHaveLength(2)
    points.value = ['One', 'Two']
    await nextTick()
    await until(() => result.model.value!.slides.length === 3)
    expect(result.loading.value).toBe(false)
    scope.stop()
  })
})

describe('useSlides incremental updates', () => {
  afterEach(() => vi.useRealTimers())

  const Many = defineComponent({
    props: { second: { type: String, required: true } },
    setup: (props) => () =>
      h(Presentation, null, () =>
        ['One', props.second, 'Three', 'Four', 'Five'].map((text, index) =>
          h(Slide, { key: index }, () => h(Text, { x: 0, y: 0, w: 3, h: 1 }, () => text)),
        ),
      ),
  })

  it('keeps the objects of unchanged slides', async () => {
    vi.useFakeTimers()
    const second = ref('Two')
    const scope = effectScope()
    const result = scope.run(() => useSlides(Many, () => ({ second: second.value }), { debounce: 10 }))!
    await vi.waitFor(() => expect(result.model.value).not.toBeNull(), { interval: 5 })
    const before = result.model.value!
    second.value = 'Changed'
    await nextTick()
    await vi.advanceTimersByTimeAsync(20)
    await vi.waitFor(() => expect(result.model.value).not.toBe(before), { interval: 5 })
    const after = result.model.value!
    for (const index of [0, 2, 3, 4]) expect(after.slides[index]).toBe(before.slides[index])
    expect(after.slides[1]).not.toBe(before.slides[1])
    expect(after.masters).toBe(before.masters)
    expect(after.layout).toBe(before.layout)
    scope.stop()
  })
})

describe('warnings in composables', () => {
  // Shows loose text on the slide once `loose` is on: a reactive change that starts to warn.
  const Loose = defineComponent({
    props: { loose: Boolean },
    setup: (props) => () =>
      h(Presentation, null, () => [
        h(Slide, null, () => [h(Text, { x: 0, y: 0, w: 2, h: 1 }, () => 'Kept'), props.loose ? 'Loose' : null]),
      ]),
  })

  it('useSlides keeps the warnings of the latest update next to the model', async () => {
    const loose = ref(false)
    const seen: PptxWarning[] = []
    const scope = effectScope()
    const result = scope.run(() =>
      useSlides(Loose, () => ({ loose: loose.value }), { debounce: 0, onWarn: (warning) => seen.push(warning) }),
    )!
    await until(() => result.model.value !== null)
    expect(result.warnings.value).toEqual([])
    loose.value = true
    await nextTick()
    await until(() => result.warnings.value.length === 1)
    expect(result.warnings.value[0]).toMatchObject({ code: 'dropped-text', slide: 1 })
    expect(result.model.value!.slides[0]!.objects).toHaveLength(1)
    expect(seen).toHaveLength(1)
    scope.stop()
  })

  it('usePptx exposes warnings and logs them without onWarn', async () => {
    const log = vi.spyOn(console, 'warn').mockImplementation(() => undefined)
    const scope = effectScope()
    const result = scope.run(() => usePptx(Loose, { loose: true }, { auto: false }))!
    await result.update()
    expect(result.warnings.value.map((warning) => warning.code)).toEqual(['dropped-text'])
    expect(log).toHaveBeenCalledTimes(1)
    log.mockRestore()
    scope.stop()
  })

  it('usePptx reports a strict failure as its error', async () => {
    const scope = effectScope()
    const result = scope.run(() => usePptx(Loose, { loose: true }, { auto: false, strict: true }))!
    expect(await result.update()).toBeNull()
    expect(String(result.error.value)).toContain('dropped-text')
    scope.stop()
  })
})
