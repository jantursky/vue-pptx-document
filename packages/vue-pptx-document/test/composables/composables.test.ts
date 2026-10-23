import { describe, expect, it, vi } from 'vitest'
import { effectScope, nextTick, ref } from 'vue'

import { usePptx, useSlides } from '../../src'
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
