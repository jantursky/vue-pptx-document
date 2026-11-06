import { mount } from '@vue/test-utils'
import { describe, expect, it } from 'vitest'
import { defineComponent, effectScope, h } from 'vue'

import { Preview, Presentation, Slide, Text, useSlides } from '../../src'
import { clearMeasureCache, measureCacheSize } from '../../src/core/measure'

const Many = defineComponent(
  () => () =>
    h(Presentation, null, () =>
      Array.from({ length: 30 }, (_, index) =>
        h(Slide, { key: index }, () => h(Text, { x: 0.5, y: 0.5, w: 6, h: 1 }, () => `Slide number ${index + 1}`)),
      ),
    ),
)

const drawn = (root: Element) => root.querySelectorAll('[data-pptx-slide]').length

describe('the preview in a real browser', () => {
  it('draws only the slides near the viewport, each figure with its final height', async () => {
    const host = document.createElement('div')
    host.style.cssText = 'width: 480px; height: 300px; overflow: auto'
    document.body.append(host)
    const wrapper = mount(Preview, { props: { document: Many, debounce: 0 }, attachTo: host })
    await expect.poll(() => host.querySelectorAll('figure').length).toBe(30)
    await expect.poll(() => drawn(host)).toBeGreaterThan(0)
    expect(drawn(host)).toBeLessThan(30)
    const heights = [...host.querySelectorAll('figure')].map((figure) => figure.getBoundingClientRect().height)
    expect(new Set(heights.map(Math.round)).size).toBe(1)
    wrapper.unmount()
    host.remove()
  })

  it('draws every slide with lazy off', async () => {
    const host = document.createElement('div')
    document.body.append(host)
    const wrapper = mount(Preview, { props: { document: Many, debounce: 0, lazy: false }, attachTo: host })
    await expect.poll(() => drawn(host)).toBe(30)
    wrapper.unmount()
    host.remove()
  })

  it('measures again after web fonts finish loading', async () => {
    const scope = effectScope()
    const result = scope.run(() => useSlides(Many, undefined, { debounce: 0 }))!
    await expect.poll(() => result.model.value).not.toBeNull()
    const first = result.model.value
    expect(measureCacheSize()).toBeGreaterThanOrEqual(0)
    clearMeasureCache()
    document.fonts.dispatchEvent(new Event('loadingdone'))
    await expect.poll(() => result.model.value).not.toBe(first)
    expect(result.model.value!.slides[0]).toBe(first!.slides[0])
    scope.stop()
  })
})
