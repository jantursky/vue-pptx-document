import { mount } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { defineComponent, effectScope, h, nextTick, ref } from 'vue'

import { Notes, Preview, Presentation, Slide, Text, downloadPptx, usePptx, useSlides } from '../../src'

const Titled = defineComponent({
  props: { text: { type: String, default: 'Hello' } },
  setup: (props) => () =>
    h(Presentation, null, () =>
      h(Slide, null, () => [
        h(Text, { x: 0, y: 0, w: 3, h: 1 }, () => props.text),
        h(Notes, null, () => 'Speaker note'),
      ]),
    ),
})

// A root that is not a <Presentation> fails when it is resolved.
const Broken = defineComponent(() => () => h(Slide))

afterEach(() => {
  vi.useRealTimers()
  vi.restoreAllMocks()
})

describe('object URLs', () => {
  it('revokes the previous URL a while after a new render', async () => {
    let count = 0
    vi.spyOn(URL, 'createObjectURL').mockImplementation(() => `blob:test/${++count}`)
    const revoke = vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => undefined)
    // Catch the delayed revoke instead of waiting ten seconds for it.
    const delayed: (() => void)[] = []
    const real = globalThis.setTimeout
    vi.spyOn(globalThis, 'setTimeout').mockImplementation(((callback: () => void, delay?: number) => {
      if (delay === 10_000) {
        delayed.push(callback)
        return 0 as unknown as ReturnType<typeof setTimeout>
      }
      return real(callback, delay)
    }) as typeof setTimeout)

    const text = ref('One')
    const scope = effectScope()
    const result = scope.run(() => usePptx(Titled, () => ({ text: text.value }), { debounce: 0 }))!
    await vi.waitFor(() => expect(result.url.value).toBe('blob:test/1'))

    text.value = 'Two'
    await nextTick()
    await vi.waitFor(() => expect(result.url.value).toBe('blob:test/2'))
    expect(revoke).not.toHaveBeenCalledWith('blob:test/1')
    expect(delayed).toHaveLength(1)

    delayed[0]!()
    expect(revoke).toHaveBeenCalledWith('blob:test/1')
    scope.stop()
  })
})

describe('useSlides', () => {
  it('measures again when web fonts have loaded', async () => {
    const listeners = new Map<string, () => void>()
    const fonts = {
      addEventListener: (name: string, listener: () => void) => listeners.set(name, listener),
      removeEventListener: (name: string) => listeners.delete(name),
    }
    vi.stubGlobal('document', Object.assign(Object.create(document) as Document, { fonts }))
    try {
      const scope = effectScope()
      const result = scope.run(() => useSlides(Titled, undefined, { debounce: 0 }))!
      await vi.waitFor(() => expect(result.model.value).not.toBeNull())
      const first = result.model.value
      listeners.get('loadingdone')?.()
      expect(result.loading.value).toBe(true)
      await vi.waitFor(() => expect(result.loading.value).toBe(false))
      expect(result.model.value).not.toBeNull()
      expect(first).not.toBeNull()
      scope.stop()
      expect(listeners.has('loadingdone')).toBe(false)
    } finally {
      vi.unstubAllGlobals()
    }
  })

  it('keeps the error of a deck that throws', async () => {
    const scope = effectScope()
    const result = scope.run(() => useSlides(Broken, undefined, { debounce: 0 }))!
    await vi.waitFor(() => expect(result.error.value).toBeInstanceOf(Error))
    expect(result.model.value).toBeNull()
    expect(result.loading.value).toBe(false)
    scope.stop()
  })
})

describe('Preview', () => {
  it('shows the speaker notes of each slide when asked', async () => {
    const wrapper = mount(Preview, { props: { document: Titled, debounce: 0, lazy: false, notes: true } })
    await vi.waitFor(() => expect(wrapper.find('[data-pptx-preview-notes]').exists()).toBe(true))
    expect(wrapper.find('[data-pptx-preview-notes]').text()).toBe('Speaker note')
    wrapper.unmount()
  })

  it('shows the message of a deck that throws', async () => {
    const wrapper = mount(Preview, { props: { document: Broken, debounce: 0 } })
    await vi.waitFor(() => expect(wrapper.find('[data-pptx-preview-error]').exists()).toBe(true))
    expect(wrapper.find('[role="alert"]').text()).toContain('<Presentation>')
    wrapper.unmount()
  })

  it('hands warnings to onWarn', async () => {
    const seen: string[] = []
    const Loose = defineComponent(
      () => () =>
        h(Presentation, null, () => h(Slide, null, () => h(Text, { x: 0, y: 0, w: 3, h: 1, bogusProp: 1 }, () => 'A'))),
    )
    const wrapper = mount(Preview, {
      props: {
        document: Loose,
        debounce: 0,
        lazy: false,
        onWarn: (warning: { code: string }) => seen.push(warning.code),
      },
    })
    await vi.waitFor(() => expect(seen).toContain('unknown-prop'))
    wrapper.unmount()
  })
})

describe('downloadPptx', () => {
  const stubLink = () => {
    const names: string[] = []
    vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:test/download')
    vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => undefined)
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (this: HTMLAnchorElement) {
      names.push(this.download)
    })
    return names
  }

  it('adds the .pptx extension when the name has none', async () => {
    const names = stubLink()
    await downloadPptx(Titled, { text: 'A' }, 'report')
    expect(names).toEqual(['report.pptx'])
  })

  it('keeps a name that already ends in .pptx', async () => {
    const names = stubLink()
    await downloadPptx(Titled, { text: 'A' }, 'report.pptx')
    await downloadPptx(Titled, { text: 'A' })
    expect(names).toEqual(['report.pptx', 'presentation.pptx'])
  })
})
