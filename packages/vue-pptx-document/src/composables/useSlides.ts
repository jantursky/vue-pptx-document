import {
  getCurrentInstance,
  onMounted,
  onScopeDispose,
  ref,
  shallowReadonly,
  shallowRef,
  type Component,
  type MaybeRefOrGetter,
  type Ref,
} from 'vue'

import type { DeckModel } from '../core/model'
import { pptx, type PptxInstance } from '../core/pptx'

type Props = Record<string, unknown>

export interface UseSlidesOptions {
  /** Milliseconds to wait after a change before resolving the slides again. */
  debounce?: number
}

export interface UseSlidesReturn {
  /** The resolved slides, as `<SlidePreview>` draws them. */
  model: Readonly<Ref<DeckModel | null>>
  loading: Readonly<Ref<boolean>>
  error: Readonly<Ref<unknown>>
  update: () => Promise<void>
  /** The mounted presentation, for `toBlob()` and friends on the same tree. */
  instance: () => PptxInstance | undefined
}

/**
 * Resolves `presentation` to its slide model and keeps it current, without writing a file.
 * This is what `<Preview>` runs; use it for thumbnails or a custom viewer.
 */
export function useSlides(
  presentation: Component,
  props?: MaybeRefOrGetter<Props | undefined>,
  options: UseSlidesOptions = {},
): UseSlidesReturn {
  const { debounce = 50 } = options
  const model = shallowRef<DeckModel | null>(null)
  const loading = ref(true)
  const error = shallowRef<unknown>(null)
  let instance: PptxInstance | undefined
  let stop: (() => void) | undefined
  let disposed = false
  let run = 0
  let timer: ReturnType<typeof setTimeout> | undefined

  const update = async (): Promise<void> => {
    if (!instance) return
    clearTimeout(timer)
    const current = ++run
    loading.value = true
    try {
      const result = await instance.toModel()
      if (current !== run) return
      model.value = result
      error.value = null
    } catch (caught) {
      if (current === run) error.value = caught
    } finally {
      if (current === run) loading.value = false
    }
  }

  const start = (): void => {
    if (disposed) return
    instance = pptx(presentation, props)
    stop = instance.onChange(() => {
      clearTimeout(timer)
      timer = setTimeout(() => void update(), debounce)
    })
    void update()
  }

  if (getCurrentInstance()) onMounted(start)
  else start()

  onScopeDispose(() => {
    disposed = true
    run++
    clearTimeout(timer)
    stop?.()
    instance?.unmount()
    instance = undefined
  })

  return {
    model: shallowReadonly(model),
    loading: shallowReadonly(loading),
    error: shallowReadonly(error),
    update,
    instance: () => instance,
  }
}
