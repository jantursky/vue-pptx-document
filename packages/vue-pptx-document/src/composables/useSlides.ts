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
import { collect, type DiagnosticOptions, type PptxWarning } from '../core/diagnostics'
import { pptx, type PptxInstance } from '../core/pptx'
import type { ReactivePropsArgs } from '../core/props'

type Props = Record<string, unknown>

export interface UseSlidesOptions extends DiagnosticOptions {
  /** Milliseconds to wait after a change before resolving the slides again. */
  debounce?: number
}

export interface UseSlidesReturn {
  /** The resolved slides, as `<SlidePreview>` draws them. */
  model: Readonly<Ref<DeckModel | null>>
  loading: Readonly<Ref<boolean>>
  error: Readonly<Ref<unknown>>
  /** The warnings of the latest finished update. */
  warnings: Readonly<Ref<readonly PptxWarning[]>>
  update: () => Promise<void>
  /** The mounted presentation, for `toBlob()` and friends on the same tree. */
  instance: () => PptxInstance | undefined
}

/**
 * Resolves `presentation` to its slide model and keeps it current, without writing a file.
 * This is what `<Preview>` runs; use it for thumbnails or a custom viewer.
 */
export function useSlides<C extends Component>(
  presentation: C,
  ...[documentProps, options = {}]: ReactivePropsArgs<C, [options?: UseSlidesOptions]>
): UseSlidesReturn {
  const props = documentProps as MaybeRefOrGetter<Props | undefined>
  const { debounce = 50, ...diagnostics } = options
  const model = shallowRef<DeckModel | null>(null)
  const loading = ref(true)
  const error = shallowRef<unknown>(null)
  const warnings = shallowRef<readonly PptxWarning[]>([])
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
    const collected = collect(diagnostics)
    try {
      const result = await instance.toModel(collected.options)
      if (current !== run) return
      model.value = result
      error.value = null
    } catch (caught) {
      if (current === run) error.value = caught
    } finally {
      if (current === run) {
        loading.value = false
        warnings.value = collected.warnings
      }
    }
  }

  const start = (): void => {
    if (disposed) return
    instance = pptx<Component>(presentation, props)
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
    warnings: shallowReadonly(warnings),
    update,
    instance: () => instance,
  }
}
