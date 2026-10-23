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

import { pptx, type PptxInstance } from '../core/pptx'
import type { WriteOptions } from '../core/write'

type Props = Record<string, unknown>

/** How long a replaced object URL stays valid, so a download that's still reading it can finish. */
const REVOKE_DELAY = 10_000

export interface UsePptxOptions extends WriteOptions {
  /** Milliseconds to wait after a change before rendering again. */
  debounce?: number
  /** Render on mount and after every change. Off: render only when you call `update()`. */
  auto?: boolean
}

export interface UsePptxReturn {
  blob: Readonly<Ref<Blob | null>>
  url: Readonly<Ref<string | null>>
  loading: Readonly<Ref<boolean>>
  error: Readonly<Ref<unknown>>
  /** Renders now and resolves with the new Blob. */
  update: () => Promise<Blob | null>
}

/**
 * Renders `presentation` to a .pptx Blob and an object URL, and renders again after reactive
 * changes. Rendering starts on mount, so server-side rendering never runs PptxGenJS.
 */
export function usePptx(
  presentation: Component,
  props?: MaybeRefOrGetter<Props | undefined>,
  options: UsePptxOptions = {},
): UsePptxReturn {
  const { debounce = 300, auto = true, ...write } = options
  const blob = shallowRef<Blob | null>(null)
  const url = ref<string | null>(null)
  const loading = ref(auto)
  const error = shallowRef<unknown>(null)
  let instance: PptxInstance | undefined
  let stop: (() => void) | undefined
  let disposed = false
  let run = 0
  let timer: ReturnType<typeof setTimeout> | undefined
  const retiring = new Map<string, ReturnType<typeof setTimeout>>()

  const retire = (old: string): void => {
    retiring.set(
      old,
      setTimeout(() => {
        retiring.delete(old)
        URL.revokeObjectURL(old)
      }, REVOKE_DELAY),
    )
  }

  const setUrl = (next: string | null): void => {
    if (url.value) retire(url.value)
    url.value = next
  }

  const start = (): void => {
    if (disposed || instance) return
    instance = pptx(presentation, props)
    if (auto) {
      stop = instance.onChange(() => {
        clearTimeout(timer)
        timer = setTimeout(() => void update(), debounce)
      })
    }
  }

  const update = async (): Promise<Blob | null> => {
    start()
    if (!instance) return null
    clearTimeout(timer)
    const current = ++run
    loading.value = true
    try {
      const result = await instance.toBlob(write)
      if (current !== run) return result
      blob.value = result
      setUrl(URL.createObjectURL(result))
      error.value = null
      return result
    } catch (caught) {
      if (current === run) error.value = caught
      return null
    } finally {
      if (current === run) loading.value = false
    }
  }

  const begin = (): void => {
    start()
    if (auto) void update()
  }

  if (getCurrentInstance()) onMounted(begin)
  else begin()

  onScopeDispose(() => {
    disposed = true
    run++
    clearTimeout(timer)
    stop?.()
    instance?.unmount()
    for (const [old, pending] of retiring) {
      clearTimeout(pending)
      URL.revokeObjectURL(old)
    }
    retiring.clear()
    if (url.value) URL.revokeObjectURL(url.value)
    url.value = null
  })

  return {
    blob: shallowReadonly(blob),
    url: shallowReadonly(url),
    loading: shallowReadonly(loading),
    error: shallowReadonly(error),
    update,
  }
}
