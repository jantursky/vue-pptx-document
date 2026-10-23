import type PptxGenJS from 'pptxgenjs'
import { h, toValue, type Component, type MaybeRefOrGetter } from 'vue'

import type { DeckModel } from './model'
import { createRoot } from './nodes'
import { renderer } from './renderer'
import { resolve } from './resolve'
import { snapshot } from './snapshot'
import { toPptxGenJS, writeModel, type WriteOptions } from './write'

type Props = Record<string, unknown>

export const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

export interface PptxInstance {
  /** The resolved slides: frames in inches and PptxGenJS options, as the preview draws them. */
  toModel(): Promise<DeckModel>
  toBlob(options?: WriteOptions): Promise<Blob>
  toBuffer(options?: WriteOptions): Promise<Uint8Array>
  toBase64(options?: WriteOptions): Promise<string>
  /** The PptxGenJS presentation, for anything the components don't cover. */
  toPptxGenJS(): Promise<PptxGenJS>
  /** Calls `listener` after every change to the tree. Returns a function that unsubscribes. */
  onChange(listener: () => void): () => void
  unmount(): void
}

/**
 * Mounts `presentation` in a Vue app of its own, with the renderer that builds the slide tree.
 * Reactive `props` re-render it; call `unmount()` when you're done.
 */
export const pptx = (presentation: Component, props?: MaybeRefOrGetter<Props | undefined>): PptxInstance => {
  const root = createRoot()
  const listeners = new Set<() => void>()
  root.onChange = () => {
    for (const listener of [...listeners]) listener()
  }
  const app = renderer.createApp({ render: () => h(presentation, toValue(props)) })
  app.mount(root)
  let mounted = true

  const toModel = (): Promise<DeckModel> => {
    if (!mounted) return Promise.reject(new Error('[vue-pptx-document] This presentation was unmounted.'))
    try {
      return resolve(snapshot(root))
    } catch (error) {
      return Promise.reject(error instanceof Error ? error : new Error(String(error)))
    }
  }

  const onRender = (blob: Blob): void => {
    const first = root.children[0]
    const handler = first?.kind === 'element' ? first.props.onRender : undefined
    if (typeof handler === 'function') (handler as (args: { blob: Blob }) => void)({ blob })
  }

  return {
    toModel,
    async toBlob(options) {
      const bytes = (await writeModel(await toModel(), 'uint8array', options)) as Uint8Array
      const blob = new Blob([bytes as BlobPart], { type: PPTX_MIME })
      onRender(blob)
      return blob
    },
    async toBuffer(options) {
      return (await writeModel(await toModel(), 'uint8array', options)) as Uint8Array
    },
    async toBase64(options) {
      return (await writeModel(await toModel(), 'base64', options)) as string
    },
    async toPptxGenJS() {
      return toPptxGenJS(await toModel())
    },
    onChange(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    unmount() {
      if (!mounted) return
      mounted = false
      listeners.clear()
      root.onChange = null
      app.unmount()
    },
  }
}

const once = async <T>(
  presentation: Component,
  props: Props | undefined,
  run: (instance: PptxInstance) => Promise<T>,
): Promise<T> => {
  const instance = pptx(presentation, props)
  try {
    return await run(instance)
  } finally {
    instance.unmount()
  }
}

/** Renders once to the bytes of a .pptx file. Works in Node, workers and the browser. */
export const renderToBuffer = (presentation: Component, props?: Props, options?: WriteOptions): Promise<Uint8Array> =>
  once(presentation, props, (instance) => instance.toBuffer(options))

/** Renders once to a Blob with the PowerPoint MIME type. */
export const renderToBlob = (presentation: Component, props?: Props, options?: WriteOptions): Promise<Blob> =>
  once(presentation, props, (instance) => instance.toBlob(options))

/** Renders once to the resolved model, without PptxGenJS. */
export const renderToModel = (presentation: Component, props?: Props): Promise<DeckModel> =>
  once(presentation, props, (instance) => instance.toModel())

/** Saves a Blob in the browser through a temporary link. */
export const saveBlob = (blob: Blob, fileName: string): void => {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = fileName
  link.rel = 'noopener'
  link.style.display = 'none'
  document.body.append(link)
  link.click()
  link.remove()
  // Safari reads the URL after the click returns.
  setTimeout(() => URL.revokeObjectURL(url), 10_000)
}

/** Renders once and saves the file in the browser. */
export const downloadPptx = async (
  presentation: Component,
  props?: Props,
  fileName = 'presentation.pptx',
  options?: WriteOptions,
): Promise<Blob> => {
  const blob = await renderToBlob(presentation, props, options)
  saveBlob(blob, fileName.endsWith('.pptx') ? fileName : `${fileName}.pptx`)
  return blob
}
