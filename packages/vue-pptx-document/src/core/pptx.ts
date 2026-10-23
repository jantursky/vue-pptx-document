import type PptxGenJS from 'pptxgenjs'
import { h, toValue, type Component, type MaybeRefOrGetter } from 'vue'

import type { DeckModel } from './model'
import { createRoot } from './nodes'
import { renderer } from './renderer'
import { createContext, inContext, type DiagnosticOptions } from './diagnostics'
import type { PropsArgs, ReactivePropsArgs } from './props'
import { resolveIn } from './resolve'
import { snapshot } from './snapshot'
import { toPptxGenJS, writeModel, type WriteOptions } from './write'

type Props = Record<string, unknown>

export const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

export interface PptxInstance {
  /** The resolved slides: frames in inches and PptxGenJS options, as the preview draws them. */
  toModel(options?: DiagnosticOptions): Promise<DeckModel>
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
export const pptx = <C extends Component>(
  presentation: C,
  ...[documentProps, options = {}]: ReactivePropsArgs<C, [options?: DiagnosticOptions]>
): PptxInstance => {
  const props = documentProps as MaybeRefOrGetter<Props | undefined>
  const root = createRoot()
  const listeners = new Set<() => void>()
  root.onChange = () => {
    for (const listener of [...listeners]) listener()
  }
  const app = renderer.createApp({ render: () => h(presentation, toValue(props)) })
  app.mount(root)
  let mounted = true

  const toModel = async (call: DiagnosticOptions = {}): Promise<DeckModel> => {
    if (!mounted) throw new Error('[vue-pptx-document] This presentation was unmounted.')
    // One context for the snapshot and the resolve, so a warning is reported once per render.
    const context = createContext({ ...options, ...call })
    return resolveIn(
      context,
      inContext(context, () => snapshot(root)),
    )
  }

  const write = async (outputType: 'uint8array' | 'base64', call: WriteOptions = {}) => {
    const merged = { ...options, ...call }
    return writeModel(await toModel(merged), outputType, merged)
  }

  const onRender = (blob: Blob): void => {
    const first = root.children[0]
    const handler = first?.kind === 'element' ? first.props.onRender : undefined
    if (typeof handler === 'function') (handler as (args: { blob: Blob }) => void)({ blob })
  }

  return {
    toModel,
    async toBlob(call) {
      const bytes = (await write('uint8array', call)) as Uint8Array
      const blob = new Blob([bytes as BlobPart], { type: PPTX_MIME })
      onRender(blob)
      return blob
    },
    async toBuffer(call) {
      return (await write('uint8array', call)) as Uint8Array
    },
    async toBase64(call) {
      return (await write('base64', call)) as string
    },
    async toPptxGenJS() {
      return toPptxGenJS(await toModel(), options)
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
  props: unknown,
  options: DiagnosticOptions | undefined,
  run: (instance: PptxInstance) => Promise<T>,
): Promise<T> => {
  const instance = pptx(presentation, props as Props | undefined, options)
  try {
    return await run(instance)
  } finally {
    instance.unmount()
  }
}

/** Renders once to the bytes of a .pptx file. Works in Node, workers and the browser. */
export const renderToBuffer = <C extends Component>(
  presentation: C,
  ...[props, options]: PropsArgs<C, [options?: WriteOptions]>
): Promise<Uint8Array> => once(presentation, props, options, (instance) => instance.toBuffer(options))

/** Renders once to a Blob with the PowerPoint MIME type. */
export const renderToBlob = <C extends Component>(
  presentation: C,
  ...[props, options]: PropsArgs<C, [options?: WriteOptions]>
): Promise<Blob> => once(presentation, props, options, (instance) => instance.toBlob(options))

/** Renders once to the resolved model, without PptxGenJS. */
export const renderToModel = <C extends Component>(
  presentation: C,
  ...[props, options]: PropsArgs<C, [options?: DiagnosticOptions]>
): Promise<DeckModel> => once(presentation, props, options, (instance) => instance.toModel())

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
export const downloadPptx = async <C extends Component>(
  presentation: C,
  ...[props, fileName = 'presentation.pptx', options]: PropsArgs<C, [fileName?: string, options?: WriteOptions]>
): Promise<Blob> => {
  const blob = await once(presentation, props, options, (instance) => instance.toBlob(options))
  saveBlob(blob, fileName.endsWith('.pptx') ? fileName : `${fileName}.pptx`)
  return blob
}
