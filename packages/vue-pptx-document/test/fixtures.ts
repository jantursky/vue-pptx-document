import JSZip from 'jszip'
import { defineComponent, h, type VNodeChild } from 'vue'

import { Presentation, Slide } from '../src'

/** A presentation whose single slide renders `children`. */
export const singleSlide = (children: () => VNodeChild, slideProps: Record<string, unknown> = {}) =>
  defineComponent(() => () => h(Presentation, null, () => h(Slide, slideProps, children)))

export const flush = () => new Promise((resolve) => setTimeout(resolve, 0))

/** The parts of a written .pptx, by path. */
export const unzip = async (bytes: Uint8Array): Promise<Record<string, string>> => {
  const zip = await JSZip.loadAsync(bytes)
  const out: Record<string, string> = {}
  for (const [name, file] of Object.entries(zip.files)) {
    if (!file.dir && /\.(xml|rels)$/.test(name)) out[name] = await file.async('string')
  }
  return out
}

/** Every `a:t` text of a slide part, in order. */
export const slideText = (xml: string): string[] => [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(([, text]) => text!)
