<script setup lang="ts" generic="C extends Component">
import { computed, onBeforeUnmount, onMounted, shallowRef, type Component } from 'vue'

import { useSlides } from '../composables/useSlides'
import type { PptxWarning } from '../core/diagnostics'
import type { DeckModel } from '../core/model'
import type { DocumentProps } from '../core/props'
import SlidePreview from './SlidePreview.vue'

const props = withDefaults(
  defineProps<{
    /** The presentation component. It's read once, when the preview mounts. */
    document: C
    /** Typed from the presentation component's props. */
    documentProps?: DocumentProps<C>
    /** Show only this slide (zero-based). */
    slide?: number
    /** Show each slide's speaker notes under it. */
    notes?: boolean
    debounce?: number
    /**
     * Draw slides only when they are near the viewport; the others stay empty boxes of the right size.
     * Turn it off to render every slide at once, for screenshots.
     */
    lazy?: boolean
    label?: string
    /** Receives each warning, as `@warn`. Without it, warnings go to the console. */
    onWarn?: (warning: PptxWarning) => void
  }>(),
  {
    documentProps: undefined,
    slide: undefined,
    notes: false,
    debounce: 50,
    lazy: true,
    label: 'Presentation preview',
    onWarn: undefined,
  },
)

defineSlots<{
  loading?: () => unknown
  error?: (props: { error: unknown }) => unknown
  /** Replaces the caption under each slide. */
  caption?: (props: { index: number; total: number; model: DeckModel }) => unknown
}>()

const { model, loading, error } = useSlides<Component>(props.document, () => props.documentProps, {
  debounce: props.debounce,
  ...(props.onWarn ? { onWarn: (warning: PptxWarning) => props.onWarn?.(warning) } : {}),
})

const indexes = computed(() => {
  const total = model.value?.slides.length ?? 0
  if (props.slide !== undefined) return props.slide >= 0 && props.slide < total ? [props.slide] : []
  return Array.from({ length: total }, (_, index) => index)
})

// Lazy rendering: one observer; a slide that came near the viewport stays rendered, so find-in-page
// and screen readers keep seeing it.
const canObserve = typeof IntersectionObserver === 'function'
const shown = shallowRef<ReadonlySet<number>>(new Set())
const everything = shallowRef(!props.lazy || !canObserve)
const isShown = (index: number): boolean => everything.value || shown.value.has(index)

const elements = new Map<number, Element>()
const indexOf = new WeakMap<Element, number>()
let observer: IntersectionObserver | undefined

const track = (index: number, element: unknown): void => {
  const previous = elements.get(index)
  if (previous && previous !== element) {
    observer?.unobserve(previous)
    elements.delete(index)
  }
  if (element instanceof Element && !elements.has(index)) {
    elements.set(index, element)
    indexOf.set(element, index)
    if (!isShown(index)) observer?.observe(element)
  }
}

const revealAll = (): void => {
  everything.value = true
}

onMounted(() => {
  if (!props.lazy || !canObserve) return
  observer = new IntersectionObserver(
    (entries) => {
      const near: number[] = []
      for (const entry of entries) {
        const index = indexOf.get(entry.target)
        if (!entry.isIntersecting || index === undefined) continue
        near.push(index)
        observer?.unobserve(entry.target)
      }
      if (near.length > 0) shown.value = new Set([...shown.value, ...near])
    },
    { rootMargin: `${window.innerHeight}px 0px` },
  )
  for (const [index, element] of elements) if (!isShown(index)) observer.observe(element)
  // The page prints what is on it, so draw the rest first.
  window.addEventListener('beforeprint', revealAll)
})

onBeforeUnmount(() => {
  observer?.disconnect()
  observer = undefined
  window.removeEventListener('beforeprint', revealAll)
})
</script>

<template>
  <section
    data-pptx-preview
    :data-loading="loading ? '' : undefined"
    :aria-label="label"
    :aria-busy="loading ? 'true' : undefined"
  >
    <slot v-if="error" name="error" :error="error">
      <p role="alert" data-pptx-preview-error>{{ error instanceof Error ? error.message : String(error) }}</p>
    </slot>
    <template v-else-if="model">
      <figure
        v-for="index in indexes"
        :key="index"
        :ref="(element) => track(index, element)"
        data-pptx-preview-slide
        :style="{ margin: 0 }"
      >
        <div
          :style="{
            aspectRatio: `${model.layout.width} / ${model.layout.height}`,
            contentVisibility: 'auto',
            width: '100%',
          }"
        >
          <SlidePreview
            v-if="isShown(index) && model.slides[index]"
            :slide="model.slides[index]"
            :masters="model.masters"
            :layout="model.layout"
            :theme="model.theme"
            :index="index"
          />
          <div v-else data-pptx-slide-placeholder aria-hidden="true" :style="{ width: '100%', height: '100%' }" />
        </div>
        <figcaption data-pptx-preview-caption>
          <slot name="caption" :index="index" :total="model.slides.length" :model="model">
            {{ index + 1 }} / {{ model.slides.length }}
          </slot>
          <p v-if="notes && model.slides[index]?.notes" data-pptx-preview-notes>{{ model.slides[index]?.notes }}</p>
        </figcaption>
      </figure>
    </template>
    <slot v-else name="loading" />
  </section>
</template>
