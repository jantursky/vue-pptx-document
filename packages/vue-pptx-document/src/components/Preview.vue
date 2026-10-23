<script setup lang="ts" generic="C extends Component">
import { computed, type Component } from 'vue'

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
    label?: string
    /** Receives each warning, as `@warn`. Without it, warnings go to the console. */
    onWarn?: (warning: PptxWarning) => void
  }>(),
  {
    documentProps: undefined,
    slide: undefined,
    notes: false,
    debounce: 50,
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
      <figure v-for="index in indexes" :key="index" data-pptx-preview-slide :style="{ margin: 0 }">
        <SlidePreview :model="model" :index="index" />
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
