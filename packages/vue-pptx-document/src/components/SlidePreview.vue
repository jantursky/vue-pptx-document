<script setup lang="ts">
import { computed, type FunctionalComponent } from 'vue'

import type { DeckModel } from '../core/model'
import { renderSlide } from '../preview/render'

const props = withDefaults(
  defineProps<{
    /** The resolved presentation, from `useSlides()` or `renderToModel()`. */
    model: DeckModel
    /** Zero-based slide index. */
    index: number
    /** Accessible name; defaults to "Slide N". */
    label?: string
  }>(),
  { label: undefined },
)

const slide = computed(() => props.model.slides[props.index])

// Rendered with h() rather than a template: each slide object maps to a few absolutely placed boxes.
const Content: FunctionalComponent = () => renderSlide(props.model, props.index, { number: props.index + 1 })
</script>

<template>
  <div
    v-if="slide"
    data-pptx-slide
    :data-hidden="slide.hidden ? '' : undefined"
    role="group"
    aria-roledescription="slide"
    :aria-label="label ?? `Slide ${index + 1}`"
    :style="{
      position: 'relative',
      width: '100%',
      aspectRatio: `${model.layout.width} / ${model.layout.height}`,
      containerType: 'inline-size',
      overflow: 'hidden',
      isolation: 'isolate',
    }"
  >
    <Content />
  </div>
</template>
