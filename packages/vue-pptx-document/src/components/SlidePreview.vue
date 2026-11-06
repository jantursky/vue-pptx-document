<script setup lang="ts">
import { computed, type FunctionalComponent } from 'vue'

import type { DeckLayout, DeckModel, MasterModel, SlideModel } from '../core/model'
import { renderSlide } from '../preview/render'

const props = withDefaults(
  defineProps<{
    /** The resolved presentation, from `useSlides()` or `renderToModel()`. Not needed with `slide`. */
    model?: DeckModel
    /** Zero-based slide index. */
    index: number
    /**
     * One slide with the parts it draws on, instead of `model`. A slide whose props are all the same
     * objects as before isn't drawn again, which is how `<Preview>` skips unchanged slides.
     */
    slide?: SlideModel
    masters?: readonly MasterModel[]
    layout?: DeckLayout
    theme?: DeckModel['theme']
    /** Accessible name; defaults to "Slide N". */
    label?: string
  }>(),
  { label: undefined, model: undefined, slide: undefined, masters: undefined, layout: undefined, theme: undefined },
)

/** The single slide, with its masters and layout, from `slide` or picked out of `model`. */
const parts = computed(() => {
  const { model } = props
  const slide = props.slide ?? model?.slides[props.index]
  const layout = props.layout ?? model?.layout
  if (!slide || !layout) return undefined
  const masters = (props.masters ?? model?.masters ?? []) as MasterModel[]
  const theme = props.theme ?? model?.theme
  const deck: DeckModel = { layout, meta: {}, masters, slides: [slide], ...(theme ? { theme } : {}) }
  return { slide, layout, deck }
})
const current = computed(() => parts.value?.slide)

// Rendered with h() rather than a template: each slide object maps to a few absolutely placed boxes.
const Content: FunctionalComponent = () =>
  parts.value ? renderSlide(parts.value.deck, 0, { number: props.index + 1 }) : []
</script>

<template>
  <div
    v-if="current"
    data-pptx-slide
    :data-hidden="current.hidden ? '' : undefined"
    role="group"
    aria-roledescription="slide"
    :aria-label="label ?? (current.title ? `Slide ${index + 1}: ${current.title}` : `Slide ${index + 1}`)"
    :style="{
      position: 'relative',
      width: '100%',
      aspectRatio: `${parts!.layout.width} / ${parts!.layout.height}`,
      containerType: 'inline-size',
      overflow: 'hidden',
      isolation: 'isolate',
    }"
  >
    <Content />
  </div>
</template>
