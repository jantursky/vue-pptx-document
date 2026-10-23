<script setup lang="ts">
import { computed, ref } from 'vue'
import { SlidePreview, usePptx, useSlides } from 'vue-pptx-document'

import HelloDeck from '../examples/HelloDeck.vue'

const name = ref('headless')
const props = computed(() => ({ name: name.value }))

// useSlides() resolves the slides for a custom viewer; usePptx() keeps a .pptx Blob current.
const { model } = useSlides(HelloDeck, props)
const { blob, url, loading, error } = usePptx(HelloDeck, props, { debounce: 400 })

const size = computed(() => (blob.value ? `${(blob.value.size / 1024).toFixed(1)} kB` : '–'))
</script>

<template>
  <div class="stack">
    <label class="field">
      <span>Name</span>
      <input v-model="name" type="text" />
    </label>
    <dl class="state">
      <div>
        <dt>loading</dt>
        <dd data-state="loading">{{ loading }}</dd>
      </div>
      <div>
        <dt>blob</dt>
        <dd data-state="size">{{ size }}</dd>
      </div>
      <div>
        <dt>error</dt>
        <dd>{{ error ? String(error) : 'null' }}</dd>
      </div>
    </dl>
    <div class="thumb">
      <SlidePreview v-if="model" :model="model" :index="0" />
    </div>
    <a
      class="button button-small"
      :href="url ?? undefined"
      :aria-disabled="url ? undefined : 'true'"
      download="headless.pptx"
    >
      Download from url
    </a>
  </div>
</template>
