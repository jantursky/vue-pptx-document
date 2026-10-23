<script setup lang="ts">
import { computed, ref } from 'vue'
import { PptxDownloadButton, Preview } from 'vue-pptx-document'

import LayoutDeck from './LayoutDeck.vue'

const all = ['Write slides as components', 'Lay them out with flexbox', 'Preview them live', 'Download a .pptx']
const count = ref(3)
const direction = ref<'row' | 'column'>('row')
const props = computed(() => ({ cards: all.slice(0, count.value), direction: direction.value }))
</script>

<template>
  <div class="stack">
    <div class="controls">
      <label class="field field-inline">
        <span>Cards: {{ count }}</span>
        <input v-model.number="count" type="range" min="1" max="4" />
      </label>
      <fieldset class="segmented">
        <legend>Flex direction</legend>
        <label><input v-model="direction" type="radio" value="row" /> row</label>
        <label><input v-model="direction" type="radio" value="column" /> column</label>
      </fieldset>
      <PptxDownloadButton
        class="button button-small"
        :document="LayoutDeck"
        :document-props="props"
        file-name="layout.pptx"
      />
    </div>
    <Preview class="viewer viewer-grid" :document="LayoutDeck" :document-props="props" />
  </div>
</template>
