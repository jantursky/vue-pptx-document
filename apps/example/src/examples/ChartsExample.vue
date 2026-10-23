<script setup lang="ts">
import { ref } from 'vue'
import { PptxDownloadButton, Preview } from 'vue-pptx-document'

import ChartsDeck from './ChartsDeck.vue'

const types = [
  'Columns',
  'Stacked bars',
  '3D',
  'Lines',
  'Area',
  'Pie',
  'Doughnut',
  'Scatter',
  'Bubble',
  'Radar',
  'Combo',
]
const slide = ref(0)
const seed = ref(7)
</script>

<template>
  <div class="stack">
    <div class="controls">
      <div class="chips" role="group" aria-label="Chart type">
        <button
          v-for="(type, index) in types"
          :key="type"
          type="button"
          class="chip"
          :aria-pressed="slide === index"
          @click="slide = index"
        >
          {{ type }}
        </button>
      </div>
      <button type="button" class="button button-small" @click="seed++">New data</button>
      <PptxDownloadButton
        class="button button-small"
        :document="ChartsDeck"
        :document-props="{ seed }"
        file-name="charts.pptx"
      />
    </div>
    <Preview class="viewer" :document="ChartsDeck" :document-props="{ seed }" :slide="slide" />
  </div>
</template>
