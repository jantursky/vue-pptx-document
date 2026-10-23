<script setup lang="ts">
import { reactive, ref } from 'vue'
import { PptxDownloadButton, Preview } from 'vue-pptx-document'

import type { Row } from './report'
import ReportDeck from './ReportDeck.vue'

const quarter = ref('Q3 2026')
const rows = reactive<Row[]>([
  { team: 'Checkout', owner: 'Mira', status: 'Done' },
  { team: 'Search', owner: 'Tomas', status: 'At risk' },
  { team: 'Mobile', owner: 'Eva', status: 'On track' },
])
const statuses: Row['status'][] = ['On track', 'At risk', 'Done']
const add = () => rows.push({ team: `Team ${rows.length + 1}`, owner: 'Someone', status: 'On track' })
</script>

<template>
  <div class="split">
    <div class="stack">
      <label class="field">
        <span>Quarter</span>
        <input v-model="quarter" type="text" />
      </label>
      <div class="rows" role="group" aria-label="Teams">
        <div v-for="(row, index) in rows" :key="index" class="row">
          <input v-model="row.team" type="text" :aria-label="`Team ${index + 1}`" />
          <select v-model="row.status" :aria-label="`Status of ${row.team}`">
            <option v-for="status in statuses" :key="status">{{ status }}</option>
          </select>
        </div>
      </div>
      <div class="actions">
        <button type="button" class="button button-small" :disabled="rows.length >= 8" @click="add">Add team</button>
        <PptxDownloadButton
          class="button button-small button-primary"
          :document="ReportDeck"
          :document-props="{ rows, quarter }"
          file-name="quarterly-review.pptx"
        />
      </div>
    </div>
    <Preview class="viewer viewer-list" :document="ReportDeck" :document-props="{ rows, quarter }" notes />
  </div>
</template>
