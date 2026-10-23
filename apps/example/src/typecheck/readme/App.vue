<!--
  Generated from packages/vue-pptx-document/README.md by scripts/samples.js; edit the source, then run
  `bun run sync:samples`. `bun run typecheck` compiles it against the library.
-->
<script setup lang="ts">
import { reactive } from 'vue'
import { PptxDownloadButton, Preview } from 'vue-pptx-document'

import ReportDeck from './ReportDeck.vue'

const report = reactive({ title: 'Q3 review', months: ['Jul', 'Aug', 'Sep'], revenue: [720, 790, 890] })
</script>

<template>
  <input v-model="report.title" />
  <Preview :document="ReportDeck" :document-props="report" />
  <PptxDownloadButton :document="ReportDeck" :document-props="report" file-name="report.pptx" />
</template>
