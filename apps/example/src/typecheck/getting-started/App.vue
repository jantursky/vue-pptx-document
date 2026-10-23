<!--
  Generated from apps/docs/docs/getting-started.mdx by scripts/samples.js; edit the source, then run
  `bun run sync:samples`. `bun run typecheck` compiles it against the library.
-->
<script setup lang="ts">
import { ref } from 'vue'
import { PptxDownloadButton, Preview } from 'vue-pptx-document'

import HelloDeck from './HelloDeck.vue'

const name = ref('Ada')
</script>

<template>
  <input v-model="name" />
  <PptxDownloadButton :document="HelloDeck" :document-props="{ name }" file-name="hello.pptx" />
  <Preview :document="HelloDeck" :document-props="{ name }" />
</template>
