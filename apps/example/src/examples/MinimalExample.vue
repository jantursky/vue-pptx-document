<script setup lang="ts">
import { ref } from 'vue'
import { PptxDownloadButton, Preview } from 'vue-pptx-document'

import HelloDeck from './HelloDeck.vue'

const name = ref('Ada')
</script>

<template>
  <div class="split">
    <div class="stack">
      <label class="field">
        <span>Name</span>
        <input v-model="name" type="text" />
      </label>
      <PptxDownloadButton
        class="button button-primary"
        :document="HelloDeck"
        :document-props="{ name }"
        file-name="hello.pptx"
      />
    </div>
    <Preview class="viewer" :document="HelloDeck" :document-props="{ name }" />
  </div>
</template>
