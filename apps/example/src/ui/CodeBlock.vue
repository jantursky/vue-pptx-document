<script setup lang="ts">
import { computed } from 'vue'

import { highlight } from './highlight'

const props = defineProps<{ code: string; fileName: string }>()

const html = computed(() => highlight(props.code, props.fileName))
</script>

<template>
  <!-- highlight.js escapes the source, and the source is the demo's own files. -->
  <!-- eslint-disable-next-line vue/no-v-html -->
  <pre class="code" tabindex="0"><code v-html="html" /></pre>
</template>
