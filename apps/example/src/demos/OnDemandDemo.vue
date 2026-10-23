<script setup lang="ts">
import { ref } from 'vue'

const state = ref<'idle' | 'loading' | 'done' | 'error'>('idle')
const message = ref('')

// Both come in by import(), so in an app where nothing else imports them they load on the click.
// On this page the other examples have loaded them already.
const download = async (): Promise<void> => {
  state.value = 'loading'
  try {
    const [{ downloadPptx }, { default: ChartsDeck }] = await Promise.all([
      import('vue-pptx-document'),
      import('../examples/ChartsDeck.vue'),
    ])
    const blob = await downloadPptx(ChartsDeck, { seed: Date.now() % 1000 }, 'charts-on-demand.pptx')
    message.value = `${(blob.size / 1024).toFixed(1)} kB written`
    state.value = 'done'
  } catch (error) {
    message.value = error instanceof Error ? error.message : String(error)
    state.value = 'error'
  }
}
</script>

<template>
  <div class="stack">
    <button type="button" class="button button-primary" :disabled="state === 'loading'" @click="download">
      {{ state === 'loading' ? 'Rendering…' : 'Download eleven charts' }}
    </button>
    <p class="note" role="status" data-on-demand-status>{{ message }}</p>
  </div>
</template>
