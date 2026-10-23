<script setup lang="ts">
import { ref, shallowRef, type Component } from 'vue'

import { downloadPptx } from '../core/pptx'

const props = withDefaults(
  defineProps<{
    /** The presentation component, rendered on each click. */
    document: Component
    documentProps?: Record<string, unknown>
    fileName?: string
  }>(),
  { documentProps: undefined, fileName: 'presentation.pptx' },
)

const emit = defineEmits<{
  done: [blob: Blob]
  error: [error: unknown]
}>()

defineSlots<{
  default?: (props: { loading: boolean; error: unknown }) => unknown
}>()

const loading = ref(false)
const error = shallowRef<unknown>(null)

// The file is built on click, not on every change, so editing a large deck stays fast.
const download = async (): Promise<void> => {
  if (loading.value) return
  loading.value = true
  error.value = null
  try {
    emit('done', await downloadPptx(props.document, props.documentProps, props.fileName))
  } catch (caught) {
    error.value = caught
    emit('error', caught)
  } finally {
    loading.value = false
  }
}
</script>

<template>
  <button
    type="button"
    data-pptx-download
    :aria-busy="loading ? 'true' : undefined"
    :data-loading="loading ? '' : undefined"
    @click="download"
  >
    <slot :loading="loading" :error="error">{{ loading ? 'Preparing…' : 'Download PPTX' }}</slot>
  </button>
</template>
