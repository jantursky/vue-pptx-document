<script setup lang="ts">
import { ref, shallowRef, type Component } from 'vue'

import { downloadPptx } from '../core/pptx'

// The attributes belong to the button, not to the wrapper around it and its announcements.
defineOptions({ inheritAttrs: false })

const props = withDefaults(
  defineProps<{
    /** The presentation component, rendered on each click. */
    document: Component
    documentProps?: Record<string, unknown>
    fileName?: string
    /** Announced to screen readers while the file is built. */
    preparingText?: string
    /** Announced when the file is saved. */
    doneText?: string
  }>(),
  {
    documentProps: undefined,
    fileName: 'presentation.pptx',
    preparingText: 'Preparing presentation…',
    doneText: 'Presentation downloaded',
  },
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
const status = ref('')

const HIDDEN = {
  position: 'absolute',
  width: '1px',
  height: '1px',
  margin: '-1px',
  padding: 0,
  overflow: 'hidden',
  clip: 'rect(0 0 0 0)',
  whiteSpace: 'nowrap',
  border: 0,
} as const

// The file is built on click, not on every change, so editing a large deck stays fast.
const download = async (): Promise<void> => {
  if (loading.value) return
  loading.value = true
  error.value = null
  status.value = props.preparingText
  try {
    emit('done', await downloadPptx(props.document, props.documentProps, props.fileName))
    status.value = props.doneText
  } catch (caught) {
    status.value = ''
    error.value = caught
    emit('error', caught)
  } finally {
    loading.value = false
  }
}
</script>

<template>
  <button
    v-bind="$attrs"
    type="button"
    data-pptx-download
    :aria-busy="loading ? 'true' : undefined"
    :data-loading="loading ? '' : undefined"
    @click="download"
  >
    <slot :loading="loading" :error="error">{{ loading ? 'Preparing…' : 'Download PPTX' }}</slot>
  </button>
  <span data-pptx-download-status role="status" aria-live="polite" :style="HIDDEN">{{ status }}</span>
  <span v-if="error" data-pptx-download-error role="alert" :style="HIDDEN">{{
    error instanceof Error ? error.message : String(error)
  }}</span>
</template>
