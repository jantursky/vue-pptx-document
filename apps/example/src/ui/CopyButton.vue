<script setup lang="ts">
import { ref } from 'vue'

const props = defineProps<{ text: string; label?: string }>()

const copied = ref(false)

const copy = async (): Promise<void> => {
  await navigator.clipboard.writeText(props.text)
  copied.value = true
  setTimeout(() => (copied.value = false), 1500)
}
</script>

<template>
  <button type="button" class="copy" :aria-label="label" @click="copy">{{ copied ? 'Copied' : 'Copy' }}</button>
</template>
