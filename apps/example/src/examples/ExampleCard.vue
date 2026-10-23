<script setup lang="ts">
import { computed, ref } from 'vue'

import CodeBlock from '../ui/CodeBlock.vue'
import CopyButton from '../ui/CopyButton.vue'

export interface SourceFile {
  name: string
  code: string
}

const props = defineProps<{ id: string; title: string; files: readonly SourceFile[] }>()

const showSource = ref(false)
const active = ref(0)
const file = computed(() => props.files[active.value] ?? props.files[0])
</script>

<template>
  <section :id="id" class="card example">
    <header class="example-head">
      <div>
        <h3>{{ title }}</h3>
        <slot name="description" />
      </div>
      <button
        type="button"
        class="button button-small"
        :aria-expanded="showSource"
        :aria-controls="`${id}-source`"
        @click="showSource = !showSource"
      >
        {{ showSource ? 'Hide source' : 'View source' }}
      </button>
    </header>

    <div v-if="showSource && file" :id="`${id}-source`" class="source">
      <div class="source-bar">
        <div role="tablist" :aria-label="`${title} source files`">
          <button
            v-for="(entry, index) in files"
            :key="entry.name"
            type="button"
            role="tab"
            :aria-selected="index === active"
            @click="active = index"
          >
            {{ entry.name }}
          </button>
        </div>
        <CopyButton :text="file.code" :label="`Copy ${file.name}`" />
      </div>
      <div role="tabpanel">
        <CodeBlock :code="file.code" :file-name="file.name" />
      </div>
    </div>

    <div class="example-live">
      <slot />
    </div>
  </section>
</template>
