<script setup lang="ts">
import { ref } from 'vue'
import { renderToModel } from 'vue-pptx-document'

import LayoutDeck from '../examples/LayoutDeck.vue'
import CodeBlock from '../ui/CodeBlock.vue'
import CopyButton from '../ui/CopyButton.vue'

const code = ref('')

// The flex slide comes out with every frame resolved, as plain absolute positions.
const generate = async (): Promise<void> => {
  const [model, { modelToVue }] = await Promise.all([
    renderToModel(LayoutDeck, { cards: ['One', 'Two', 'Three'], direction: 'row' }),
    import('vue-pptx-document/import'),
  ])
  code.value = modelToVue(model)
}
</script>

<template>
  <div class="stack">
    <div class="actions">
      <button type="button" class="button button-small" @click="generate">Generate Vue code</button>
      <CopyButton v-if="code" :text="code" label="Copy the generated component" />
    </div>
    <CodeBlock v-if="code" class="code-tall" :code="code" file-name="LayoutDeck.generated.vue" />
  </div>
</template>
