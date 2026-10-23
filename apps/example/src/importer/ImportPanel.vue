<script setup lang="ts">
import { computed, ref, shallowRef } from 'vue'
import { renderToBuffer, saveBlob, SlidePreview, type DeckModel } from 'vue-pptx-document'

import { sampleReport } from '../generator/model'
import ReportTemplate from '../generator/ReportTemplate.vue'
import CodeBlock from '../ui/CodeBlock.vue'
import CopyButton from '../ui/CopyButton.vue'

const model = shallowRef<DeckModel | null>(null)
const code = ref('')
const warnings = ref<string[]>([])
const error = ref('')
const busy = ref(false)
const fileName = ref('')
const dragging = ref(false)
const embedImages = ref(true)
const masterShapes = ref(true)

const componentName = computed(
  () =>
    (fileName.value.replace(/\.pptx$/i, '') || 'Imported')
      .split(/[^\p{L}\p{N}]+/u)
      .filter(Boolean)
      .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
      .join('') + '.vue',
)

// The importer is its own entry point, loaded only when a file arrives.
const load = async (bytes: ArrayBuffer | Uint8Array, name: string): Promise<void> => {
  busy.value = true
  error.value = ''
  try {
    const { pptxToVue } = await import('vue-pptx-document/import')
    const result = await pptxToVue(bytes, { embedImages: embedImages.value, masterShapes: masterShapes.value })
    model.value = result.model
    code.value = result.code
    warnings.value = result.warnings
    fileName.value = name
  } catch (caught) {
    error.value = caught instanceof Error ? caught.message : String(caught)
    model.value = null
    code.value = ''
  } finally {
    busy.value = false
  }
}

const onFile = async (file: File | undefined): Promise<void> => {
  if (!file) return
  await load(await file.arrayBuffer(), file.name)
}

const onInput = (event: Event): void => {
  const input = event.target as HTMLInputElement
  void onFile(input.files?.[0])
  input.value = ''
}

const onDrop = (event: DragEvent): void => {
  dragging.value = false
  void onFile(event.dataTransfer?.files[0])
}

// A round trip: render the generator's sample deck with PptxGenJS, then read it back.
const loadSample = async (): Promise<void> => {
  busy.value = true
  const bytes = await renderToBuffer(ReportTemplate, { report: sampleReport() })
  await load(bytes, 'sample-report.pptx')
}

const download = (): void => saveBlob(new Blob([code.value], { type: 'text/plain' }), componentName.value)
</script>

<template>
  <div class="importer">
    <div
      class="dropzone"
      :data-dragging="dragging ? '' : undefined"
      @dragover.prevent="dragging = true"
      @dragleave="dragging = false"
      @drop.prevent="onDrop"
    >
      <p><strong>Drop a .pptx file here</strong> or pick one. It is read in this browser and never uploaded.</p>
      <div class="actions">
        <label class="button button-primary file-button">
          Choose a file
          <input
            type="file"
            accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
            data-field="pptx"
            @change="onInput"
          />
        </label>
        <button type="button" class="button" :disabled="busy" data-import-sample @click="loadSample">
          Try the sample deck
        </button>
      </div>
      <div class="options">
        <label class="check"><input v-model="embedImages" type="checkbox" /> Embed pictures as data URIs</label>
        <label class="check"
          ><input v-model="masterShapes" type="checkbox" /> Copy layout and master shapes onto slides</label
        >
      </div>
    </div>

    <p v-if="busy" class="note" role="status">Reading…</p>
    <p v-if="error" class="note note-error" role="alert">{{ error }}</p>

    <template v-if="model && !busy">
      <div class="import-summary" role="status" data-import-summary>
        <strong>{{ fileName }}</strong
        >: {{ model.slides.length }} slides, {{ model.layout.width }} × {{ model.layout.height }} in
      </div>
      <ul v-if="warnings.length" class="warnings" aria-label="Import warnings">
        <li v-for="warning in warnings" :key="warning">{{ warning }}</li>
      </ul>
      <div class="import-grid">
        <section aria-label="Slides read from the file">
          <ol class="thumbs">
            <li v-for="(_, index) in model.slides" :key="index">
              <SlidePreview :model="model" :index="index" />
              <span>{{ index + 1 }}</span>
            </li>
          </ol>
        </section>
        <section class="generated" aria-label="Generated Vue component">
          <div class="source-bar">
            <span>{{ componentName }}</span>
            <span class="actions">
              <CopyButton :text="code" label="Copy the generated component" />
              <button type="button" class="copy" @click="download">Download .vue</button>
            </span>
          </div>
          <CodeBlock class="code-tall" :code="code" :file-name="componentName" />
        </section>
      </div>
    </template>
  </div>
</template>
