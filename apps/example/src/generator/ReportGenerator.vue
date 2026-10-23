<script setup lang="ts">
import { computed, reactive, ref, toRaw, watch } from 'vue'
import { PptxDownloadButton, Preview, renderToModel, saveBlob } from 'vue-pptx-document'

import CodeBlock from '../ui/CodeBlock.vue'
import CopyButton from '../ui/CopyButton.vue'
import { ui as t } from './labels'
import {
  emptyReport,
  isReport,
  loadDraft,
  MAX_KPIS,
  MAX_REGIONS,
  MAX_SERIES,
  sampleReport,
  saveDraft,
  type ChartKind,
  type DeckLayout,
  type Report,
} from './model'
import ReportTemplate from './ReportTemplate.vue'

const report = reactive<Report>(loadDraft() ?? sampleReport())
const message = ref('')

watch(report, () => saveDraft(report), { deep: true })

const replace = (next: Report): void => {
  Object.assign(report, structuredClone(next))
  message.value = ''
}

const documentProps = computed(() => ({ report }))
const fileBase = computed(() =>
  (report.title.trim() || 'report')
    .replace(/[^\p{L}\p{N}]+/gu, '-')
    .replace(/^-|-$/g, '')
    .toLowerCase(),
)

const months = computed({
  get: () => report.months.join(', '),
  set: (value: string) => {
    report.months = value
      .split(',')
      .map((month) => month.trim())
      .filter(Boolean)
      .slice(0, 12)
  },
})

const move = (list: unknown[], index: number, step: number): void => {
  const target = index + step
  if (target < 0 || target >= list.length) return
  const [item] = list.splice(index, 1)
  if (item !== undefined) list.splice(target, 0, item)
}

const onLogo = (event: Event): void => {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  input.value = ''
  if (!file) return
  if (file.size > 1.5 * 1024 * 1024) {
    message.value = t.logoTooBig
    return
  }
  const reader = new FileReader()
  reader.onload = () => {
    if (typeof reader.result === 'string') report.logo = reader.result
  }
  reader.readAsDataURL(file)
}

const exportJson = (): void => {
  saveBlob(new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' }), `${fileBase.value}.json`)
}

const importJson = async (event: Event): Promise<void> => {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  input.value = ''
  if (!file) return
  try {
    const parsed: unknown = JSON.parse(await file.text())
    if (!isReport(parsed)) throw new Error('invalid')
    replace(parsed)
  } catch {
    message.value = t.invalidJson
  }
}

// "Export Vue code": the deck as it is now, written out as a standalone component.
const code = ref('')
const exporting = ref(false)
const exportVue = async (): Promise<void> => {
  if (code.value) {
    code.value = ''
    return
  }
  exporting.value = true
  try {
    const [model, { modelToVue }] = await Promise.all([
      renderToModel(ReportTemplate, { report: structuredClone(toRaw(report)) }),
      import('vue-pptx-document/import'),
    ])
    code.value = modelToVue(model)
  } finally {
    exporting.value = false
  }
}
watch(report, () => (code.value = ''), { deep: true })
const downloadVue = (): void => {
  const name = fileBase.value
    .split('-')
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join('')
  saveBlob(new Blob([code.value], { type: 'text/plain' }), `${name}Deck.vue`)
}

const layouts: DeckLayout[] = ['16x9', '4x3', 'wide']
const charts: ChartKind[] = ['bar', 'line', 'area']
</script>

<template>
  <div class="generator">
    <form class="generator-form" @submit.prevent>
      <fieldset class="card">
        <legend>{{ t.deck }}</legend>
        <label class="field">
          <span>{{ t.layout }}</span>
          <select v-model="report.layout" data-field="layout">
            <option v-for="layout in layouts" :key="layout" :value="layout">{{ layout }}</option>
          </select>
        </label>
        <label class="field">
          <span>{{ t.title }}</span>
          <input v-model="report.title" type="text" data-field="title" />
        </label>
        <label class="field">
          <span>{{ t.subtitle }}</span>
          <input v-model="report.subtitle" type="text" />
        </label>
        <div class="grid-2">
          <label class="field">
            <span>{{ t.company }}</span>
            <input v-model="report.company" type="text" />
          </label>
          <label class="field">
            <span>{{ t.author }}</span>
            <input v-model="report.author" type="text" />
          </label>
          <label class="field">
            <span>{{ t.date }}</span>
            <input v-model="report.date" type="date" />
          </label>
          <label class="field">
            <span>{{ t.accent }}</span>
            <input v-model="report.accent" type="color" data-field="accent" />
          </label>
        </div>
        <div class="logo-row">
          <label class="field">
            <span>{{ t.logo }}</span>
            <input type="file" accept="image/png,image/jpeg" data-field="logo" @change="onLogo" />
          </label>
          <img v-if="report.logo" class="logo-thumb" :src="report.logo" alt="" />
          <button v-if="report.logo" type="button" class="button button-small" @click="report.logo = null">
            {{ t.removeLogo }}
          </button>
        </div>
      </fieldset>

      <fieldset class="card">
        <legend>{{ t.highlights }}</legend>
        <div v-for="(_, index) in report.highlights" :key="index" class="row">
          <input v-model="report.highlights[index]" type="text" :aria-label="`${t.highlights} ${index + 1}`" />
          <button
            type="button"
            class="icon-button"
            :aria-label="`${t.remove} ${index + 1}`"
            @click="report.highlights.splice(index, 1)"
          >
            ×
          </button>
        </div>
        <button
          type="button"
          class="button button-small"
          :disabled="report.highlights.length >= 8"
          @click="report.highlights.push('')"
        >
          {{ t.addHighlight }}
        </button>
      </fieldset>

      <fieldset class="card">
        <legend>{{ t.kpis }}</legend>
        <div class="table-head kpi-row" aria-hidden="true">
          <span>{{ t.label }}</span
          ><span>{{ t.value }}</span
          ><span>{{ t.delta }}</span
          ><span />
        </div>
        <div v-for="(kpi, index) in report.kpis" :key="index" class="kpi-row">
          <input v-model="kpi.label" type="text" :aria-label="`${t.label} ${index + 1}`" />
          <input v-model="kpi.value" type="text" :aria-label="`${t.value} ${index + 1}`" />
          <input v-model.number="kpi.delta" type="number" step="1" :aria-label="`${t.delta} ${index + 1}`" />
          <button
            type="button"
            class="icon-button"
            :aria-label="`${t.remove} ${kpi.label || index + 1}`"
            @click="report.kpis.splice(index, 1)"
          >
            ×
          </button>
        </div>
        <button
          type="button"
          class="button button-small"
          :disabled="report.kpis.length >= MAX_KPIS"
          @click="report.kpis.push({ label: '', value: '', delta: 0 })"
        >
          {{ t.addKpi }}
        </button>
      </fieldset>

      <fieldset class="card">
        <legend>{{ t.chart }}</legend>
        <div class="grid-2">
          <label class="field">
            <span>{{ t.chartType }}</span>
            <select v-model="report.chart" data-field="chart">
              <option v-for="chart in charts" :key="chart" :value="chart">{{ chart }}</option>
            </select>
          </label>
          <label class="field">
            <span>{{ t.months }}</span>
            <input v-model.lazy="months" type="text" data-field="months" />
          </label>
        </div>
        <div v-for="(series, index) in report.series" :key="index" class="series">
          <div class="row">
            <input v-model="series.name" type="text" :aria-label="`${t.series} ${index + 1}`" />
            <button
              type="button"
              class="icon-button"
              :aria-label="`${t.remove} ${series.name || index + 1}`"
              @click="report.series.splice(index, 1)"
            >
              ×
            </button>
          </div>
          <div class="values">
            <label v-for="(month, monthIndex) in report.months" :key="month" class="field field-compact">
              <span>{{ month }}</span>
              <input v-model.number="series.values[monthIndex]" type="number" min="0" />
            </label>
          </div>
        </div>
        <button
          type="button"
          class="button button-small"
          :disabled="report.series.length >= MAX_SERIES"
          @click="report.series.push({ name: '', values: report.months.map(() => 0) })"
        >
          {{ t.addSeries }}
        </button>
      </fieldset>

      <fieldset class="card">
        <legend>{{ t.regions }}</legend>
        <div class="table-head region-row" aria-hidden="true">
          <span>{{ t.label }}</span
          ><span>{{ t.revenue }}</span
          ><span>{{ t.target }}</span
          ><span />
        </div>
        <div v-for="(region, index) in report.regions" :key="index" class="region-row">
          <input v-model="region.name" type="text" :aria-label="`${t.regions} ${index + 1}`" />
          <input v-model.number="region.revenue" type="number" min="0" :aria-label="`${t.revenue} ${region.name}`" />
          <input v-model.number="region.target" type="number" min="0" :aria-label="`${t.target} ${region.name}`" />
          <span class="row-actions">
            <button
              type="button"
              class="icon-button"
              :aria-label="`${t.up}: ${region.name}`"
              :disabled="index === 0"
              @click="move(report.regions, index, -1)"
            >
              ↑
            </button>
            <button
              type="button"
              class="icon-button"
              :aria-label="`${t.down}: ${region.name}`"
              :disabled="index === report.regions.length - 1"
              @click="move(report.regions, index, 1)"
            >
              ↓
            </button>
            <button
              type="button"
              class="icon-button"
              :aria-label="`${t.remove} ${region.name || index + 1}`"
              @click="report.regions.splice(index, 1)"
            >
              ×
            </button>
          </span>
        </div>
        <button
          type="button"
          class="button button-small"
          :disabled="report.regions.length >= MAX_REGIONS"
          @click="report.regions.push({ name: '', revenue: 0, target: 0 })"
        >
          {{ t.addRegion }}
        </button>
      </fieldset>

      <fieldset class="card">
        <legend>{{ t.closing }}</legend>
        <label class="field">
          <span>{{ t.closing }}</span>
          <textarea v-model="report.closing" rows="2" />
        </label>
        <label class="check">
          <input v-model="report.notes" type="checkbox" />
          {{ t.notes }}
        </label>
      </fieldset>
    </form>

    <aside class="generator-preview" :aria-label="t.preview">
      <div class="toolbar">
        <PptxDownloadButton
          class="button button-primary"
          :document="ReportTemplate"
          :document-props="documentProps"
          :file-name="`${fileBase}.pptx`"
        >
          <template #default="{ loading }">{{ loading ? '…' : t.download }}</template>
        </PptxDownloadButton>
        <button
          type="button"
          class="button"
          :aria-expanded="code !== ''"
          :disabled="exporting"
          data-export-vue
          @click="exportVue"
        >
          {{ code ? t.hideVue : t.exportVue }}
        </button>
        <button type="button" class="button button-small" @click="exportJson">{{ t.exportJson }}</button>
        <label class="button button-small file-button">
          {{ t.importJson }}
          <input type="file" accept="application/json,.json" data-field="import" @change="importJson" />
        </label>
        <button type="button" class="button button-small" @click="replace(sampleReport())">{{ t.sample }}</button>
        <button type="button" class="button button-small" @click="replace(emptyReport())">{{ t.clear }}</button>
      </div>
      <p v-if="message" class="note note-error" role="alert">{{ message }}</p>
      <div v-if="code" class="generated">
        <div class="source-bar">
          <span>{{ fileBase }}.vue</span>
          <span class="actions">
            <CopyButton :text="code" label="Copy the generated component" />
            <button type="button" class="copy" @click="downloadVue">{{ t.downloadVue }}</button>
          </span>
        </div>
        <CodeBlock class="code-tall" :code="code" file-name="Deck.vue" />
      </div>
      <Preview
        class="viewer viewer-list"
        tabindex="0"
        :document="ReportTemplate"
        :document-props="documentProps"
        :notes="report.notes"
        :debounce="120"
      >
        <template #caption="{ index, total }">
          <span data-status>{{ index + 1 }} / {{ total }}</span>
        </template>
      </Preview>
    </aside>
  </div>
</template>
