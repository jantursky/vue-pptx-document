<script setup lang="ts">
import { computed } from 'vue'
import {
  Chart,
  Image,
  LAYOUTS,
  Master,
  Notes,
  Paragraph,
  Presentation,
  Shape,
  Slide,
  Table,
  Text,
  View,
} from 'vue-pptx-document'

import { labels as t } from './labels'
import type { Report } from './model'

const props = defineProps<{ report: Report }>()

const size = computed(() => LAYOUTS[props.report.layout])
const number = (value: number) => new Intl.NumberFormat('en-US').format(value)
const date = computed(() => {
  const parsed = new Date(props.report.date)
  return Number.isNaN(parsed.getTime()) ? props.report.date : parsed.toLocaleDateString('en-US', { dateStyle: 'long' })
})

const ink = '#111111'
const body = '#414141'
const muted = '#737373'
const hairline = '#C1C1C1'
const paper = '#FAFAFA'
const positive = '#0B6EC5'

const highlights = computed(() => props.report.highlights.filter((item) => item.trim() !== ''))
const kpis = computed(() => props.report.kpis.filter((kpi) => kpi.label.trim() !== '' || kpi.value.trim() !== ''))
const series = computed(() =>
  props.report.series
    .filter((item) => item.name.trim() !== '')
    .map((item) => ({
      name: item.name,
      labels: props.report.months,
      values: props.report.months.map((_, index) => item.values[index] ?? 0),
    })),
)
const regions = computed(() => props.report.regions.filter((region) => region.name.trim() !== ''))
const rows = computed(() => [
  [t.region, t.revenue, t.target, t.ofTarget].map((text) => ({
    text,
    options: { bold: true, color: paper, fill: { color: ink } },
  })),
  ...regions.value.map((region) => {
    const ratio = region.target ? region.revenue / region.target : 0
    return [
      region.name,
      { text: number(region.revenue), options: { align: 'right' as const } },
      { text: number(region.target), options: { align: 'right' as const } },
      {
        text: `${Math.round(ratio * 100)} %`,
        options: { align: 'right' as const, bold: true, color: ratio >= 1 ? positive : props.report.accent },
      },
    ]
  }),
])
const share = computed(() => [
  {
    name: t.share,
    labels: regions.value.map((region) => region.name),
    values: regions.value.map((region) => region.revenue),
  },
])
const palette = computed(() => [props.report.accent, ink, muted, hairline, positive, '#F2B600'])
const footer = computed(() => [props.report.company, props.report.title].filter(Boolean).join(' · '))
</script>

<template>
  <Presentation
    :layout="report.layout"
    :title="report.title"
    :author="report.author"
    :company="report.company"
    head-font-face="Arial"
    body-font-face="Arial"
  >
    <Master
      name="REPORT"
      :background="paper"
      :slide-number="{
        x: size.width - 1.1,
        y: size.height - 0.45,
        w: 0.6,
        h: 0.3,
        fontSize: 9,
        color: muted,
        align: 'right',
      }"
    >
      <Shape type="rect" :x="0" :y="0" :w="size.width" :h="0.08" :fill="report.accent" />
      <Shape
        type="line"
        :x="0.5"
        :y="size.height - 0.5"
        :w="size.width - 1"
        :h="0"
        :line="{ color: hairline, width: 0.75 }"
      />
      <Text :x="0.5" :y="size.height - 0.45" :w="size.width - 2" :h="0.3" :font-size="9" :color="muted" :margin="0">{{
        footer
      }}</Text>
    </Master>

    <!-- Title: absolute positions, measured from the slide size. -->
    <Slide :background="paper">
      <Shape type="rect" :x="0" :y="0" :w="size.width" :h="0.08" :fill="report.accent" />
      <Image
        v-if="report.logo"
        :src="report.logo"
        :x="size.width - 1.7"
        :y="0.45"
        :w="1.2"
        :h="0.6"
        :sizing="{ type: 'contain', w: 1.2, h: 0.6 }"
      />
      <Shape type="rect" :x="0.6" :y="size.height / 2 - 0.85" :w="0.06" :h="1.3" :fill="report.accent" />
      <Text
        :x="0.85"
        :y="size.height / 2 - 0.95"
        :w="size.width - 1.7"
        :h="0.9"
        :font-size="40"
        bold
        :color="ink"
        valign="bottom"
        :margin="0"
      >
        {{ report.title }}
      </Text>
      <Text
        :x="0.85"
        :y="size.height / 2 + 0.05"
        :w="size.width - 1.7"
        :h="0.5"
        :font-size="18"
        :color="body"
        valign="top"
        :margin="0"
      >
        {{ report.subtitle }}
      </Text>
      <Text
        :x="0.85"
        :y="size.height - 1"
        :w="size.width - 1.7"
        :h="0.4"
        :font-size="12"
        :color="muted"
        valign="top"
        :margin="0"
      >
        {{ [report.company, report.author, date].filter(Boolean).join(' · ') }}
      </Text>
      <Notes v-if="report.notes">{{ t.notesTitle }}</Notes>
    </Slide>

    <!-- Highlights: flex, so the list fills whatever space the layout leaves. -->
    <Slide
      v-if="highlights.length"
      master="REPORT"
      layout="flex"
      :style="{ padding: 0.6, paddingBottom: 0.8, gap: 0.3 }"
    >
      <Text :font-size="28" bold :color="ink" valign="top" :margin="0">{{ t.highlights }}</Text>
      <Text :font-size="18" :color="body" valign="top" :para-space-after="10" :margin="0" :style="{ flexGrow: 1 }">
        <Paragraph v-for="item in highlights" :key="item" bullet>{{ item }}</Paragraph>
      </Text>
      <Notes v-if="report.notes">{{ t.notesHighlights }}</Notes>
    </Slide>

    <!-- Metrics: one card per KPI in a flex row. -->
    <Slide v-if="kpis.length" master="REPORT" layout="flex" :style="{ padding: 0.6, paddingBottom: 0.8, gap: 0.3 }">
      <Text :font-size="28" bold :color="ink" valign="top" :margin="0">{{ t.metrics }}</Text>
      <View :style="{ flexDirection: 'row', gap: 0.25, flexGrow: 1, alignItems: 'center' }">
        <View
          v-for="kpi in kpis"
          :key="kpi.label"
          fill="#FFFFFF"
          :line="{ color: hairline, width: 0.75 }"
          :style="{ flex: 1, height: 1.9, padding: 0.2, gap: 0.05 }"
        >
          <Text :font-size="12" :color="muted" valign="top" :margin="0">{{ kpi.label }}</Text>
          <Text :font-size="32" bold :color="ink" valign="middle" :margin="0" :style="{ flexGrow: 1 }">{{
            kpi.value
          }}</Text>
          <Text :font-size="12" bold :color="kpi.delta >= 0 ? positive : report.accent" valign="bottom" :margin="0">
            {{ kpi.delta >= 0 ? '▲' : '▼' }} {{ Math.abs(kpi.delta) }} % {{ t.vsLast }}
          </Text>
        </View>
      </View>
      <Notes v-if="report.notes">{{ t.notesMetrics }}</Notes>
    </Slide>

    <Slide v-if="series.length" master="REPORT" layout="flex" :style="{ padding: 0.6, paddingBottom: 0.8, gap: 0.2 }">
      <Text :font-size="28" bold :color="ink" valign="top" :margin="0">{{ t.trend }}</Text>
      <Chart
        :type="report.chart"
        :data="series"
        :chart-colors="palette"
        show-legend
        legend-pos="b"
        :legend-font-size="11"
        :cat-axis-label-font-size="11"
        :val-axis-label-font-size="11"
        :val-grid-line="{ color: 'E5E5E5', size: 1 }"
        :line-size="3"
        :line-data-symbol-size="8"
        show-val-axis-title
        :val-axis-title="`${t.revenue} (${t.thousands})`"
        :val-axis-title-font-size="10"
        :style="{ flexGrow: 1 }"
      />
      <Notes v-if="report.notes">{{ t.notesTrend }}</Notes>
    </Slide>

    <Slide v-if="regions.length" master="REPORT" layout="flex" :style="{ padding: 0.6, paddingBottom: 0.8, gap: 0.3 }">
      <Text :font-size="28" bold :color="ink" valign="top" :margin="0">{{ t.regions }}</Text>
      <View :style="{ flexDirection: 'row', gap: 0.4, flexGrow: 1 }">
        <Table
          :rows="rows"
          :font-size="12"
          :color="body"
          :border="{ type: 'solid', color: 'C1C1C1', pt: 0.75 }"
          :style="{ flex: 3, alignSelf: 'flex-start' }"
        />
        <Chart
          type="doughnut"
          :data="share"
          :chart-colors="palette"
          :hole-size="58"
          show-percent
          show-legend
          legend-pos="b"
          :legend-font-size="10"
          :data-label-font-size="10"
          :data-label-color="paper"
          :style="{ flex: 2 }"
        />
      </View>
      <Notes v-if="report.notes">{{ t.notesRegions }}</Notes>
    </Slide>

    <Slide :background="ink">
      <Shape type="rect" :x="0.6" :y="size.height / 2 - 0.8" :w="0.06" :h="1.6" :fill="report.accent" />
      <Text
        :x="0.85"
        :y="size.height / 2 - 0.85"
        :w="size.width - 1.7"
        :h="0.8"
        :font-size="36"
        bold
        :color="paper"
        valign="top"
        :margin="0"
      >
        {{ report.closing ? t.next : t.thanks }}
      </Text>
      <Text
        :x="0.85"
        :y="size.height / 2"
        :w="size.width - 1.7"
        :h="0.9"
        :font-size="18"
        color="#C1C1C1"
        valign="top"
        :margin="0"
      >
        {{ report.closing || report.company }}
      </Text>
      <Notes v-if="report.notes">{{ t.notesClosing }}</Notes>
    </Slide>
  </Presentation>
</template>
