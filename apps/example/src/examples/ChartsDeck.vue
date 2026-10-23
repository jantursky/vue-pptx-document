<script setup lang="ts">
import { computed } from 'vue'
import { Chart, Presentation, Slide, Text, type ChartSeries, type ComboChart } from 'vue-pptx-document'

const props = defineProps<{ seed: number }>()

const colors = ['#42B883', '#35495E', '#9FE0BF', '#2A8A62', '#737373']
const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']

// A small deterministic generator, so the same seed draws the same charts.
const random = (seed: number) => {
  let state = seed * 9301 + 49297
  return () => {
    state = (state * 9301 + 49297) % 233280
    return state / 233280
  }
}

const data = computed(() => {
  const next = random(props.seed)
  const series = (name: string): ChartSeries => ({
    name,
    labels: months,
    values: months.map(() => Math.round(20 + next() * 80)),
  })
  const xs = Array.from({ length: 8 }, () => Math.round(next() * 100))
  const north = series('North')
  const south = series('South')
  return {
    north,
    south,
    two: [north, south],
    three: [north, south, series('West')],
    share: [{ name: 'Share', labels: ['Web', 'Mobile', 'Desktop', 'Other'], values: [44, 31, 17, 8] }],
    scatter: [
      { name: 'X', values: xs },
      { name: 'Accounts', values: xs.map((x) => Math.round(x * 0.6 + next() * 30)) },
    ],
    bubble: [
      { name: 'X', values: xs.slice(0, 5) },
      {
        name: 'Markets',
        values: xs.slice(0, 5).map(() => Math.round(next() * 100)),
        sizes: xs.slice(0, 5).map(() => Math.round(4 + next() * 20)),
      },
    ],
  }
})

const combo = computed<ComboChart[]>(() => [
  { type: 'bar', data: [data.value.north], options: { chartColors: ['#C1C1C1'] } },
  {
    type: 'line',
    data: [data.value.south],
    options: { chartColors: ['#42B883'], secondaryValAxis: true, secondaryCatAxis: true },
  },
])

const charts = computed(
  () =>
    [
      { title: 'Clustered columns', type: 'bar', data: data.value.two, options: { barDir: 'col' } },
      {
        title: 'Stacked bars',
        type: 'bar',
        data: data.value.three,
        options: { barDir: 'bar', barGrouping: 'stacked' },
      },
      { title: '3D columns', type: 'bar3D', data: data.value.two, options: {} },
      { title: 'Lines', type: 'line', data: data.value.two, options: { lineSmooth: true } },
      { title: 'Stacked area', type: 'area', data: data.value.three, options: { barGrouping: 'stacked' } },
      { title: 'Pie', type: 'pie', data: data.value.share, options: { showPercent: true } },
      { title: 'Doughnut', type: 'doughnut', data: data.value.share, options: { holeSize: 60, showPercent: true } },
      { title: 'Scatter', type: 'scatter', data: data.value.scatter, options: { lineSize: 0, lineDataSymbolSize: 8 } },
      { title: 'Bubbles', type: 'bubble', data: data.value.bubble, options: {} },
      { title: 'Radar', type: 'radar', data: data.value.two, options: { radarStyle: 'filled' } },
    ] as const,
)
</script>

<template>
  <Presentation title="Charts" layout="16x9">
    <Slide v-for="chart in charts" :key="chart.title" background="#FAFAFA">
      <Text :x="0.5" :y="0.3" :w="9" :h="0.6" :font-size="24" bold valign="top">{{ chart.title }}</Text>
      <Chart
        :type="chart.type"
        :data="chart.data"
        :x="0.5"
        :y="1"
        :w="9"
        :h="4.3"
        :chart-colors="colors"
        show-legend
        legend-pos="b"
        :cat-axis-label-font-size="11"
        :val-axis-label-font-size="11"
        :legend-font-size="11"
        :val-grid-line="{ color: 'E5E5E5', size: 1 }"
        v-bind="chart.options"
      />
    </Slide>
    <Slide background="#FAFAFA">
      <Text :x="0.5" :y="0.3" :w="9" :h="0.6" :font-size="24" bold valign="top">Combo: bars and a line</Text>
      <Chart
        :type="combo"
        :x="0.5"
        :y="1"
        :w="9"
        :h="4.3"
        show-legend
        legend-pos="b"
        :legend-font-size="11"
        :val-axes="[{ showValAxisTitle: false }, { showValAxisTitle: false }]"
        :cat-axes="[{}, { catAxisHidden: true }]"
      />
    </Slide>
  </Presentation>
</template>
