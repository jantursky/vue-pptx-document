<!--
  Generated from packages/vue-pptx-document/README.md by scripts/samples.js; edit the source, then run
  `bun run sync:samples`. `bun run typecheck` compiles it against the library.
-->
<script setup lang="ts">
import { Chart, Presentation, Shape, Slide, Text, View } from 'vue-pptx-document'

defineProps<{ title: string; months: string[]; revenue: number[] }>()
</script>

<template>
  <Presentation :title="title" layout="16x9">
    <!-- Absolute positions: numbers are inches, strings take %, pt, px, cm or mm. -->
    <Slide background="#FAFAFA">
      <Shape type="rect" :x="0" :y="0" w="100%" :h="0.08" fill="#F35815" />
      <Text :x="0.6" :y="2" :w="8.8" :h="1" :font-size="40" bold color="#111111">{{ title }}</Text>
    </Slide>

    <!-- Flexbox: the slide is a column, the chart takes the space the title leaves. -->
    <Slide layout="flex" :style="{ padding: 0.6, gap: 0.3 }">
      <Text :font-size="28" bold>Monthly revenue</Text>
      <View :style="{ flexDirection: 'row', flexGrow: 1, gap: 0.3 }">
        <Chart type="bar" :data="[{ name: 'Revenue', labels: months, values: revenue }]" :style="{ flex: 2 }" />
        <Text :font-size="16" :style="{ flex: 1 }">Up {{ revenue.at(-1) }}k in the last month.</Text>
      </View>
    </Slide>
  </Presentation>
</template>
