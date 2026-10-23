<script setup lang="ts">
import { Presentation, Shape, Slide, Text, View } from 'vue-pptx-document'

const props = defineProps<{ cards: string[]; direction: 'row' | 'column' }>()

const accent = '#F35815'
</script>

<template>
  <Presentation title="Layout" layout="16x9">
    <!-- Absolute: every object has x, y, w and h in inches. -->
    <Slide background="#FAFAFA">
      <Text :x="0.5" :y="0.35" :w="9" :h="0.6" :font-size="24" bold valign="top">Absolute positions</Text>
      <template v-for="(card, index) in props.cards" :key="card">
        <Shape
          type="rect"
          :x="0.5 + index * 2.3"
          :y="1.4"
          :w="2.1"
          :h="2.4"
          fill="#FFFFFF"
          :line="{ color: '#C1C1C1', width: 1 }"
        />
        <Shape type="rect" :x="0.5 + index * 2.3" :y="1.4" :w="0.06" :h="2.4" :fill="accent" />
        <Text :x="0.7 + index * 2.3" :y="1.55" :w="1.8" :h="2.1" :font-size="14" valign="top">{{ card }}</Text>
      </template>
    </Slide>

    <!-- Flex: Yoga places the same cards, so adding one reflows the row. -->
    <Slide background="#FAFAFA" layout="flex" :style="{ padding: 0.5, gap: 0.3 }">
      <Text :font-size="24" bold valign="top">Flexbox with Yoga</Text>
      <View :style="{ flexDirection: props.direction, flexGrow: 1, gap: 0.2 }">
        <View
          v-for="card in props.cards"
          :key="card"
          fill="#FFFFFF"
          :line="{ color: '#C1C1C1', width: 1 }"
          :style="{ flex: 1, flexDirection: 'row' }"
        >
          <Shape type="rect" :fill="accent" :style="{ width: 0.06 }" />
          <Text :font-size="14" valign="top" :style="{ flex: 1 }">{{ card }}</Text>
        </View>
      </View>
    </Slide>
  </Presentation>
</template>
