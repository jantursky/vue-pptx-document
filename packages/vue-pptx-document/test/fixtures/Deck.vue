<script setup lang="ts">
import { Presentation, Slide, Text } from '../../src'

defineProps<{ title: string; points: string[] }>()
</script>

<template>
  <Presentation :title="title">
    <Slide>
      <Text :x="0.5" :y="0.5" :w="9" :h="1" :font-size="32" bold>{{ title }}</Text>
    </Slide>
    <Slide v-for="point in points" :key="point">
      <Text :x="0.5" :y="0.5" :w="9" :h="1">{{ point }}</Text>
    </Slide>
  </Presentation>
</template>
