<script setup lang="ts">
import { Presentation, Shape, Slide, Text } from 'vue-pptx-document'

defineProps<{ name: string }>()
</script>

<template>
  <Presentation title="Hello" layout="16x9">
    <Slide background="#FAFAFA">
      <Shape type="rect" :x="0.6" :y="1.9" :w="0.06" :h="1.1" fill="#F35815" />
      <Text :x="0.85" :y="1.8" :w="8.5" :h="0.8" :font-size="40" bold color="#000000" valign="top">
        Hello, {{ name || 'world' }}
      </Text>
      <Text :x="0.85" :y="2.6" :w="8.5" :h="0.5" :font-size="18" color="#414141" valign="top">
        A slide written as a Vue component.
      </Text>
    </Slide>
  </Presentation>
</template>
