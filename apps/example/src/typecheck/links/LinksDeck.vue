<!--
  Generated from apps/docs/docs/guides/links.mdx by scripts/samples.js; edit the source, then run
  `bun run sync:samples`. `bun run typecheck` compiles it against the library.
-->
<script setup lang="ts">
import { Presentation, Run, Shape, Slide, Table, TableCell, TableRow, Text } from 'vue-pptx-document'
</script>

<template>
  <Presentation title="Links">
    <Slide title="Start">
      <Text :x="0.5" :y="0.5" :w="9" :h="0.6">
        See <Run :href="{ slide: 'Pricing' }">the prices</Run> or <Run href="https://vuejs.org">Vue</Run>.
      </Text>
      <Shape type="roundRect" :x="0.5" :y="1.5" :w="3" :h="1" fill="#f35815" :href="{ slide: 2, tooltip: 'Prices' }" />
    </Slide>
    <Slide title="Pricing">
      <Table :x="0.5" :y="1.2" :w="6">
        <TableRow href="https://vuejs.org/guide/">
          <TableCell>Vue guide</TableCell>
          <TableCell href="https://vuejs.org/api/">Own link</TableCell>
        </TableRow>
      </Table>
    </Slide>
  </Presentation>
</template>
