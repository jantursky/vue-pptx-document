<script setup lang="ts">
import {
  Master,
  Notes,
  Paragraph,
  Presentation,
  Run,
  Shape,
  Slide,
  Table,
  TableCell,
  TableRow,
  Text,
} from 'vue-pptx-document'

import type { Row } from './report'

defineProps<{ rows: Row[]; quarter: string }>()

const status = { 'On track': '#2A8A62', 'At risk': '#D64545', Done: '#737373' } as const
</script>

<template>
  <Presentation
    title="Quarterly review"
    author="vue-pptx-document"
    layout="16x9"
    head-font-face="Arial"
    body-font-face="Arial"
  >
    <!-- A master: every slide that names it gets the band, the footer and the number. -->
    <Master
      name="BRAND"
      background="#FAFAFA"
      :slide-number="{ x: 9, y: 5.2, w: 0.6, h: 0.3, fontSize: 10, color: '#737373' }"
    >
      <Shape type="rect" :x="0" :y="0" :w="10" :h="0.12" fill="#35495E" />
      <Text :x="0.5" :y="5.2" :w="6" :h="0.3" :font-size="10" color="#737373">Quarterly review · internal</Text>
      <Shape type="line" :x="0.5" :y="5.15" :w="9" :h="0" :line="{ color: '#C1C1C1', width: 0.75 }" />
    </Master>

    <Slide master="BRAND" section="Overview">
      <Shape type="rect" :x="0.5" :y="1.7" :w="0.06" :h="1.2" fill="#42B883" />
      <Text :x="0.75" :y="1.6" :w="8" :h="0.9" :font-size="40" bold valign="top">Quarterly review</Text>
      <Text :x="0.75" :y="2.45" :w="8" :h="0.5" :font-size="18" color="#414141" valign="top"
        >{{ quarter }} · Product teams</Text
      >
      <Notes>Welcome everyone. This deck was rendered from a Vue component.</Notes>
    </Slide>

    <Slide master="BRAND" section="Overview">
      <Text :x="0.5" :y="0.4" :w="9" :h="0.6" :font-size="24" bold valign="top">What changed</Text>
      <Text :x="0.5" :y="1.1" :w="9" :h="3.6" :font-size="16" valign="top" :para-space-after="8">
        <Paragraph bullet
          >Shipped <Run bold>{{ rows.filter((row) => row.status === 'Done').length }}</Run> projects this
          quarter</Paragraph
        >
        <Paragraph bullet>
          <Run :color="status['At risk']" bold>{{ rows.filter((row) => row.status === 'At risk').length }} at risk</Run
          >, with owners named in the table
        </Paragraph>
        <Paragraph :bullet="{ type: 'number' }" :indent-level="1">Review scope in week one</Paragraph>
        <Paragraph :bullet="{ type: 'number' }" :indent-level="1">Re-plan by week three</Paragraph>
        <Paragraph bullet
          >Read the full notes on <Run href="https://github.com/jantursky/vue-pptx-document">GitHub</Run></Paragraph
        >
      </Text>
      <Notes>Keep this one short.</Notes>
    </Slide>

    <Slide master="BRAND" section="Details">
      <Text :x="0.5" :y="0.4" :w="9" :h="0.6" :font-size="24" bold valign="top">Status by team</Text>
      <Table
        :x="0.5"
        :y="1.15"
        :w="9"
        :col-w="[3.5, 3.5, 2]"
        :font-size="12"
        :border="{ type: 'solid', color: 'C1C1C1', pt: 0.75 }"
      >
        <TableRow bold color="#FAFAFA" :fill="{ color: '35495E' }">
          <TableCell>Team</TableCell>
          <TableCell>Owner</TableCell>
          <TableCell>Status</TableCell>
        </TableRow>
        <TableRow v-for="row in rows" :key="row.team">
          <TableCell>{{ row.team }}</TableCell>
          <TableCell>{{ row.owner }}</TableCell>
          <TableCell bold :color="status[row.status]">{{ row.status }}</TableCell>
        </TableRow>
      </Table>
    </Slide>
  </Presentation>
</template>
