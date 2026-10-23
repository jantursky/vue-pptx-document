<p align="center"><img src="https://raw.githubusercontent.com/jantursky/vue-pptx-document/main/apps/docs/public/icon.svg" alt="" width="64" height="64" /></p>

# vue-pptx-document

[![npm](https://img.shields.io/npm/v/vue-pptx-document)](https://www.npmjs.com/package/vue-pptx-document) [![CI](https://github.com/jantursky/vue-pptx-document/actions/workflows/ci.yml/badge.svg)](https://github.com/jantursky/vue-pptx-document/actions/workflows/ci.yml) [![types](https://img.shields.io/npm/types/vue-pptx-document)](https://www.npmjs.com/package/vue-pptx-document) [![license](https://img.shields.io/npm/l/vue-pptx-document)](https://github.com/jantursky/vue-pptx-document/blob/main/LICENSE) [![codecov](https://codecov.io/gh/jantursky/vue-pptx-document/graph/badge.svg)](https://codecov.io/gh/jantursky/vue-pptx-document)

[Docs](https://vue-pptx-document-docs.vercel.app) · [Demo](https://vue-pptx-document.vercel.app) · [Changelog](https://github.com/jantursky/vue-pptx-document/blob/main/packages/vue-pptx-document/CHANGELOG.md) · [GitHub](https://github.com/jantursky/vue-pptx-document)

Create PowerPoint files with Vue 3 components. Write `<Presentation>`, `<Slide>`, `<Text>`, `<Shape>`, `<Image>`, `<Chart>` and `<Table>` in a template, with props, `v-for`, `v-if` and computed values, and get an editable .pptx, in the browser or in Node.

Under the hood it's a custom Vue renderer (`createRenderer`) that builds a plain tree of slides and objects. The tree is resolved to frames in inches, laid out with [Yoga](https://yogalayout.dev) where you ask for flexbox, and handed to [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) to write the file. The same resolved tree drives `<Preview>`, an HTML and SVG rendering of the slides, so you can show the deck as people edit it without writing a file.

## Install

```sh
bun add vue-pptx-document
# or: npm install vue-pptx-document, pnpm add vue-pptx-document, yarn add vue-pptx-document
```

It requires Vue `^3.5` and, on the server, Node.js 20.19 or newer. It is ESM only. PptxGenJS and Yoga load on first use and add about 165 kB gzipped to a browser bundle; the preview alone doesn't need PptxGenJS.

## Quick start

```vue
<!-- ReportDeck.vue -->
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
```

```vue
<!-- App.vue -->
<script setup lang="ts">
import { reactive } from 'vue'
import { PptxDownloadButton, Preview } from 'vue-pptx-document'

import ReportDeck from './ReportDeck.vue'

const report = reactive({ title: 'Q3 review', months: ['Jul', 'Aug', 'Sep'], revenue: [720, 790, 890] })
</script>

<template>
  <input v-model="report.title" />
  <Preview :document="ReportDeck" :document-props="report" />
  <PptxDownloadButton :document="ReportDeck" :document-props="report" file-name="report.pptx" />
</template>
```

The preview follows every change to the data, debounced. The file is written only when someone clicks the button.

In Node, render straight to bytes:

```ts
import { writeFile } from 'node:fs/promises'
import { renderToBuffer } from 'vue-pptx-document'

import ReportDeck from './ReportDeck.vue'

await writeFile(
  'report.pptx',
  await renderToBuffer(ReportDeck, { title: 'Q3 review', months: ['Jul'], revenue: [720] }),
)
```

## Import a .pptx as Vue code

`vue-pptx-document/import` reads a PowerPoint file (text with its runs and bullets, shapes, pictures, tables, charts, backgrounds, sections and notes) and writes a component that renders the same slides:

```ts
import { pptxToVue } from 'vue-pptx-document/import'

const { code, warnings } = await pptxToVue(await file.arrayBuffer())
```

The same importer runs from the command line:

```sh
npx vue-pptx-document import deck.pptx -o Deck.vue
```

## API

| export                                                                 | what it does                                                                                 |
| ---------------------------------------------------------------------- | -------------------------------------------------------------------------------------------- |
| `Presentation`, `Slide`, `Master`, `Placeholder`, `Notes`              | The deck: layout and metadata, slides with backgrounds and sections, masters, speaker notes. |
| `Text`, `Paragraph`, `Run`, `Br`                                       | Text boxes, with paragraphs (bullets, numbering), styled runs and links.                     |
| `Shape`, `Image`, `Media`, `Chart`, `Table`, `TableRow`, `TableCell`   | The PptxGenJS objects. Every PptxGenJS option passes through as a prop.                      |
| `View`                                                                 | A flexbox container, with an optional background rectangle.                                  |
| `<Preview>`, `<SlidePreview>`                                          | HTML and SVG rendering of the slides, with `loading`, `error` and `caption` slots.           |
| `<PptxDownloadButton>`                                                 | A button that writes the file on click and saves it.                                         |
| `usePptx(document, props?, options?)`                                  | `blob`, `url`, `loading`, `error` and `warnings` refs that follow reactive changes.          |
| `useSlides(document, props?, options?)`                                | The resolved slide model for your own viewer or thumbnails.                                  |
| `renderToBuffer`, `renderToBlob`, `renderToModel`, `downloadPptx`      | Render once; `renderToBuffer` works in Node and workers.                                     |
| `pptx(document, props?, options?)`                                     | The low-level handle: `toBlob()`, `toBuffer()`, `toModel()`, `toPptxGenJS()`, `onChange()`.  |
| `pptxToVue`, `parsePptx`, `modelToVue` from `vue-pptx-document/import` | Read a .pptx, and write a model as a Vue component.                                          |

In templates, kebab-case attributes reach PptxGenJS in camelCase (`:font-size` is `fontSize`), flags such as `bold`, `bullet` and `show-legend` work as bare attributes, and colors take `#RRGGBB`, `#RGB` or CSS names.

`props` is typed from the presentation component, and every entry point takes `onWarn` and `strict` for warnings such as dropped text or an unknown prop, each with its slide and component.

The presentation is mounted as its own Vue app, so it doesn't see `provide`/`inject` or plugins from the page that shows it; pass what it needs as props.

## License

MIT. Files are written by [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (MIT) by Brent Ely and contributors; flex layout is [Yoga](https://github.com/facebook/yoga) (MIT) by Meta.
