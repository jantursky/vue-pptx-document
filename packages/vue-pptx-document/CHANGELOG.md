# Changelog

All notable changes to `vue-pptx-document` are listed here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.0] - 2026-10-23

### Added

- A Vue 3 custom renderer that turns `<Presentation>`, `<Slide>`, `<Master>`, `<Placeholder>`, `<Notes>`, `<Text>`, `<Paragraph>`, `<Run>`, `<Br>`, `<Shape>`, `<Image>`, `<Media>`, `<Chart>`, `<Table>`, `<TableRow>`, `<TableCell>` and `<View>` into a .pptx written by PptxGenJS 4.
- Absolute positions by default: numbers are inches, strings take `%`, `in`, `cm`, `mm`, `pt`, `px` or `emu`. Text and tables without a height get one measured from their content.
- Flexbox through Yoga in `<View>` and `<Slide layout="flex">`, mixed freely with absolute objects; `position: 'absolute'` children inside flex containers.
- Every PptxGenJS chart type, including combo charts with a secondary axis.
- `<Preview>` and `<SlidePreview>`: an HTML and SVG rendering of the resolved slides, with text, preset shapes, pictures, tables, charts, masters, placeholders, backgrounds, slide numbers and notes.
- `<PptxDownloadButton>`, `usePptx()`, `useSlides()`, `pptx()`, `renderToBuffer()`, `renderToBlob()`, `renderToModel()` and `downloadPptx()`, in the browser, in workers and in Node.
- `vue-pptx-document/import`: `pptxToVue()`, `parsePptx()` and `modelToVue()` read a .pptx (runs, bullets, shapes, groups, pictures, tables, charts, theme colors, backgrounds, layouts and masters, sections, notes, slide numbers) and write a Vue component that renders it.
- A `vue-pptx-document import deck.pptx -o Deck.vue` command line.

[Unreleased]: https://github.com/jantursky/vue-pptx-document/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/jantursky/vue-pptx-document/releases/tag/v1.0.0
