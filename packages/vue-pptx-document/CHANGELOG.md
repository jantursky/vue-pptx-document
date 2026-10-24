# Changelog

All notable changes to `vue-pptx-document` are listed here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.1.2] - 2026-10-24

### Fixed

- The `vue-pptx-document` command's path in `package.json` no longer triggers npm's "script name was invalid and removed" warning on publish.

## [1.1.1] - 2026-10-24

### Changed

- The README on npm shows the demo screenshots and the new logo.

## [1.1.0] - 2026-10-23

### Added

- Diagnostics: every warning is `{ code, message, slide, object }` and goes to an `onWarn` option, which `pptx()`, the render functions, `usePptx`, `useSlides` and `<Preview @warn>` take. `strict: true` rejects the render on the first warning instead.
- `usePptx` and `useSlides` return the latest render's `warnings`.
- Errors name the slide and component they come from, including images PptxGenJS can't read, such as `ENOENT … (slide 3, <Image>)`.
- In development builds, a prop a component doesn't know warns with code `unknown-prop` and suggests the closest name (`fontsize` → `fontSize`). Bundlers drop the check from production builds.
- `props` is typed from the presentation component in `pptx()`, the render functions, the composables and `<Preview :document-props>`, and is required when the deck has required props.
- `<Chart>` props are typed from PptxGenJS's chart options.

### Changed

- Warnings are logged once per message for each render, instead of once per process, so a server that renders many decks sees them every time.
- Warning messages name components as written (`<Slide>`, `<TableRow>`) instead of their internal node types.
- `<Chart>` props no longer accept keys outside PptxGenJS's chart options. PptxGenJS ignored those keys anyway.
- The library's own size budget is 26 kB gzipped (was 24 kB); it measures 25 kB with this release.

### Fixed

- `transparent` and the alpha of an eight-digit hex color now become `transparency` instead of an opaque color. `transparent` used to write white.
- Every CSS color name is known (`lightgray`, `tomato`, `rebeccapurple`, ...). Unknown color values are dropped with a warning instead of being written into the file, where PowerPoint asked to repair it.
- An `<Image>` without a `src` is skipped with a warning instead of reaching PptxGenJS with an empty path.

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

[Unreleased]: https://github.com/jantursky/vue-pptx-document/compare/1.1.2...HEAD
[1.1.2]: https://github.com/jantursky/vue-pptx-document/compare/1.1.1...1.1.2
[1.1.1]: https://github.com/jantursky/vue-pptx-document/compare/1.1.0...1.1.1
[1.1.0]: https://github.com/jantursky/vue-pptx-document/compare/1.0.0...1.1.0
[1.0.0]: https://github.com/jantursky/vue-pptx-document/releases/tag/1.0.0
