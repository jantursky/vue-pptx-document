# Contributing

Thanks for helping. Bug reports, docs fixes and pull requests are all welcome.

## Before you start

- For anything bigger than a small fix, open an issue first so we can agree on the approach.
- If a file opens wrong in PowerPoint, try the same options with PptxGenJS directly (`pptx(Deck).toPptxGenJS()` gives you its presentation object). If it fails there too, the bug belongs [upstream](https://github.com/gitbrent/PptxGenJS/issues).
- If only the preview looks wrong, compare with the file in PowerPoint or Keynote: the file is the reference, the preview approximates it.

## Setup

You need [Bun](https://bun.sh) 1.3+ and Node.js 22.19+ (the version is in `.node-version`).

```sh
bun install
bunx playwright install chromium firefox webkit   # once, for the browser and e2e tests
bun run dev                                       # demo app, reading the library's source
```

## Where things live

| path                                                  |                                                                                                     |
| ----------------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| `packages/vue-pptx-document/src/core`                 | The Vue renderer, the snapshot, layout (absolute and Yoga flex), the model, the PptxGenJS writer.   |
| `packages/vue-pptx-document/src/components`           | Primitives, `Preview`, `SlidePreview`, `PptxDownloadButton`.                                        |
| `packages/vue-pptx-document/src/preview`              | The HTML and SVG renderer behind the preview: text, preset shapes, tables, charts.                  |
| `packages/vue-pptx-document/src/import`, `src/cli.ts` | The .pptx reader, the Vue code writer and the command line.                                         |
| `packages/vue-pptx-document/test`                     | Unit tests (Node + happy-dom) and `test/browser` (real engines).                                    |
| `apps/example`                                        | The demo, its Playwright e2e tests in `e2e/`, and `src/typecheck` with the README and docs samples. |
| `apps/docs`                                           | The Blume docs site. Live examples come from `apps/example/src/examples/*Example.vue`.              |

## Checks

```sh
bun run check          # samples in sync, format, lint, typecheck, unit tests with coverage
bun run test:browser   # writing and measuring in Chromium, Firefox and WebKit
bun run test:e2e       # builds the demo and drives it with Playwright
bun run build          # fails over the size budgets in .size-limit.json
bun run lint:package && bun run test:pack   # after build: publint, types, install the tarball, run the CLI
```

CI runs all of them on every pull request.

## Conventions

- TypeScript strict, `<script setup>` in components, Prettier formatting (`bun run format`).
- Every behavior change comes with a test. Unit tests write real .pptx files; `test/fixtures.ts` has helpers to unzip them and read the slide XML back.
- Code samples in `packages/vue-pptx-document/README.md` and `apps/docs/docs/getting-started.mdx` are type-checked. After editing them, run `bun run sync:samples`.
- Add a line under `## [Unreleased]` in `packages/vue-pptx-document/CHANGELOG.md` for anything users will notice.
- Breaking changes need a new major version and an entry under **Changed**.

## License

By contributing you agree that your contributions are licensed under the MIT license.
