<p align="center"><img src="apps/docs/public/icon.svg" alt="" width="72" height="72" /></p>

# vue-pptx-document

[![npm](https://img.shields.io/npm/v/vue-pptx-document)](https://www.npmjs.com/package/vue-pptx-document) [![CI](https://github.com/jantursky/vue-pptx-document/actions/workflows/ci.yml/badge.svg)](https://github.com/jantursky/vue-pptx-document/actions/workflows/ci.yml) [![types](https://img.shields.io/npm/types/vue-pptx-document)](https://www.npmjs.com/package/vue-pptx-document) [![license](https://img.shields.io/npm/l/vue-pptx-document)](https://github.com/jantursky/vue-pptx-document/blob/main/LICENSE) [![codecov](https://codecov.io/gh/jantursky/vue-pptx-document/graph/badge.svg)](https://codecov.io/gh/jantursky/vue-pptx-document)

[Docs](https://vue-pptx-document-docs.vercel.app) · [Demo](https://vue-pptx-document.vercel.app) · [npm](https://www.npmjs.com/package/vue-pptx-document) · [Changelog](packages/vue-pptx-document/CHANGELOG.md)

Create PowerPoint files with Vue 3 components, and turn PowerPoint files back into Vue components. This repository holds the library, a demo app and the docs site.

| folder                       | what it is                                                                                        |
| ---------------------------- | ------------------------------------------------------------------------------------------------- |
| `packages/vue-pptx-document` | The library, published to npm. Its [README](packages/vue-pptx-document/README.md) covers the API. |
| `apps/example`               | The demo app: examples with their source, a report generator and a .pptx importer, on three tabs. |
| `apps/docs`                  | The docs site, built with [Blume](https://www.npmjs.com/package/blume).                           |

## The demo

- **Overview**: live previews next to their source. A minimal deck, absolute and flex layouts side by side, every chart type, and a report with a master, sections, notes, bullets and a table. Below them, the headless (`useSlides()` + `usePptx()`), on-demand (`downloadPptx()` behind a dynamic `import()`) and model-to-code patterns.
- **Report generator** ([`#generator`](https://vue-pptx-document.vercel.app/#generator)): a six-slide business review from one template, `apps/example/src/generator/ReportTemplate.vue`. Title, highlights, KPI cards in a flex row, a revenue chart, a table next to a doughnut and a closing slide, on a master with a footer and slide numbers. An accent color, a logo, three slide sizes, speaker notes, a draft saved in `localStorage`, JSON export and import, and **Export Vue code**, which writes the current deck out as a standalone component.
- **Import .pptx** ([`#import`](https://vue-pptx-document.vercel.app/#import)): drop a PowerPoint file and get thumbnails of its slides and the Vue component that renders them. The file is read in the browser and never uploaded.

## How it works

1. `createRenderer` from Vue builds and patches a plain tree of `{ type, props, children }` nodes from the components (`src/core/renderer.ts`). Templates, slots, `v-for` and reactivity work as in any Vue app.
2. A snapshot of the tree (`src/core/snapshot.ts`) is resolved into a model of slides and objects with frames in inches (`src/core/resolve.ts`). Absolute objects take their `x`, `y`, `w` and `h`; `<View>` and `<Slide layout="flex">` go through [Yoga](https://yogalayout.dev), which loads on first use (`src/core/flex.ts`). Text without a height is measured, with a canvas in the browser and glyph widths in Node.
3. The model feeds two renderers: PptxGenJS, which writes the file (`src/core/write.ts`), and the HTML and SVG preview (`src/preview`), so what you see while editing is the same layout that ends up in the file.
4. The importer (`src/import`) goes the other way: it unzips a .pptx with JSZip, reads the slide, layout, master, theme and chart XML into the same model, and writes that model out as a component.

### Prior art

Checked on 2026-10-22, before starting:

| project                                                              | approach                                     | state                                                                                                  |
| -------------------------------------------------------------------- | -------------------------------------------- | ------------------------------------------------------------------------------------------------------ |
| [`react-pptx`](https://github.com/wyozi/react-pptx)                  | React components rendered to PptxGenJS calls | 2.20.1 from August 2023, about 7,600 downloads a month. Absolute positions only, with an HTML preview. |
| [`vue-pptx`](https://www.npmjs.com/package/vue-pptx)                 | Vue wrapper, Chinese docs                    | 0.0.9, about 15 downloads a month; the name is why this package is `vue-pptx-document`.                |
| [`@vue-office/pptx`](https://www.npmjs.com/package/@vue-office/pptx) | Viewer for existing files, not a generator   | 1.0.1 from January 2025.                                                                               |
| [`pptxgenjs`](https://github.com/gitbrent/PptxGenJS)                 | The upstream writer, imperative API          | 4.0.1, about 14 million downloads a month.                                                             |

This library adds what those lack for Vue: a renderer with reactivity, flexbox next to absolute positions, a preview that shares the layout with the file, and an importer that turns existing decks into components.

## Requirements

- [Bun](https://bun.sh) 1.3 or newer, for installing and running scripts.
- Node.js 22.19 or newer. Vitest's coverage and the docs build run on it, and `bun run` hands those tools to Node.

## Commands

Run these from the repository root.

```sh
bun install            # install every workspace
bun run dev            # demo app with hot reload, reading the library's source
bun run dev:docs       # docs site with hot reload
bun run build          # library with size-limit budgets, then the demo app
bun run build:docs     # library, then the docs site
bun run check          # format check, lint, typecheck, tests with coverage
bun run test:browser   # writing and measuring in real Chromium, Firefox and WebKit
bun run test:e2e       # build the demo and drive it with Playwright in all three engines
bun run lint:package   # publint and type resolution (after build)
bun run test:pack      # install the packed tarball into an empty project, render, import and run the CLI (after build)
bun run sync:samples   # copy the README and docs code samples into the demo, where they're type-checked
bun run format         # format everything with Prettier
bun run lint:fix       # fix what ESLint can fix
```

`bun run build` ends with [size-limit](https://github.com/ai/size-limit) and fails if the library's own code grows past 26 kB minified and gzipped (about 25 kB today), if the library with PptxGenJS and Yoga passes 200 kB (about 191 kB), or if the importer passes 70 kB (about 62 kB). The budgets are in `packages/vue-pptx-document/.size-limit.json`, and `.github/workflows/size.yml` comments the change on every pull request. Coverage goes to [Codecov](https://codecov.io/gh/jantursky/vue-pptx-document) from CI. Coverage has minimum thresholds, so `bun run check` fails if they drop.

Tests come in seven layers:

- **Unit** (`packages/vue-pptx-document/test`, Vitest + happy-dom): the renderer, the snapshot, units and colors, text measuring, absolute and flex layout, the writer, the components, the composables, the preview of every shape and chart type, and the importer, code writer and CLI. They write real .pptx files in Node and read the slide XML back, or import them again.
- **Types** (`test/types.test.ts`): `expectTypeOf` and `@ts-expect-error` assertions that `bun run typecheck` enforces.
- **Browser** (`test/browser`): writing a Blob, measuring text with the canvas and fetching `usePptx()`'s object URL in Chromium, Firefox and WebKit.
- **End-to-end** (`apps/example/e2e`, Playwright): the built demo, with live previews, layout switching, chart types, downloads that open as zip files, the report generator's draft, JSON import and Vue export, and the importer's round trip.
- **Accessibility** (`apps/example/e2e/a11y.spec.ts`, axe-core): the demo's three tabs against WCAG 2.2 A and AA, in the light and dark themes.
- **Vue versions** (CI): the library's typecheck and unit tests against Vue 3.5.0, the oldest version its peer range allows, and the newest one on npm.
- **Package** (`scripts/pack-test.js`): the tarball installed into an empty project, imported and required from plain Node, with the importer and the `vue-pptx-document` command.

The code samples in the package README and the getting-started page are copied into `apps/example/src/typecheck` by `bun run sync:samples` and compiled with the demo; `bun run check` fails if a copy is stale.

CI (`.github/workflows/ci.yml`) runs all of it, plus `build`, `lint:package` and `build:docs`, on pushes to `main` and on pull requests.

The browser tests need Playwright's engines once per machine: `bunx playwright install chromium firefox webkit`.

## Releasing a new version

1. Move the entries under `## [Unreleased]` in `packages/vue-pptx-document/CHANGELOG.md` into a new section, such as `## [1.1.0] - 2026-11-03`, and add its link at the bottom. Leave an empty `## [Unreleased]` section above it.
2. Set the same version in `packages/vue-pptx-document/package.json`. The demo app and the docs read it from there.
3. Run `bun run check && bun run build:docs`, commit, and push.
4. Tag the commit and push the tag:

   ```sh
   git tag v1.1.0
   git push origin v1.1.0
   ```

   `.github/workflows/release.yml` checks that the tag matches the version, runs every test suite, publishes to npm with [provenance](https://docs.npmjs.com/generating-provenance-statements), and creates a GitHub release from the changelog section.

The workflow publishes through npm's [trusted publishing](https://docs.npmjs.com/trusted-publishers), so no token is stored in the repository. Set it up once on npmjs.com under the package's **Settings → Trusted publisher**: GitHub Actions, repository `jantursky/vue-pptx-document`, workflow `release.yml`, environment `npm`. The very first version has to be published by hand, since the package must exist before a trusted publisher can be added.

To publish by hand instead, run from the package folder. `prepublishOnly` runs the typecheck, both test suites, the build, the package checks and the tarball test first, and stops the publish if any fail:

```sh
cd packages/vue-pptx-document
bun publish --access public
```

Versions follow semantic versioning: a breaking change needs a new major version, listed under **Changed** in the changelog.

## Deploying to Vercel

The demo app and the docs deploy as two separate Vercel projects from the same repository. Both build the library from source first, so neither depends on a published npm version.

### Demo app

1. In Vercel, choose **Add New → Project** and import the repository.
2. Leave **Root Directory** at the repository root.
3. Leave the build settings alone. `vercel.json` at the root sets them:
   - install: `bun install --frozen-lockfile`
   - build: `bun run build`
   - output: `apps/example/dist`
4. Under **Settings → Build and Deployment**, set **Node.js Version** to 22.x.
5. Deploy.

### Docs site

1. Import the same repository again as a second project.
2. Set **Root Directory** to `apps/docs`.
3. Keep **Include source files outside of the Root Directory in the Build Step** turned on (it's the default). The build needs the library and the demo's example files.
4. Leave the build settings alone. `apps/docs/vercel.json` sets them:
   - install: `bun install --frozen-lockfile`, run from the repository root
   - build: `bun run build:docs`, run from the repository root
   - output: `dist`
5. Set **Node.js Version** to 22.x, as for the demo app.
6. Deploy. Blume picks up the production domain from Vercel's environment, so the sitemap and social cards get absolute URLs without extra config.

### From the command line

With the [Vercel CLI](https://vercel.com/docs/cli) installed and logged in:

```sh
vercel link              # once, at the repository root, for the demo project
vercel                   # preview deployment
vercel --prod            # production deployment

cd apps/docs
vercel link              # once, for the docs project
vercel --prod
```

After the first setup, every push to `main` deploys to production and every pull request gets a preview URL.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). Security issues go through [SECURITY.md](SECURITY.md), and everyone involved follows the [code of conduct](CODE_OF_CONDUCT.md).

## License

MIT. Files are written by [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (MIT); flex layout is [Yoga](https://github.com/facebook/yoga) (MIT).
