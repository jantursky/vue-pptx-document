import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vite'

const lib = (path: string) => fileURLToPath(new URL(`../../packages/vue-pptx-document/${path}`, import.meta.url))
const { version } = JSON.parse(readFileSync(lib('package.json'), 'utf8')) as { version: string }

// The dev server reads the library's source so edits hot-reload across packages.
// Builds resolve `vue-pptx-document` through node_modules, i.e. the library's real dist output.
export default defineConfig(({ command }) => ({
  plugins: [vue()],
  // PptxGenJS with JSZip is about 500 kB minified; it loads with the first download.
  build: { chunkSizeWarningLimit: 700 },
  define: {
    __VUE_PPTX_VERSION__: JSON.stringify(version),
  },
  resolve: {
    alias:
      command === 'serve'
        ? [
            { find: /^vue-pptx-document$/, replacement: lib('src/index.ts') },
            { find: /^vue-pptx-document\/import$/, replacement: lib('src/import/index.ts') },
          ]
        : [],
  },
}))
