import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vite'
import dts from 'vite-plugin-dts'

const root = (path: string): string => resolve(import.meta.dirname, path)

const pkg = JSON.parse(readFileSync(root('package.json'), 'utf8')) as { dependencies: Record<string, string> }
const external = ['vue', ...Object.keys(pkg.dependencies)]
const isExternal = (id: string): boolean =>
  id.startsWith('node:') || external.some((name) => id === name || id.startsWith(`${name}/`))

export default defineConfig({
  plugins: [
    vue(),
    dts({
      tsconfigPath: './tsconfig.build.json',
      include: ['src'],
      bundleTypes: true,
    }),
  ],
  build: {
    lib: {
      // The importer and the CLI are entries of their own, so apps that only render never load them.
      entry: { index: root('src/index.ts'), import: root('src/import/index.ts'), cli: root('src/cli.ts') },
      // ESM only, like Vue's own runtime build and yoga-layout.
      formats: ['es'],
      fileName: (_format, name) => `${name}.js`,
    },
    rollupOptions: {
      external: isExternal,
      output: {
        banner: (chunk) => (chunk.name === 'cli' ? '#!/usr/bin/env node' : ''),
        chunkFileNames: 'chunks/[name]-[hash].js',
      },
    },
    sourcemap: true,
    minify: true,
    target: 'es2020',
    emptyOutDir: true,
  },
})
