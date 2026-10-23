import { readFileSync } from 'node:fs'

import { defineConfig } from 'blume'

const pkg = JSON.parse(
  readFileSync(new URL('../../packages/vue-pptx-document/package.json', import.meta.url), 'utf8'),
) as {
  version: string
}

export default defineConfig({
  title: 'vue-pptx-document',
  description: 'Create PowerPoint files with Vue 3 components.',
  logo: '/logo.svg',
  github: {
    owner: 'jantursky',
    repo: 'vue-pptx-document',
    dir: 'apps/docs',
  },
  navigation: {
    actions: [
      { href: 'https://vue-pptx-document.vercel.app', label: 'Demo' },
      { href: 'https://www.npmjs.com/package/vue-pptx-document', label: 'npm' },
    ],
  },
  content: {
    root: 'docs',
  },
  theme: {
    accent: 'orange',
    mode: 'system',
  },
  variables: {
    version: pkg.version,
  },
  // The live previews run the demo app's example files, so the docs and the demo never drift apart.
  examples: {
    source: '../example/src/examples/*Example.vue',
    css: 'examples.css',
  },
})
