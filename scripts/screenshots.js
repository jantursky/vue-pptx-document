// Builds the demo, serves the build and captures the README screenshots into .github/assets:
// each view in light and dark at 1280×800 (device scale 2), plus one slide of the generated deck.
// Run it again whenever the demo's look changes: `bun run screenshots`.
// The callbacks passed to evaluate() run in the page, hence the browser globals.
/* global window, document */
import { execSync } from 'node:child_process'
import { mkdirSync, readFileSync, statSync } from 'node:fs'
import { createServer } from 'node:http'
import { extname, join, normalize, relative } from 'node:path'

import { chromium } from 'playwright'
import sharp from 'sharp'

const root = join(import.meta.dirname, '..')
const dist = join(root, 'apps/example/dist')
const out = join(root, '.github/assets')

const views = [
  { name: 'overview', hash: '' },
  {
    name: 'generator',
    hash: '#generator',
    // Start at the form, below the sticky header, so the live preview sits next to it.
    prepare: (page) => page.locator('.generator').evaluate((el) => window.scrollTo(0, el.offsetTop - 120)),
  },
  {
    name: 'import',
    hash: '#import',
    prepare: async (page) => {
      await page.locator('[data-import-sample]').click()
      await page.locator('[data-import-summary]').waitFor()
      await page.locator('.import-grid').evaluate((el) => window.scrollTo(0, el.offsetTop - 120))
    },
  },
]

const types = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.woff2': 'font/woff2',
}

const serve = () =>
  new Promise((resolve) => {
    const server = createServer((req, res) => {
      const path = normalize(decodeURIComponent(new URL(req.url, 'http://localhost').pathname))
      let file = join(dist, path)
      if (relative(dist, file).startsWith('..') || !statSync(file, { throwIfNoEntry: false })?.isFile()) {
        file = join(dist, 'index.html')
      }
      res.writeHead(200, { 'content-type': types[extname(file)] ?? 'application/octet-stream' })
      res.end(readFileSync(file))
    })
    server.listen(0, () => resolve(server))
  })

const save = async (buffer, name) => {
  const file = join(out, `${name}.png`)
  await sharp(buffer).png({ palette: true, quality: 90, effort: 10, compressionLevel: 9 }).toFile(file)
  console.log(`${relative(root, file)}  ${Math.round(statSync(file).size / 1024)} KB`)
}

// Fonts loaded, no requests in flight, and every preview past its debounce.
const settle = async (page) => {
  await page.waitForLoadState('networkidle')
  await page.evaluate(() => document.fonts.ready)
  await page.waitForTimeout(600)
}

execSync('bun run build', { cwd: root, stdio: 'inherit' })
mkdirSync(out, { recursive: true })
const server = await serve()
const base = `http://localhost:${server.address().port}/`
const browser = await chromium.launch()

try {
  for (const theme of ['light', 'dark']) {
    const context = await browser.newContext({
      viewport: { width: 1280, height: 800 },
      deviceScaleFactor: 2,
      colorScheme: theme,
      reducedMotion: 'reduce',
    })
    await context.addInitScript((value) => localStorage.setItem('vue-pptx-document:theme', value), theme)
    const page = await context.newPage()

    for (const view of views) {
      await page.goto(base + view.hash)
      await settle(page)
      await view.prepare?.(page)
      await settle(page)
      await save(await page.screenshot(), `${view.name}-${theme}`)
    }

    // The deck itself looks the same in both themes, so one crop is enough.
    if (theme === 'light') {
      await page.goto(`${base}#generator`)
      await settle(page)
      const slide = page.locator('.generator [data-pptx-slide]').nth(2)
      await slide.scrollIntoViewIfNeeded()
      await settle(page)
      await save(await slide.screenshot(), 'slide')
    }
    await context.close()
  }
} finally {
  await browser.close()
  server.close()
}
