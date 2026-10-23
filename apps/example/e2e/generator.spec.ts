import { readFile } from 'node:fs/promises'

import { expect, test, type Page } from '@playwright/test'

import { download, readPptx, watchErrors } from './helpers'

// A 1×1 PNG, for the logo upload.
const PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==',
  'base64',
)

let errors: string[] = []

test.beforeEach(async ({ page }) => {
  errors = watchErrors(page)
  await page.goto('/#generator')
  await expect(slides(page)).toHaveCount(6, { timeout: 15_000 })
})

test.afterEach(() => {
  expect(errors).toEqual([])
})

const slides = (page: Page) => page.locator('.generator-preview [data-pptx-preview-slide]')

test('writes the six-slide report', async ({ page }) => {
  await expect(slides(page).first()).toContainText('Q3 business review')
  await expect(slides(page).nth(0).locator('[data-pptx-preview-notes]')).toBeVisible()
  const file = await download(page, 'q3-business-review.pptx', () =>
    page.getByRole('button', { name: 'Download PPTX' }).click(),
  )
  const deck = await readPptx(file)
  expect(deck).toHaveLength(6)
  expect(deck[0]).toContain('Q3 business review')
  expect(deck[1]).toContain('Churn dropped below 2%')
  expect(deck[4]).toEqual(expect.arrayContaining(['Region', 'Slovakia']))
})

test('follows edits and drops empty slides', async ({ page }) => {
  await page.locator('[data-field="title"]').fill('Annual review')
  await expect(slides(page).first()).toContainText('Annual review')
  await page.locator('[data-field="layout"]').selectOption('4x3')
  await expect(slides(page).first().locator('[data-pptx-slide]')).toHaveAttribute('style', /aspect-ratio: 10 \/ 7\.5/)
  await page.getByRole('button', { name: 'Clear' }).click()
  // Without highlights, metrics, series or regions, the title and closing slides remain.
  await expect(slides(page)).toHaveCount(2)
})

test('keeps the draft after a reload', async ({ page }) => {
  await page.locator('[data-field="title"]').fill('Kept')
  await page.reload()
  await expect(page.locator('[data-field="title"]')).toHaveValue('Kept')
})

test('exports and imports JSON', async ({ page }) => {
  await page.locator('[data-field="title"]').fill('Round trip')
  const file = await download(page, 'round-trip.json', () => page.getByRole('button', { name: 'Export JSON' }).click())
  const json = JSON.parse(await readFile(await file.path(), 'utf8'))
  expect(json.title).toBe('Round trip')
  await page.getByRole('button', { name: 'Load sample' }).click()
  await expect(page.locator('[data-field="title"]')).toHaveValue('Q3 business review')
  await page
    .locator('[data-field="import"]')
    .setInputFiles({ name: 'report.json', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(json)) })
  await expect(page.locator('[data-field="title"]')).toHaveValue('Round trip')
  await page
    .locator('[data-field="import"]')
    .setInputFiles({ name: 'bad.json', mimeType: 'application/json', buffer: Buffer.from('{"nope":1}') })
  await expect(page.getByRole('alert')).toContainText("isn't a report")
})

test('takes a logo and exports Vue code', async ({ page }) => {
  await page.locator('[data-field="logo"]').setInputFiles({ name: 'logo.png', mimeType: 'image/png', buffer: PNG })
  await expect(slides(page).first().locator('img')).toBeVisible()
  await page.locator('[data-export-vue]').click()
  const code = page.locator('.generator-preview .generated')
  await expect(code).toContainText('<Presentation')
  await expect(code).toContainText('data:image/png;base64')
  const file = await download(page, 'Q3BusinessReviewDeck.vue', () =>
    code.getByRole('button', { name: 'Download .vue' }).click(),
  )
  expect(await readFile(await file.path(), 'utf8')).toContain("from 'vue-pptx-document'")
})
