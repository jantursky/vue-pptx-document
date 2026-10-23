import { readFile } from 'node:fs/promises'

import { expect, test } from '@playwright/test'

import { download, watchErrors } from './helpers'

let errors: string[] = []

test.beforeEach(async ({ page }) => {
  errors = watchErrors(page)
  await page.goto('/#import')
})

test.afterEach(() => {
  expect(errors).toEqual([])
})

test('reads the sample deck back into a component', async ({ page }) => {
  await page.locator('[data-import-sample]').click()
  await expect(page.locator('[data-import-summary]')).toContainText('6 slides', { timeout: 15_000 })
  await expect(page.locator('.thumbs [data-pptx-slide]')).toHaveCount(6)
  await expect(page.locator('.thumbs [data-pptx-slide]').nth(1)).toContainText('Highlights')
  const code = page.locator('.generated')
  await expect(code).toContainText('<Presentation')
  await expect(code).toContainText(':slide-number=')
  const file = await download(page, 'SampleReport.vue', () =>
    code.getByRole('button', { name: 'Download .vue' }).click(),
  )
  expect(await readFile(await file.path(), 'utf8')).toContain('<Chart')
})

test('reads a file the user picks, and reports bad ones', async ({ page }) => {
  // A deck written by the report generator, downloaded first.
  await page.goto('/#generator')
  const deck = await download(page, 'q3-business-review.pptx', () =>
    page.getByRole('button', { name: 'Download PPTX' }).click(),
  )
  await page.goto('/#import')
  const buffer = await readFile(await deck.path())
  await page
    .locator('[data-field="pptx"]')
    .setInputFiles({ name: 'q3-business-review.pptx', mimeType: 'application/octet-stream', buffer })
  await expect(page.locator('[data-import-summary]')).toContainText('q3-business-review.pptx', { timeout: 15_000 })
  await page
    .locator('[data-field="pptx"]')
    .setInputFiles({ name: 'broken.pptx', mimeType: 'application/octet-stream', buffer: Buffer.from('not a zip') })
  await expect(page.getByRole('alert')).toBeVisible()
})
