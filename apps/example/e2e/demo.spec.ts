import { expect, test } from '@playwright/test'

import { download, readPptx, watchErrors } from './helpers'

let errors: string[] = []

test.beforeEach(async ({ page }) => {
  errors = watchErrors(page)
  await page.goto('/')
})

test.afterEach(() => {
  expect(errors).toEqual([])
})

test('previews every example', async ({ page }) => {
  await expect(page.getByRole('heading', { level: 1 })).toHaveText('PowerPoint files from Vue components')
  for (const id of ['#example-minimal', '#example-layout', '#example-charts', '#example-report']) {
    await expect(page.locator(`${id} [data-pptx-slide]`).first()).toBeVisible()
  }
  await expect(page.locator('#example-layout [data-pptx-slide]')).toHaveCount(2)
})

test('draws the slide again when the data changes, and downloads it', async ({ page }) => {
  const example = page.locator('#example-minimal')
  await example.getByLabel('Name').fill('Grace')
  await expect(example.locator('[data-pptx-slide]')).toContainText('Grace')
  const file = await download(page, 'hello.pptx', () => example.getByRole('button', { name: 'Download PPTX' }).click())
  const slides = await readPptx(file)
  expect(slides).toHaveLength(1)
  expect(slides[0]!.join(' ')).toContain('Grace')
})

test('lays out more cards in flex and absolute alike', async ({ page }) => {
  const example = page.locator('#example-layout')
  await example.getByRole('slider').fill('4')
  await example.getByLabel('column').check()
  for (const slide of await example.locator('[data-pptx-slide]').all())
    await expect(slide).toContainText('Download a .pptx')
})

test('switches chart types', async ({ page }) => {
  const example = page.locator('#example-charts')
  const chips = example.getByRole('group', { name: 'Chart type' }).getByRole('button')
  const count = await chips.count()
  expect(count).toBeGreaterThanOrEqual(10)
  for (let index = 0; index < count; index++) {
    await chips.nth(index).click()
    await expect(chips.nth(index)).toHaveAttribute('aria-pressed', 'true')
    await expect(example.locator('[data-pptx-slide] svg').first()).toBeVisible()
  }
})

test('shows the source of an example', async ({ page }) => {
  const example = page.locator('#example-report')
  await example.getByRole('button', { name: 'View source' }).click()
  await expect(example.locator('.source')).toContainText('<Presentation')
  await example.getByRole('button', { name: 'Hide source' }).click()
  await expect(example.locator('.source')).toHaveCount(0)
})

test('runs the headless, on-demand and codegen patterns', async ({ page }) => {
  await expect(page.locator('#demo-headless [data-state="loading"]')).toHaveText('false')
  const file = await download(page, 'charts-on-demand.pptx', () =>
    page.locator('#demo-on-demand').getByRole('button').click(),
  )
  expect((await readPptx(file)).length).toBeGreaterThan(0)
  await page.locator('#demo-codegen').getByRole('button', { name: 'Generate Vue code' }).click()
  await expect(page.locator('#demo-codegen')).toContainText('<Presentation')
})

test('remembers the theme', async ({ page }) => {
  const toggle = page.getByRole('button', { name: /^Theme:/ })
  await toggle.click()
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark')
  await page.reload()
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark')
})
