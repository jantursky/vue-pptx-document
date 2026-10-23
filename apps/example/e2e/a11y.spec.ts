import AxeBuilder from '@axe-core/playwright'
import { expect, test, type Page } from '@playwright/test'

const TAGS = ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa']

// WCAG 2.2 A and AA. Inside the slide previews, colors are the deck author's choice, so
// contrast is checked everywhere except there; every other rule covers the slides too.
const audit = async (page: Page) => {
  const page_ = await new AxeBuilder({ page }).withTags(TAGS).exclude('[data-pptx-slide]').analyze()
  const slides = await new AxeBuilder({ page })
    .withTags(TAGS)
    .include('[data-pptx-slide]')
    .disableRules(['color-contrast'])
    .analyze()
  return [...page_.violations, ...slides.violations].map(
    ({ id, help, nodes }) => `${id}: ${help} (${nodes.map((node) => node.target.join(' ')).join(', ')})`,
  )
}

for (const colorScheme of ['light', 'dark'] as const) {
  test.describe(`${colorScheme} theme`, () => {
    test.use({ colorScheme })

    test('the overview has no accessibility violations', async ({ page }) => {
      await page.goto('/')
      await expect(page.locator('#example-report [data-pptx-slide]').first()).toBeVisible()
      await page.locator('#example-report').getByRole('button', { name: 'View source' }).click()
      expect(await audit(page)).toEqual([])
    })

    test('the report generator has no accessibility violations', async ({ page }) => {
      await page.goto('/#generator')
      await expect(page.locator('.generator-preview [data-pptx-preview-slide]')).toHaveCount(6, { timeout: 15_000 })
      expect(await audit(page)).toEqual([])
    })

    test('the importer has no accessibility violations', async ({ page }) => {
      await page.goto('/#import')
      await page.locator('[data-import-sample]').click()
      await expect(page.locator('[data-import-summary]')).toBeVisible({ timeout: 15_000 })
      expect(await audit(page)).toEqual([])
    })
  })
}
