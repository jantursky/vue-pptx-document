import vue from '@vitejs/plugin-vue'
import { playwright } from '@vitest/browser-playwright'
import { defineConfig } from 'vitest/config'

// The unit suite runs in happy-dom under Node, where PptxGenJS loads its Node build.
// These tests run the browser path (Blob output, canvas text measurement) in real engines.
export default defineConfig({
  plugins: [vue()],
  test: {
    include: ['test/browser/**/*.test.ts'],
    browser: {
      enabled: true,
      headless: true,
      provider: playwright(),
      instances: [{ browser: 'chromium' }, { browser: 'firefox' }, { browser: 'webkit' }],
    },
  },
})
