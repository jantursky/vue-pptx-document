import { readFile } from 'node:fs/promises'

import { expect, type Download, type Page } from '@playwright/test'
import JSZip from 'jszip'

/** Fails the test on uncaught errors and console errors. */
export const watchErrors = (page: Page): string[] => {
  const errors: string[] = []
  page.on('pageerror', (error) => errors.push(error.message))
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text())
  })
  return errors
}

/** Clicks `trigger` and waits for the download with this file name. */
export const download = async (page: Page, fileName: string, trigger: () => Promise<void>): Promise<Download> => {
  const [file] = await Promise.all([page.waitForEvent('download'), trigger()])
  expect(file.suggestedFilename()).toBe(fileName)
  return file
}

/** Opens a downloaded .pptx and returns the text of each slide, in order. */
export const readPptx = async (file: Download): Promise<string[][]> => {
  const zip = await JSZip.loadAsync(await readFile(await file.path()))
  const names = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => Number(/(\d+)\.xml/.exec(a)![1]) - Number(/(\d+)\.xml/.exec(b)![1]))
  return Promise.all(
    names.map(async (name) =>
      [...(await zip.file(name)!.async('string')).matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(([, text]) => text!),
    ),
  )
}
