import type { DeckModel } from '../core/model'
import { modelToVue, type CodegenOptions } from './codegen'
import { parsePptx, type ParseOptions, type PptxInput } from './parse'

export { modelToVue, literal, type CodegenOptions } from './codegen'
export { parsePptx, type ParseOptions, type ParseResult, type PptxInput } from './parse'

export interface PptxToVueResult {
  /** The Vue single-file component. */
  code: string
  model: DeckModel
  warnings: string[]
}

/** Reads a .pptx file and writes a Vue component that renders the same slides. */
export const pptxToVue = async (
  input: PptxInput,
  options: ParseOptions & CodegenOptions = {},
): Promise<PptxToVueResult> => {
  const { model, warnings } = await parsePptx(input, options)
  return { code: modelToVue(model, options), model, warnings }
}
