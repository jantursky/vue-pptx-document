import { readFile, writeFile } from 'node:fs/promises'
import { basename } from 'node:path'
import { parseArgs } from 'node:util'

import { pptxToVue } from './import'

const HELP = `Usage: vue-pptx-document import <file.pptx> [options]

Writes a Vue component that renders the slides of a PowerPoint file.

Options:
  -o, --out <file>      Write the component here (default: print it)
      --no-images       Refer to pictures as media/<name> instead of embedding them
      --no-master       Skip the shapes slides inherit from their layout and master
      --from <module>   Import the components from this module (default: vue-pptx-document)
  -h, --help            Show this help
`

const main = async (): Promise<number> => {
  const { values, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      out: { type: 'string', short: 'o' },
      'no-images': { type: 'boolean', default: false },
      'no-master': { type: 'boolean', default: false },
      from: { type: 'string' },
      help: { type: 'boolean', short: 'h', default: false },
    },
  })
  const [command, file] = positionals
  if (values.help || command !== 'import' || !file) {
    process.stdout.write(HELP)
    return values.help ? 0 : 1
  }
  const { code, warnings, model } = await pptxToVue(await readFile(file), {
    embedImages: !values['no-images'],
    masterShapes: !values['no-master'],
    ...(values.from ? { importFrom: values.from } : {}),
  })
  for (const warning of warnings) process.stderr.write(`warning: ${warning}\n`)
  if (values.out) {
    await writeFile(values.out, code)
    process.stderr.write(`${basename(file)}: ${model.slides.length} slides written to ${values.out}\n`)
  } else process.stdout.write(code)
  return 0
}

main().then(
  (code) => {
    process.exitCode = code
  },
  (error: unknown) => {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`)
    process.exitCode = 1
  },
)
