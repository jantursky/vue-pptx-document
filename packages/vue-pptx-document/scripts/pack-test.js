// Installs the packed tarball into an empty project, the way a consumer gets it from
// npm, and writes a presentation from it in plain Node. Run after `bun run build`.
import { execFileSync } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

const packageDir = resolve(import.meta.dirname, '..')
const dir = mkdtempSync(join(tmpdir(), 'vue-pptx-document-pack-'))
const run = (command, args, cwd) => execFileSync(command, args, { cwd, encoding: 'utf8', stdio: 'pipe' })

try {
  const [packed] = JSON.parse(run('npm', ['pack', '--json', '--pack-destination', dir], packageDir))
  const files = packed.files.map((file) => file.path).sort()
  const allowed =
    /^(dist\/(index|import|cli)\.(js|js\.map|d\.ts)|dist\/chunks\/[\w.-]+\.(js|js\.map)|dist\/[\w/.-]+\.d\.ts|README\.md|CHANGELOG\.md|LICENSE|package\.json)$/
  const unexpected = files.filter((file) => !allowed.test(file))
  if (unexpected.length > 0) throw new Error(`Unexpected files in the tarball: ${unexpected.join(', ')}`)
  console.log(`tarball: ${packed.filename}, ${files.length} files, ${(packed.size / 1024).toFixed(1)} kB`)

  writeFileSync(join(dir, 'package.json'), JSON.stringify({ name: 'consumer', private: true, type: 'module' }))
  run('npm', ['install', '--no-audit', '--no-fund', '--loglevel=error', `./${packed.filename}`, 'vue@^3.5.0'], dir)

  writeFileSync(
    join(dir, 'consumer.js'),
    `import { writeFileSync } from 'node:fs'
import { h } from 'vue'
import { Presentation, Slide, Text, renderToBuffer } from 'vue-pptx-document'
import { pptxToVue } from 'vue-pptx-document/import'

const Hello = { render: () => h(Presentation, { title: 'Pack test' }, () => h(Slide, () => h(Text, { x: 1, y: 1, w: 8, h: 1 }, () => 'Hello'))) }
const bytes = await renderToBuffer(Hello)
if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error('Not a zip file')
writeFileSync('deck.pptx', bytes)
console.log('rendered ' + bytes.length + ' bytes from the installed package')
const { code } = await pptxToVue(bytes)
if (!code.includes('Hello')) throw new Error('The importer lost the text')
console.log('imported it back as ' + code.split('\\n').length + ' lines of Vue')
`,
  )
  console.log(run('node', ['consumer.js'], dir).trim())

  // The bin the package installs.
  run('npx', ['--no-install', 'vue-pptx-document', 'import', 'deck.pptx', '-o', 'Deck.vue'], dir)
  if (!readFileSync(join(dir, 'Deck.vue'), 'utf8').includes('<Presentation'))
    throw new Error('The CLI wrote no component')
  console.log('the CLI works')

  // CommonJS consumers load the ESM build through Node's require(esm).
  writeFileSync(
    join(dir, 'consumer.cjs'),
    `const { renderToBuffer, Presentation } = require('vue-pptx-document')
if (typeof renderToBuffer !== 'function' || !Presentation) throw new Error('require() returned no exports')
console.log('require() works')
`,
  )
  console.log(run('node', ['consumer.cjs'], dir).trim())
} finally {
  rmSync(dir, { recursive: true, force: true })
}
