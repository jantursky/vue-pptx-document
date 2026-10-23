// Copies the Vue code samples from the README and the docs into apps/example/src/typecheck,
// where `bun run typecheck` compiles them against the library. With `--check` it only
// reports copies that are out of date, which `bun run check` treats as a failure.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, relative } from 'node:path'

const root = join(import.meta.dirname, '..')
const out = join(root, 'apps/example/src/typecheck')
const check = process.argv.includes('--check')

const sources = [
  // The README marks each file with a leading `<!-- Name.vue -->` comment.
  {
    file: 'packages/vue-pptx-document/README.md',
    dir: 'readme',
    pattern: /```vue\n<!-- (\S+\.vue) -->\n([\s\S]*?)```/g,
  },
  // The docs put the file name after the language: ```vue App.vue
  { file: 'apps/docs/docs/getting-started.mdx', dir: 'getting-started', pattern: /```vue (\S+\.vue)\n([\s\S]*?)```/g },
]

const stale = []
for (const { file, dir, pattern } of sources) {
  const text = readFileSync(join(root, file), 'utf8')
  for (const [, name, code] of text.matchAll(pattern)) {
    const target = join(out, dir, name)
    const content = `<!--\n  Generated from ${file} by scripts/samples.js; edit the source, then run\n  \`bun run sync:samples\`. \`bun run typecheck\` compiles it against the library.\n-->\n${code}`
    if (existsSync(target) && readFileSync(target, 'utf8') === content) continue
    if (check) {
      stale.push(relative(root, target))
      continue
    }
    mkdirSync(dirname(target), { recursive: true })
    writeFileSync(target, content)
    console.log(`wrote ${relative(root, target)}`)
  }
}

if (stale.length > 0) {
  console.error(`Out of date, run \`bun run sync:samples\`:\n  ${stale.join('\n  ')}`)
  process.exit(1)
}
