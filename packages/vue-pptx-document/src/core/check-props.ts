import { warn } from './diagnostics'
import { KNOWN_PROPS } from './known-props'
import { tag } from './nodes'

// Vue's own keys and event listeners aren't PptxGenJS options.
const IGNORED = /^(on[A-Z]|key$|ref$|ref_)/

let known: Map<string, Set<string>> | undefined

const loose = (key: string): string => key.toLowerCase().replace(/[-_]/g, '')

/** Warns about props a primitive doesn't know, with the closest known name. Development builds only. */
export const checkProps = (type: string, props: Record<string, unknown>): void => {
  known ??= new Map(Object.entries(KNOWN_PROPS).map(([name, list]) => [name, new Set(list.split(' '))]))
  const names = known.get(type)
  if (!names) return
  for (const key of Object.keys(props)) {
    if (names.has(key) || IGNORED.test(key)) continue
    const suggestion = [...names].find((name) => loose(name) === loose(key))
    const hint = suggestion ? ` Did you mean \`${suggestion}\`?` : ''
    warn('unknown-prop', `${tag(type)} has no \`${key}\` prop; PptxGenJS ignores it.${hint}`)
  }
}
