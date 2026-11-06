import type { DeckModel } from './model'

/** Longer strings (image data, mostly) are hashed by length and ends, not in full. */
const LONG = 256
const END = 64

const digest = (_key: string, value: unknown): unknown =>
  typeof value === 'string' && value.length > LONG
    ? `\u0000${value.length}:${value.slice(0, END)}:${value.slice(-END)}`
    : value

/** A stable string for a resolved slide, master or layout. */
const fingerprint = (value: unknown): string => JSON.stringify(value, digest)

const sameKind = <T extends object>(previous: readonly T[], next: readonly T[]): T[] => {
  const pool = new Map<string, T[]>()
  for (const item of previous) {
    const key = fingerprint(item)
    const bucket = pool.get(key)
    if (bucket) bucket.push(item)
    else pool.set(key, [item])
  }
  const merged = next.map((item) => pool.get(fingerprint(item))?.shift() ?? item)
  // The same array too when nothing moved, so a prop holding it compares equal.
  return merged.length === previous.length && merged.every((item, index) => item === previous[index])
    ? (previous as T[])
    : merged
}

const sameValue = <T>(previous: T, next: T): T => (fingerprint(previous) === fingerprint(next) ? previous : next)

/**
 * Puts the previous slide, master and layout objects back wherever their content is
 * unchanged, so a consumer can tell what changed by comparing identities.
 */
export const reuseModel = (previous: DeckModel | null, next: DeckModel): DeckModel => {
  if (!previous) return next
  return {
    ...next,
    layout: sameValue(previous.layout, next.layout),
    theme: sameValue(previous.theme, next.theme),
    masters: sameKind(previous.masters, next.masters),
    slides: sameKind(previous.slides, next.slides),
  }
}
