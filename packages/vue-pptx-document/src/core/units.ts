/** A length: a number is inches; strings take a unit (`in`, `cm`, `mm`, `pt`, `px`, `emu`) or `%`. */
export type Length = number | `${number}` | `${number}${'in' | 'cm' | 'mm' | 'pt' | 'px' | 'emu' | '%'}`

export const PX_PER_IN = 96
export const PT_PER_IN = 72
export const EMU_PER_IN = 914400

const PER_INCH: Record<string, number> = { in: 1, cm: 2.54, mm: 25.4, pt: PT_PER_IN, px: PX_PER_IN, emu: EMU_PER_IN }

const pattern = /^(-?\d*\.?\d+(?:e-?\d+)?)\s*(in|cm|mm|pt|px|emu|%)?$/i

/**
 * Converts a length to inches. Percentages are of `base`, the parent's width or height.
 * Returns undefined for anything that isn't a length, such as `auto`.
 */
export const toInches = (value: unknown, base = 0): number | undefined => {
  if (typeof value === 'number') return Number.isFinite(value) ? value : undefined
  if (typeof value !== 'string') return undefined
  const match = pattern.exec(value.trim())
  if (!match) return undefined
  const amount = Number(match[1])
  const unit = (match[2] ?? 'in').toLowerCase()
  if (unit === '%') return (amount / 100) * base
  return amount / (PER_INCH[unit] ?? 1)
}

export const isPercent = (value: unknown): value is `${number}%` =>
  typeof value === 'string' && /^-?\d*\.?\d+%$/.test(value.trim())

/** Rounds to 1/10000 inch, so frames stay readable in tests and generated code. */
export const round = (value: number): number => Math.round(value * 10000) / 10000
