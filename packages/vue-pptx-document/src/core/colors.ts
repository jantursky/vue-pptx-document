/** PowerPoint's theme color slots, which PptxGenJS passes through unchanged. */
export const THEME_COLORS = new Set([
  'tx1',
  'tx2',
  'bg1',
  'bg2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
])

const NAMED: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  orange: 'FFA500',
  purple: '800080',
  gray: '808080',
  grey: '808080',
  silver: 'C0C0C0',
  navy: '000080',
  teal: '008080',
  transparent: 'FFFFFF',
}

/**
 * Normalizes a CSS-style color (`#f35815`, `#fff`, `F35815`, `white`) to the six-digit
 * hex PptxGenJS writes. Theme colors such as `accent1` pass through.
 */
export const toHex = (value: unknown): string | undefined => {
  if (typeof value !== 'string') return undefined
  const color = value.trim()
  if (THEME_COLORS.has(color)) return color
  const named = NAMED[color.toLowerCase()]
  if (named) return named
  const hex = color.replace(/^#/, '')
  if (/^[0-9a-f]{3}$/i.test(hex)) {
    return hex
      .split('')
      .map((digit) => digit + digit)
      .join('')
      .toUpperCase()
  }
  if (/^[0-9a-f]{6}$/i.test(hex)) return hex.toUpperCase()
  // An 8-digit hex keeps only the color; use `transparency` for the alpha.
  if (/^[0-9a-f]{8}$/i.test(hex)) return hex.slice(0, 6).toUpperCase()
  return color
}

const isColorKey = (key: string): boolean => key === 'color' || key === 'highlight' || key.endsWith('Color')

/** Walks options and normalizes every color value, including `chartColors` arrays. */
export const normalizeColors = <T>(value: T): T => {
  if (Array.isArray(value)) return value.map((item: unknown) => normalizeColors(item)) as T
  if (value == null || typeof value !== 'object') return value
  const out: Record<string, unknown> = {}
  for (const [key, item] of Object.entries(value)) {
    if (isColorKey(key) && typeof item === 'string') out[key] = toHex(item)
    else if ((key === 'chartColors' || key === 'invertedColors') && Array.isArray(item)) {
      out[key] = item.map((color: unknown) => toHex(color) ?? color)
    } else out[key] = normalizeColors(item)
  }
  return out as T
}

/** CSS color for the preview: `#RRGGBB`, or a stand-in for theme colors. */
export const cssColor = (value: string | undefined, fallback = 'transparent'): string => {
  if (!value) return fallback
  const hex = toHex(value)
  if (!hex) return fallback
  if (THEME_COLORS.has(hex)) return THEME_PREVIEW[hex] ?? fallback
  return /^[0-9A-F]{6}$/.test(hex) ? `#${hex}` : hex
}

/** Office's default theme, used to preview theme colors. */
const THEME_PREVIEW: Record<string, string> = {
  tx1: '#000000',
  tx2: '#44546A',
  bg1: '#FFFFFF',
  bg2: '#E7E6E6',
  accent1: '#4472C4',
  accent2: '#ED7D31',
  accent3: '#A5A5A5',
  accent4: '#FFC000',
  accent5: '#5B9BD5',
  accent6: '#70AD47',
}

/** `transparency` is 0 to 100 in PptxGenJS. */
export const opacity = (transparency: unknown): number =>
  typeof transparency === 'number' ? Math.max(0, Math.min(1, 1 - transparency / 100)) : 1
