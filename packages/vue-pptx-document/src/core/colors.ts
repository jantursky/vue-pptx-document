import { warn } from './diagnostics'

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

// The CSS named colors, as name and hex pairs, unpacked on first use.
const NAMES =
  'aliceblue f0f8ff antiquewhite faebd7 aqua 00ffff aquamarine 7fffd4 azure f0ffff beige f5f5dc bisque ffe4c4 ' +
  'black 000000 blanchedalmond ffebcd blue 0000ff blueviolet 8a2be2 brown a52a2a burlywood deb887 cadetblue 5f9ea0 ' +
  'chartreuse 7fff00 chocolate d2691e coral ff7f50 cornflowerblue 6495ed cornsilk fff8dc crimson dc143c cyan 00ffff ' +
  'darkblue 00008b darkcyan 008b8b darkgoldenrod b8860b darkgray a9a9a9 darkgreen 006400 darkgrey a9a9a9 ' +
  'darkkhaki bdb76b darkmagenta 8b008b darkolivegreen 556b2f darkorange ff8c00 darkorchid 9932cc darkred 8b0000 ' +
  'darksalmon e9967a darkseagreen 8fbc8f darkslateblue 483d8b darkslategray 2f4f4f darkslategrey 2f4f4f ' +
  'darkturquoise 00ced1 darkviolet 9400d3 deeppink ff1493 deepskyblue 00bfff dimgray 696969 dimgrey 696969 ' +
  'dodgerblue 1e90ff firebrick b22222 floralwhite fffaf0 forestgreen 228b22 fuchsia ff00ff gainsboro dcdcdc ' +
  'ghostwhite f8f8ff gold ffd700 goldenrod daa520 gray 808080 green 008000 greenyellow adff2f grey 808080 ' +
  'honeydew f0fff0 hotpink ff69b4 indianred cd5c5c indigo 4b0082 ivory fffff0 khaki f0e68c lavender e6e6fa ' +
  'lavenderblush fff0f5 lawngreen 7cfc00 lemonchiffon fffacd lightblue add8e6 lightcoral f08080 lightcyan e0ffff ' +
  'lightgoldenrodyellow fafad2 lightgray d3d3d3 lightgreen 90ee90 lightgrey d3d3d3 lightpink ffb6c1 ' +
  'lightsalmon ffa07a lightseagreen 20b2aa lightskyblue 87cefa lightslategray 778899 lightslategrey 778899 ' +
  'lightsteelblue b0c4de lightyellow ffffe0 lime 00ff00 limegreen 32cd32 linen faf0e6 magenta ff00ff ' +
  'maroon 800000 mediumaquamarine 66cdaa mediumblue 0000cd mediumorchid ba55d3 mediumpurple 9370db ' +
  'mediumseagreen 3cb371 mediumslateblue 7b68ee mediumspringgreen 00fa9a mediumturquoise 48d1cc ' +
  'mediumvioletred c71585 midnightblue 191970 mintcream f5fffa mistyrose ffe4e1 moccasin ffe4b5 ' +
  'navajowhite ffdead navy 000080 oldlace fdf5e6 olive 808000 olivedrab 6b8e23 orange ffa500 orangered ff4500 ' +
  'orchid da70d6 palegoldenrod eee8aa palegreen 98fb98 paleturquoise afeeee palevioletred db7093 ' +
  'papayawhip ffefd5 peachpuff ffdab9 peru cd853f pink ffc0cb plum dda0dd powderblue b0e0e6 purple 800080 ' +
  'rebeccapurple 663399 red ff0000 rosybrown bc8f8f royalblue 4169e1 saddlebrown 8b4513 salmon fa8072 ' +
  'sandybrown f4a460 seagreen 2e8b57 seashell fff5ee sienna a0522d silver c0c0c0 skyblue 87ceeb ' +
  'slateblue 6a5acd slategray 708090 slategrey 708090 snow fffafa springgreen 00ff7f steelblue 4682b4 tan d2b48c ' +
  'teal 008080 thistle d8bfd8 tomato ff6347 turquoise 40e0d0 violet ee82ee wheat f5deb3 white ffffff ' +
  'whitesmoke f5f5f5 yellow ffff00 yellowgreen 9acd32'

let named: Map<string, string> | undefined
const namedColor = (name: string): string | undefined => {
  if (!named) {
    named = new Map()
    for (const [, key, hex] of NAMES.matchAll(/(\w+) (\w{6})/g)) named.set(key ?? '', (hex ?? '').toUpperCase())
  }
  return named.get(name)
}

interface ParsedColor {
  /** Six-digit hex or a theme color. */
  hex: string
  /** 0 to 1. */
  alpha: number
}

/** Reads a CSS-style color. Unknown values are undefined. */
const parseColor = (value: string): ParsedColor | undefined => {
  const color = value.trim()
  if (THEME_COLORS.has(color)) return { hex: color, alpha: 1 }
  const lower = color.toLowerCase()
  if (lower === 'transparent') return { hex: 'FFFFFF', alpha: 0 }
  const name = namedColor(lower)
  if (name) return { hex: name, alpha: 1 }
  const hex = color.replace(/^#/, '')
  if (/^[0-9a-f]{3}$/i.test(hex)) {
    return {
      hex: hex
        .split('')
        .map((digit) => digit + digit)
        .join('')
        .toUpperCase(),
      alpha: 1,
    }
  }
  if (/^[0-9a-f]{6}$/i.test(hex)) return { hex: hex.toUpperCase(), alpha: 1 }
  if (/^[0-9a-f]{8}$/i.test(hex)) return { hex: hex.slice(0, 6).toUpperCase(), alpha: parseInt(hex.slice(6), 16) / 255 }
  return undefined
}

/**
 * Normalizes a CSS-style color (`#f35815`, `#fff`, `F35815`, `white`, `transparent`) to the
 * six-digit hex PptxGenJS writes. Theme colors such as `accent1` pass through. The alpha of an
 * eight-digit hex or `transparent` is dropped here; `normalizeColors` turns it into `transparency`.
 */
export const toHex = (value: unknown): string | undefined =>
  typeof value === 'string' ? parseColor(value)?.hex : undefined

const isColorKey = (key: string): boolean => key === 'color' || key === 'highlight' || key.endsWith('Color')

/** Walks options and normalizes every color value, including `chartColors` arrays. */
export const normalizeColors = <T>(value: T): T => {
  if (Array.isArray(value)) return value.map((item: unknown) => normalizeColors(item)) as T
  if (value == null || typeof value !== 'object') return value
  const out: Record<string, unknown> = {}
  for (const [key, item] of Object.entries(value)) {
    if (isColorKey(key) && typeof item === 'string') {
      const parsed = parseColor(item)
      if (!parsed) {
        warn('unknown-color', `Unknown color "${item}" is ignored. Use a CSS color name or hex.`)
        continue
      }
      out[key] = parsed.hex
      // `color` sits next to the object's `transparency` (0 to 100); an explicit one wins.
      if (key === 'color' && parsed.alpha < 1 && (value as Record<string, unknown>).transparency === undefined) {
        out.transparency = Math.round((1 - parsed.alpha) * 100)
      }
    } else if ((key === 'chartColors' || key === 'invertedColors') && Array.isArray(item)) {
      out[key] = item.map((color: unknown) => toHex(color) ?? color)
    } else out[key] = normalizeColors(item)
  }
  return out as T
}

/** CSS color for the preview: `#RRGGBB`, `rgba()` when it has an alpha, or a stand-in for theme colors. */
export const cssColor = (value: string | undefined, fallback = 'transparent'): string => {
  const parsed = value ? parseColor(value) : undefined
  if (!parsed) return fallback
  const { hex, alpha } = parsed
  if (THEME_COLORS.has(hex)) return THEME_PREVIEW[hex] ?? fallback
  if (alpha === 0) return 'transparent'
  if (alpha < 1) {
    const [r, g, b] = [0, 2, 4].map((offset) => parseInt(hex.slice(offset, offset + 2), 16))
    return `rgba(${r}, ${g}, ${b}, ${Math.round(alpha * 100) / 100})`
  }
  return `#${hex}`
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
