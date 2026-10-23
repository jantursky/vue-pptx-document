import { THEME_COLORS } from '../core/colors'
import type { DeckModel, MasterModel, SlideModel, SlideObject, TextRun } from '../core/model'
import { LAYOUTS } from '../core/resolve'
import { round } from '../core/units'

type Options = Record<string, unknown>

export interface CodegenOptions {
  /** Where the components are imported from. */
  importFrom?: string
  /** Values longer than this many characters move to a constant in `<script setup>`. */
  inlineLimit?: number
}

const PARAGRAPH_KEYS = new Set([
  'align',
  'bullet',
  'indentLevel',
  'paraSpaceBefore',
  'paraSpaceAfter',
  'lineSpacing',
  'lineSpacingMultiple',
  'rtlMode',
])

/** Keys PptxGenJS adds itself, or that the model stores elsewhere. */
const INTERNAL = new Set(['path', 'data', 'breakLine', '_bodyProp', '_field'])

const kebab = (key: string): string => key.replace(/[A-Z]/g, (char) => `-${char.toLowerCase()}`)

const isColorKey = (key: string): boolean => key === 'color' || key === 'highlight' || key.endsWith('Color')

/** Writes hex colors back in CSS form: `F35815` becomes `#F35815`. */
const cssColors = (value: unknown, key = ''): unknown => {
  if (typeof value === 'string' && isColorKey(key) && /^[0-9A-F]{6}$/i.test(value) && !THEME_COLORS.has(value))
    return `#${value}`
  if (Array.isArray(value)) {
    return value.map((item) =>
      key === 'chartColors' && typeof item === 'string' ? cssColors(item, 'color') : cssColors(item),
    )
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([name, item]) => [name, cssColors(item, name)]))
  }
  return value
}

/** A JavaScript literal with single-quoted strings, to sit inside a double-quoted attribute. */
export const literal = (value: unknown, indent = ''): string => {
  if (value === undefined) return 'undefined'
  if (value === null) return 'null'
  if (typeof value === 'number') return String(round(value))
  if (typeof value === 'boolean') return String(value)
  if (typeof value === 'string')
    return `'${value.replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/\n/g, '\\n').replace(/"/g, '\\u0022')}'`
  if (Array.isArray(value)) {
    const items = value.map((item) => literal(item, `${indent}  `))
    const flat = `[${items.join(', ')}]`
    if (flat.length <= 80 && !flat.includes('\n')) return flat
    return `[\n${items.map((item) => `${indent}  ${item}`).join(',\n')},\n${indent}]`
  }
  const entries = Object.entries(value as Options).filter(([, item]) => item !== undefined)
  if (entries.length === 0) return '{}'
  const parts = entries.map(
    ([key, item]) => `${/^[A-Za-z_$][\w$]*$/.test(key) ? key : literal(key)}: ${literal(item, `${indent}  `)}`,
  )
  const flat = `{ ${parts.join(', ')} }`
  if (flat.length <= 80 && !flat.includes('\n')) return flat
  return `{\n${parts.map((part) => `${indent}  ${part}`).join(',\n')},\n${indent}}`
}

const escapeText = (text: string): string =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\{\{/g, '{&#123;')

const escapeAttribute = (text: string): string => text.replace(/&/g, '&amp;').replace(/"/g, '&quot;')

class Writer {
  readonly constants: string[] = []
  private readonly names = new Map<string, string>()
  private readonly counters = new Map<string, number>()
  readonly used = new Set<string>()

  constructor(private readonly limit: number) {}

  /** Moves a long value to a named constant, reusing one with the same content. */
  hoist(prefix: string, value: unknown): string {
    const code = literal(value)
    const existing = this.names.get(code)
    if (existing) return existing
    const count = (this.counters.get(prefix) ?? 0) + 1
    this.counters.set(prefix, count)
    const name = `${prefix}${count}`
    this.names.set(code, name)
    this.constants.push(`const ${name} = ${code}`)
    return name
  }

  /** Attributes for a component: strings as-is, everything else bound. */
  attributes(props: Options, hoistAs = 'value'): string {
    const out: string[] = []
    for (const [key, raw] of Object.entries(props)) {
      if (raw === undefined || INTERNAL.has(key)) continue
      const value = cssColors(raw, key)
      const name = kebab(key)
      if (value === true) out.push(name)
      else if (typeof value === 'string' && !value.includes('\n') && value.length <= this.limit)
        out.push(`${name}="${escapeAttribute(value)}"`)
      else {
        // `fill="#F35815"` for a plain color.
        if (
          (key === 'fill' || key === 'line') &&
          value &&
          typeof value === 'object' &&
          Object.keys(value).length === 1 &&
          'color' in value
        ) {
          out.push(`${name}="${(value as { color: string }).color}"`)
          continue
        }
        const code = literal(value, '')
        out.push(
          code.length > this.limit || code.includes('\n')
            ? `:${name}="${this.hoist(hoistAs, value)}"`
            : `:${name}="${code}"`,
        )
      }
    }
    return out.join(' ')
  }

  tag(name: string, attributes: string, children: string[] | string | null, indent: string): string {
    this.used.add(name)
    const open = attributes ? `<${name} ${attributes}` : `<${name}`
    if (children === null || children.length === 0) return `${indent}${open} />`
    if (typeof children === 'string') return `${indent}${open}>${children}</${name}>`
    return `${indent}${open}>\n${children.join('\n')}\n${indent}</${name}>`
  }
}

const frameProps = (object: SlideObject): Options => ({
  x: object.frame.x,
  y: object.frame.y,
  w: object.frame.w,
  h: object.frame.h,
})

const pickRun = (options: Options | undefined): Options =>
  Object.fromEntries(Object.entries(options ?? {}).filter(([key]) => !PARAGRAPH_KEYS.has(key) && key !== 'breakLine'))

const pickParagraph = (options: Options | undefined): Options =>
  Object.fromEntries(Object.entries(options ?? {}).filter(([key]) => PARAGRAPH_KEYS.has(key)))

const same = (a: unknown, b: unknown): boolean => JSON.stringify(a) === JSON.stringify(b)

/** Run options every run shares move up to the text box. */
const sharedOptions = (runs: TextRun[]): Options => {
  const first = runs[0]?.options as Options | undefined
  if (!first) return {}
  return Object.fromEntries(
    Object.entries(first).filter(
      ([key, value]) =>
        key !== 'breakLine' && runs.every((run) => same((run.options as Options | undefined)?.[key], value)),
    ),
  )
}

const writeRuns = (writer: Writer, runs: TextRun[], shared: Options, indent: string): string[] | string => {
  const groups: TextRun[][] = [[]]
  for (const run of runs) {
    groups[groups.length - 1]?.push(run)
    if (run.options?.breakLine) groups.push([])
  }
  if (groups.length > 1 && groups[groups.length - 1]?.length === 0) groups.pop()
  const strip = (options: Options): Options =>
    Object.fromEntries(Object.entries(options).filter(([key, value]) => !(key in shared && same(shared[key], value))))
  const inline = (group: TextRun[]): string =>
    group
      .map((run) => {
        const own = strip(pickRun(run.options as Options | undefined))
        const text = escapeText(run.text)
        if (Object.keys(own).length === 0) return text
        writer.used.add('Run')
        const attributes = writer.attributes(own)
        return `<Run ${attributes}>${text}</Run>`
      })
      .join('')
  if (groups.length === 1) {
    const [group] = groups as [TextRun[]]
    const paragraph = strip(pickParagraph(group[0]?.options as Options | undefined))
    if (Object.keys(paragraph).length === 0) return inline(group)
  }
  return groups.map((group) => {
    const paragraph = strip(pickParagraph(group[0]?.options as Options | undefined))
    return writer.tag('Paragraph', writer.attributes(paragraph), inline(group) || null, indent)
  })
}

const writeObject = (writer: Writer, object: SlideObject, indent: string): string => {
  const options = object.options as Options
  switch (object.kind) {
    case 'text': {
      const shared = sharedOptions(object.runs)
      const props = { ...frameProps(object), ...options, ...pickRun(shared), ...pickParagraph(shared) }
      const children = writeRuns(writer, object.runs, shared, `${indent}  `)
      return writer.tag('Text', writer.attributes(props), children, indent)
    }
    case 'shape':
      return writer.tag(
        'Shape',
        writer.attributes({ type: object.shape, ...frameProps(object), ...options }),
        null,
        indent,
      )
    case 'image': {
      const src = object.src.length > 120 ? writer.hoist('image', object.src) : undefined
      const props = { ...frameProps(object), ...options }
      const attributes = [src ? `:src="${src}"` : `src="${escapeAttribute(object.src)}"`, writer.attributes(props)]
        .filter(Boolean)
        .join(' ')
      return writer.tag('Image', attributes, null, indent)
    }
    case 'table': {
      const rows = writer.hoist('rows', cssColors(object.rows))
      return writer.tag(
        'Table',
        `:rows="${rows}" ${writer.attributes({ ...frameProps(object), ...options })}`,
        null,
        indent,
      )
    }
    case 'chart': {
      const type = Array.isArray(object.type)
        ? `:type="${writer.hoist('chart', cssColors(object.type))}"`
        : `type="${object.type}"`
      const data =
        Array.isArray(object.type) || object.data.length === 0 ? '' : ` :data="${writer.hoist('data', object.data)}"`
      return writer.tag(
        'Chart',
        `${type}${data} ${writer.attributes({ ...frameProps(object), ...options }, 'chartOptions')}`,
        null,
        indent,
      )
    }
    case 'media': {
      const { type, link, data, path, ...rest } = options as Options & {
        type?: string
        link?: string
        data?: string
        path?: string
      }
      const source = link ?? (data ? `data:${data}` : path)
      const src =
        source && source.length > 120
          ? `:src="${writer.hoist('media', source)}"`
          : source
            ? `src="${escapeAttribute(source)}"`
            : ''
      return writer.tag(
        'Media',
        [`type="${type ?? 'video'}"`, src, writer.attributes({ ...frameProps(object), ...rest })]
          .filter(Boolean)
          .join(' '),
        null,
        indent,
      )
    }
    case 'placeholder': {
      const { name, type, ...rest } = options as Options & { name?: string; type?: string }
      return writer.tag(
        'Placeholder',
        writer.attributes({ name: name ?? object.name, type, ...frameProps(object), ...rest }),
        object.text ? escapeText(object.text) : null,
        indent,
      )
    }
  }
}

const backgroundProp = (writer: Writer, background: SlideModel['background']): Options => {
  if (!background) return {}
  if (background.src)
    return { background: background.src.length > 120 ? undefined : background.src, __background: background.src }
  if (background.transparency === undefined && background.color) return { background: `#${background.color}` }
  return { background: cssColors(background) }
}

const containerAttributes = (writer: Writer, props: Options): string => {
  const { __background, ...rest } = props
  const parts: string[] = []
  if (typeof __background === 'string' && rest.background === undefined)
    parts.push(`:background="${writer.hoist('image', __background)}"`)
  const attributes = writer.attributes(rest)
  if (attributes) parts.push(attributes)
  return parts.join(' ')
}

const writeMaster = (writer: Writer, master: MasterModel, indent: string): string => {
  const props = {
    name: master.name,
    ...backgroundProp(writer, master.background),
    margin: master.margin,
    slideNumber: master.slideNumber,
  }
  return writer.tag(
    'Master',
    containerAttributes(writer, props),
    master.objects.map((object) => writeObject(writer, object, `${indent}  `)),
    indent,
  )
}

const writeSlide = (writer: Writer, slide: SlideModel, indent: string): string => {
  const props = {
    ...backgroundProp(writer, slide.background),
    color: slide.color,
    master: slide.master,
    section: slide.section,
    hidden: slide.hidden,
    slideNumber: slide.slideNumber,
  }
  const children = slide.objects.map((object) => writeObject(writer, object, `${indent}  `))
  if (slide.notes)
    children.push(writer.tag('Notes', '', escapeText(slide.notes).replace(/\n/g, '<Br />'), `${indent}  `))
  if (slide.notes?.includes('\n')) writer.used.add('Br')
  return writer.tag('Slide', containerAttributes(writer, props), children, indent)
}

/** Writes a Vue single-file component that renders the model with the library's components. */
export const modelToVue = (model: DeckModel, options: CodegenOptions = {}): string => {
  const writer = new Writer(options.inlineLimit ?? 100)
  const known = Object.entries(LAYOUTS).find(([, layout]) => layout.name === model.layout.name)
  const presentation: Options = {
    layout: known ? known[0] : { width: model.layout.width, height: model.layout.height },
    ...model.meta,
    rtl: model.rtl,
    headFontFace: model.theme?.headFontFace,
    bodyFontFace: model.theme?.bodyFontFace,
  }
  const children = [
    ...model.masters.map((master) => writeMaster(writer, master, '    ')),
    ...model.slides.map((slide) => writeSlide(writer, slide, '    ')),
  ]
  const body = writer.tag('Presentation', writer.attributes(presentation), children, '  ')
  const imports = [...writer.used].sort()
  const script = [
    `import { ${imports.join(', ')} } from '${options.importFrom ?? 'vue-pptx-document'}'`,
    ...(writer.constants.length > 0 ? ['', ...writer.constants] : []),
  ].join('\n')
  return `<script setup lang="ts">\n${script}\n</script>\n\n<template>\n${body}\n</template>\n`
}
