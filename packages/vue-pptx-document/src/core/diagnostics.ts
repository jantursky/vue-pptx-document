/** What a warning is about. Match on these rather than on the wording, which may change. */
export type WarningCode =
  | 'dropped-text'
  | 'misplaced-object'
  | 'unknown-master'
  | 'master-object-dropped'
  | 'missing-src'
  | 'unknown-color'
  | 'unknown-prop'

export interface PptxWarning {
  code: WarningCode
  message: string
  /** 1-based number of the slide it happened on. */
  slide?: number
  /** The component it is about, such as `<Image>`. */
  object?: string
}

export interface DiagnosticOptions {
  /** Receives every warning. Without it, warnings go to `console.warn`, once per message and render. */
  onWarn?: (warning: PptxWarning) => void
  /** Rejects the render on the first warning. */
  strict?: boolean
}

interface Location {
  slide?: number
  object?: string
}

export interface Context extends Location {
  options: DiagnosticOptions
  seen: Set<string>
}

// Set only while synchronous work runs (see `inContext`), so concurrent renders never share it.
let active: Context | undefined

const PREFIX = '[vue-pptx-document]'

const where = ({ slide, object }: Location): string => {
  const parts = [slide === undefined ? undefined : `slide ${slide}`, object].filter(Boolean)
  return parts.length > 0 ? ` (${parts.join(', ')})` : ''
}

const located = new WeakSet<Error>()

const report = (warning: PptxWarning, context: Context | undefined): void => {
  const text = `${PREFIX} ${warning.message}${where(warning)}`
  if (context?.options.strict) {
    const error = new Error(`${PREFIX} ${warning.code}: ${warning.message}${where(warning)}`)
    located.add(error)
    throw error
  }
  if (context?.options.onWarn) {
    context.options.onWarn(warning)
    return
  }
  if (context) {
    if (context.seen.has(text)) return
    context.seen.add(text)
  }
  console.warn(text)
}

/**
 * Options that also keep every warning of one render in `warnings`, for composables that expose
 * them. Without a user `onWarn`, warnings still go to the console, once per message.
 */
export const collect = (options: DiagnosticOptions = {}): { options: DiagnosticOptions; warnings: PptxWarning[] } => {
  const warnings: PptxWarning[] = []
  const seen = new Set<string>()
  const onWarn = (warning: PptxWarning): void => {
    warnings.push(warning)
    if (options.onWarn) {
      options.onWarn(warning)
      return
    }
    const text = `${PREFIX} ${warning.message}${where(warning)}`
    if (seen.has(text)) return
    seen.add(text)
    console.warn(text)
  }
  return { options: { ...options, onWarn }, warnings }
}

/** Reports a warning on the render that is running, with its slide and object. */
export const warn = (code: WarningCode, message: string): void => {
  const context = active
  const warning: PptxWarning = { code, message }
  if (context?.slide !== undefined) warning.slide = context.slide
  if (context?.object !== undefined) warning.object = context.object
  report(warning, context)
}

/** A fresh context for one render: its own deduplication and options. */
export const createContext = (options: DiagnosticOptions = {}): Context => ({ options, seen: new Set() })

/** Runs synchronous work with `context` active. Never pass a function that awaits. */
export const inContext = <T>(context: Context, run: () => T): T => {
  const previous = active
  active = context
  try {
    return run()
  } finally {
    active = previous
  }
}

/** Adds where an error happened to its message, once, keeping the original as `cause`. */
export const locate = (error: unknown, location: Location): Error => {
  const original = error instanceof Error ? error : new Error(String(error))
  if (located.has(original) || (location.slide === undefined && location.object === undefined)) return original
  const message = original.message.startsWith(PREFIX) ? original.message.slice(PREFIX.length + 1) : original.message
  const wrapped = new Error(`${PREFIX} ${message}${where(location)}`, { cause: original })
  located.add(wrapped)
  return wrapped
}

/** Runs `run` with the slide or object recorded on the active context, and locates what it throws. */
export const at = <T>(location: Location, run: () => T): T => {
  const context = active
  const previous: Location = { slide: context?.slide, object: context?.object }
  const current: Location = {
    slide: location.slide ?? previous.slide,
    object: location.object ?? previous.object,
  }
  if (context) {
    context.slide = current.slide
    context.object = current.object
  }
  try {
    return run()
  } catch (error) {
    throw locate(error, current)
  } finally {
    if (context) {
      context.slide = previous.slide
      context.object = previous.object
    }
  }
}
