import type { AllowedComponentProps, Component, FunctionalComponent, MaybeRefOrGetter, VNodeProps } from 'vue'

type Builtin = keyof VNodeProps | keyof AllowedComponentProps

type Declared<C> = C extends abstract new (...args: never) => { $props: infer P }
  ? P
  : C extends FunctionalComponent<infer P>
    ? P
    : never

/**
 * The props a presentation component declares, for `renderToBuffer(Deck, props)` and the other
 * entry points. Components without typed props, and the wide `Component` type, take any object.
 */
export type DocumentProps<C> = Component extends C
  ? Record<string, unknown>
  : [Exclude<keyof Declared<C>, Builtin>] extends [never]
    ? Record<string, unknown>
    : Omit<Declared<C>, Builtin>

// A prop that accepts `undefined` counts as optional.
type RequiredKeys<T> = { [K in keyof T]-?: undefined extends T[K] ? never : K }[keyof T]

/** `props` is required when the component has required props, and optional otherwise. */
export type PropsArgs<C, Rest extends unknown[]> = [RequiredKeys<DocumentProps<C>>] extends [never]
  ? [props?: DocumentProps<C>, ...rest: Rest]
  : [props: DocumentProps<C>, ...rest: Rest]

/** The same for entry points that also take a ref or a getter. */
export type ReactivePropsArgs<C, Rest extends unknown[]> = [RequiredKeys<DocumentProps<C>>] extends [never]
  ? [props?: MaybeRefOrGetter<DocumentProps<C> | undefined>, ...rest: Rest]
  : [props: MaybeRefOrGetter<DocumentProps<C>>, ...rest: Rest]
