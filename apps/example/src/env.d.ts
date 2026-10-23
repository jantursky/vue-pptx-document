// vue-tsc resolves .vue imports itself; this declaration is for tools that do not,
// such as the type-aware linter, so `createApp(App)` is not seen as untyped.
declare module '*.vue' {
  import type { DefineComponent } from 'vue'
  const component: DefineComponent
  export default component
}

/** The library version, from packages/vue-pptx-document/package.json (see vite.config.ts). */
declare const __VUE_PPTX_VERSION__: string
