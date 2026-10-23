import hljs from 'highlight.js/lib/core'
import css from 'highlight.js/lib/languages/css'
import javascript from 'highlight.js/lib/languages/javascript'
import typescript from 'highlight.js/lib/languages/typescript'
import xml from 'highlight.js/lib/languages/xml'

// Only the grammars the demo shows. A Vue file is highlighted as markup, which hands
// <script> and <style> blocks to the JavaScript and CSS grammars.
hljs.registerLanguage('xml', xml)
hljs.registerLanguage('javascript', javascript)
hljs.registerLanguage('typescript', typescript)
hljs.registerLanguage('css', css)

const languageOf = (fileName: string) => {
  if (fileName.endsWith('.ts')) return 'typescript'
  if (fileName.endsWith('.css')) return 'css'
  return 'xml'
}

export const highlight = (code: string, fileName: string): string =>
  hljs.highlight(code, { language: languageOf(fileName), ignoreIllegals: true }).value
