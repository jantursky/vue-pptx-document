<script setup lang="ts">
import { defineAsyncComponent, onBeforeUnmount, ref, watch } from 'vue'

import CodegenDemo from './demos/CodegenDemo.vue'
import HeadlessDemo from './demos/HeadlessDemo.vue'
import OnDemandDemo from './demos/OnDemandDemo.vue'
import ExamplesSection from './examples/ExamplesSection.vue'
import helloSource from './examples/HelloDeck.vue?raw'
import CodeBlock from './ui/CodeBlock.vue'
import CopyButton from './ui/CopyButton.vue'

// The generator and the importer load with their tabs.
const ReportGenerator = defineAsyncComponent(() => import('./generator/ReportGenerator.vue'))
const ImportPanel = defineAsyncComponent(() => import('./importer/ImportPanel.vue'))

const version = __VUE_PPTX_VERSION__
const install = 'npm install vue-pptx-document'

type Tab = 'overview' | 'generator' | 'import'
const tabs: { id: Tab; label: string }[] = [
  { id: 'overview', label: 'Overview' },
  { id: 'generator', label: 'Report generator' },
  { id: 'import', label: 'Import .pptx' },
]

// The tab lives in the URL hash, so #generator links straight to it and Back works.
const fromHash = (): Tab =>
  location.hash === '#generator' ? 'generator' : location.hash === '#import' ? 'import' : 'overview'
const tab = ref<Tab>(fromHash())
const onHashChange = () => {
  tab.value = fromHash()
}
window.addEventListener('hashchange', onHashChange)
onBeforeUnmount(() => window.removeEventListener('hashchange', onHashChange))

const select = (next: Tab) => {
  tab.value = next
  history.pushState(null, '', next === 'overview' ? location.pathname : `#${next}`)
}

const tabButtons = ref<HTMLButtonElement[]>([])
const onTabKey = (event: KeyboardEvent, index: number) => {
  const step = event.key === 'ArrowRight' ? 1 : event.key === 'ArrowLeft' ? -1 : 0
  if (!step) return
  event.preventDefault()
  const next = tabs[(index + step + tabs.length) % tabs.length]
  if (!next) return
  select(next.id)
  tabButtons.value[tabs.indexOf(next)]?.focus()
}

const open = (next: Tab) => {
  select(next)
  document.getElementById('tabs')?.scrollIntoView({ behavior: 'smooth' })
}

// Light, dark, or the system's choice, kept on <html data-theme>.
type Theme = 'system' | 'light' | 'dark'
const readTheme = (): Theme => {
  try {
    const stored = localStorage.getItem('vue-pptx-document:theme')
    return stored === 'light' || stored === 'dark' ? stored : 'system'
  } catch {
    return 'system'
  }
}
const theme = ref<Theme>(readTheme())
watch(
  theme,
  (value) => {
    if (value === 'system') document.documentElement.removeAttribute('data-theme')
    else document.documentElement.setAttribute('data-theme', value)
    try {
      if (value === 'system') localStorage.removeItem('vue-pptx-document:theme')
      else localStorage.setItem('vue-pptx-document:theme', value)
    } catch {
      // The choice just isn't remembered.
    }
  },
  { immediate: true },
)
const nextTheme = () => {
  theme.value = theme.value === 'system' ? 'dark' : theme.value === 'dark' ? 'light' : 'system'
}
</script>

<template>
  <a class="skip" href="#main">Skip to content</a>
  <div class="banner">
    <div class="container">
      <span class="pill">New</span>
      Import a .pptx and get a Vue component back.
      <button type="button" class="banner-link" @click="open('import')">Try it</button>
    </div>
  </div>
  <header class="nav">
    <div class="container nav-inner">
      <a class="brand" href="/">
        <svg class="brand-mark" viewBox="0 0 32 32" aria-hidden="true">
          <rect x="4" y="7" width="24" height="18" fill="none" stroke="currentColor" stroke-width="2" />
          <rect x="8" y="11" width="12" height="3" style="fill: var(--ember)" />
          <rect x="8" y="17" width="8" height="2" fill="currentColor" />
        </svg>
        vue-pptx-document
        <span class="version">{{ version }}</span>
      </a>
      <nav class="nav-links" aria-label="Project">
        <a href="https://vue-pptx-document-docs.vercel.app">Docs</a>
        <a href="https://github.com/jantursky/vue-pptx-document">GitHub</a>
        <a href="https://www.npmjs.com/package/vue-pptx-document">npm</a>
        <button type="button" class="theme-toggle" title="Change theme" @click="nextTheme">
          <span class="sr-only">Theme: </span>{{ theme === 'system' ? 'auto' : theme }}
        </button>
      </nav>
    </div>
  </header>

  <main id="main">
    <section class="band hero">
      <div class="container hero-grid">
        <div class="hero-copy">
          <p class="tag">Vue 3 · PowerPoint</p>
          <h1 class="headline">PowerPoint files from Vue components</h1>
          <p class="lead">
            Write slides with <code>&lt;Presentation&gt;</code>, <code>&lt;Slide&gt;</code>, <code>&lt;Text&gt;</code>,
            <code>&lt;Shape&gt;</code>, <code>&lt;Image&gt;</code>, <code>&lt;Chart&gt;</code> and
            <code>&lt;Table&gt;</code>. Place them in inches or lay them out with flexbox. A custom Vue renderer builds
            the tree, <code>&lt;Preview&gt;</code> draws it live, and PptxGenJS writes the .pptx, in the browser or in
            Node.
          </p>
          <div class="actions">
            <a class="button button-primary" href="https://vue-pptx-document-docs.vercel.app/getting-started"
              >Get started</a
            >
            <button type="button" class="button" @click="open('generator')">Try the report generator</button>
          </div>
          <div class="install">
            <code><span class="install-prompt" aria-hidden="true">$</span> {{ install }}</code>
            <CopyButton :text="install" label="Copy the install command" />
          </div>
        </div>
        <figure class="hero-code">
          <figcaption>HelloDeck.vue</figcaption>
          <CodeBlock :code="helloSource" file-name="HelloDeck.vue" />
        </figure>
      </div>
    </section>

    <div id="tabs" class="tabs-bar">
      <div class="container">
        <div class="tabs" role="tablist" aria-label="Demo">
          <button
            v-for="(item, index) in tabs"
            :id="`tab-${item.id}`"
            :key="item.id"
            ref="tabButtons"
            type="button"
            role="tab"
            :aria-selected="tab === item.id"
            :aria-controls="`panel-${item.id}`"
            :tabindex="tab === item.id ? 0 : -1"
            @click="select(item.id)"
            @keydown="onTabKey($event, index)"
          >
            {{ item.label }}
          </button>
        </div>
      </div>
    </div>

    <div v-if="tab === 'generator'" id="panel-generator" role="tabpanel" aria-labelledby="tab-generator">
      <section class="band">
        <div class="container container-wide">
          <header class="section-head">
            <p class="tag">Full demo</p>
            <h2 class="headline">Report generator</h2>
            <p>
              A six-slide business review from one template: a title slide, highlights, KPI cards laid out with flexbox,
              a revenue chart, a table next to a doughnut, and a closing slide, on a master with a footer and slide
              numbers. Your accent color and logo, three slide sizes. The draft is saved in this browser; export it as
              JSON to keep it, or as a Vue component to start your own template.
            </p>
          </header>
          <ReportGenerator />
        </div>
      </section>
    </div>

    <div v-else-if="tab === 'import'" id="panel-import" role="tabpanel" aria-labelledby="tab-import">
      <section class="band">
        <div class="container container-wide">
          <header class="section-head">
            <p class="tag">Codegen</p>
            <h2 class="headline">From .pptx to a Vue component</h2>
            <p>
              <code>pptxToVue()</code> from <code>vue-pptx-document/import</code> reads text with its runs and bullets,
              shapes, pictures, tables, charts, backgrounds, sections and notes, and writes a component you can edit.
              The same importer runs from the command line:
              <code>npx vue-pptx-document import deck.pptx -o Deck.vue</code>.
            </p>
          </header>
          <ImportPanel />
        </div>
      </section>
    </div>

    <div v-else id="panel-overview" role="tabpanel" aria-labelledby="tab-overview">
      <section class="band">
        <div class="container">
          <ExamplesSection />
        </div>
      </section>

      <section class="band band-alt" aria-labelledby="demos-title">
        <div class="container">
          <header class="section-head">
            <p class="tag">Patterns</p>
            <h2 id="demos-title" class="headline">More ways to use it</h2>
          </header>
          <div class="demos">
            <section id="demo-headless" class="card demo">
              <header>
                <h3>Headless</h3>
                <p>
                  <code>useSlides()</code> resolves the slides for your own viewer, here one
                  <code>&lt;SlidePreview&gt;</code>. <code>usePptx()</code> keeps a <code>blob</code> and
                  <code>url</code> of the file current.
                </p>
              </header>
              <HeadlessDemo />
            </section>
            <section id="demo-on-demand" class="card demo">
              <header>
                <h3>On demand</h3>
                <p>
                  <code>downloadPptx()</code> behind a button, with the library and the deck behind a dynamic
                  <code>import()</code>, so an app loads them only when someone asks for the file.
                </p>
              </header>
              <OnDemandDemo />
            </section>
            <section id="demo-codegen" class="card demo demo-wide">
              <header>
                <h3>Model to code</h3>
                <p>
                  <code>renderToModel()</code> gives every object's resolved frame. <code>modelToVue()</code> writes it
                  back as a component with absolute positions, which is how the generator's <em>Export Vue code</em>
                  works.
                </p>
              </header>
              <CodegenDemo />
            </section>
          </div>
        </div>
      </section>

      <section class="band" aria-labelledby="engine-title">
        <div class="container split-60">
          <header class="section-head">
            <p class="tag">Under the hood</p>
            <h2 id="engine-title" class="headline">A Vue renderer for PptxGenJS</h2>
            <p>
              <code>createRenderer</code> from Vue builds a plain tree of slides and objects from your components, so
              templates, slots, <code>v-for</code> and reactivity work as usual. Flex containers go through Yoga, the
              layout engine of React Native and react-pdf. The resolved tree feeds both the HTML preview and PptxGenJS,
              which only loads when you write a file.
            </p>
          </header>
          <ul class="stats">
            <li class="card stat">
              <span class="tag">Preview</span>
              <strong>No file needed</strong>
              <span>HTML and SVG from the same tree, as you type</span>
            </li>
            <li class="card stat">
              <span class="tag">Layout</span>
              <strong>Inches or flexbox</strong>
              <span>absolute frames, Yoga, or both on one slide</span>
            </li>
            <li class="card stat">
              <span class="tag">One entry</span>
              <strong>Browser + Node</strong>
              <span>the same imports in apps, workers and servers</span>
            </li>
          </ul>
        </div>
      </section>
    </div>
  </main>

  <footer class="footer">
    <div class="container footer-grid">
      <div>
        <h2>vue-pptx-document</h2>
        <p>MIT licensed. Files written by PptxGenJS, flex layout by Yoga.</p>
      </div>
      <nav aria-label="Project links">
        <h2>Project</h2>
        <a href="https://vue-pptx-document-docs.vercel.app">Docs</a>
        <a href="https://github.com/jantursky/vue-pptx-document">GitHub</a>
        <a href="https://www.npmjs.com/package/vue-pptx-document">npm</a>
      </nav>
      <nav aria-label="Built on">
        <h2>Built on</h2>
        <a href="https://github.com/gitbrent/PptxGenJS">PptxGenJS</a>
        <a href="https://github.com/facebook/yoga">Yoga</a>
        <a href="https://vuejs.org">Vue</a>
      </nav>
    </div>
  </footer>
</template>
