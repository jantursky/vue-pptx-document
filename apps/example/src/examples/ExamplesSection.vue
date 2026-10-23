<script setup lang="ts">
import ChartsDeck from './ChartsDeck.vue?raw'
import ChartsExample from './ChartsExample.vue'
import chartsSource from './ChartsExample.vue?raw'
import ExampleCard from './ExampleCard.vue'
import HelloDeck from './HelloDeck.vue?raw'
import LayoutDeck from './LayoutDeck.vue?raw'
import LayoutExample from './LayoutExample.vue'
import layoutSource from './LayoutExample.vue?raw'
import MinimalExample from './MinimalExample.vue'
import minimalSource from './MinimalExample.vue?raw'
import report from './report.ts?raw'
import ReportDeck from './ReportDeck.vue?raw'
import ReportExample from './ReportExample.vue'
import reportSource from './ReportExample.vue?raw'
</script>

<template>
  <section class="examples" aria-labelledby="examples-title">
    <header class="section-head examples-head">
      <p class="tag">Examples</p>
      <h2 id="examples-title">Four decks, live</h2>
      <p>
        Each example is a presentation component plus the page that shows it. The preview is HTML drawn from the same
        tree PptxGenJS writes, so it updates as you type. Open <em>View source</em> to read the exact files that run
        here.
      </p>
      <nav class="chips" aria-label="Examples">
        <a class="chip" href="#example-minimal">Minimal</a>
        <a class="chip" href="#example-layout">Absolute and flex</a>
        <a class="chip" href="#example-charts">Charts</a>
        <a class="chip" href="#example-report">Masters and tables</a>
      </nav>
    </header>

    <div class="examples-grid">
      <ExampleCard
        id="example-minimal"
        title="Minimal"
        :files="[
          { name: 'HelloDeck.vue', code: HelloDeck },
          { name: 'MinimalExample.vue', code: minimalSource },
        ]"
      >
        <template #description>
          <p>One slide and a prop. Type a name and the preview draws it again; the button writes the .pptx.</p>
        </template>
        <MinimalExample />
      </ExampleCard>

      <ExampleCard
        id="example-layout"
        title="Absolute and flex"
        :files="[
          { name: 'LayoutDeck.vue', code: LayoutDeck },
          { name: 'LayoutExample.vue', code: layoutSource },
        ]"
      >
        <template #description>
          <p>
            The same cards twice: placed with <code>x</code>, <code>y</code>, <code>w</code> and <code>h</code> in
            inches, then laid out by Yoga inside <code>&lt;View&gt;</code> with <code>flex</code> and <code>gap</code>.
          </p>
        </template>
        <LayoutExample />
      </ExampleCard>

      <ExampleCard
        id="example-charts"
        title="Charts"
        :files="[
          { name: 'ChartsDeck.vue', code: ChartsDeck },
          { name: 'ChartsExample.vue', code: chartsSource },
        ]"
      >
        <template #description>
          <p>
            Every PptxGenJS chart type, plus a bar and line combo on two axes. The file gets native, editable PowerPoint
            charts.
          </p>
        </template>
        <ChartsExample />
      </ExampleCard>

      <ExampleCard
        id="example-report"
        title="Masters and tables"
        :files="[
          { name: 'ReportDeck.vue', code: ReportDeck },
          { name: 'ReportExample.vue', code: reportSource },
          { name: 'report.ts', code: report },
        ]"
      >
        <template #description>
          <p>
            A <code>&lt;Master&gt;</code> with a footer and slide numbers, sections, speaker notes, bullets and a table
            built from <code>&lt;TableRow&gt;</code> and <code>v-for</code>.
          </p>
        </template>
        <ReportExample />
      </ExampleCard>
    </div>
  </section>
</template>
