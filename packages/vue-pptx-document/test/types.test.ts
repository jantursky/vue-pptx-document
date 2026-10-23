// The assertions here run at compile time: `bun run typecheck` fails when the public
// types drift. At runtime the file only checks that the exports exist.
import { describe, expectTypeOf, it } from 'vitest'
import { defineComponent, h, type Ref } from 'vue'

import {
  pptx,
  renderToBlob,
  renderToBuffer,
  renderToModel,
  usePptx,
  useSlides,
  Presentation,
  type Chart,
  type ChartProps,
  type DeckModel,
  type PptxInstance,
  type Slide,
  type Text,
  type UsePptxReturn,
} from '../src'
import { pptxToVue, type PptxToVueResult } from '../src/import'
import Deck from './fixtures/Deck.vue'

type PropsOf<T extends abstract new (...args: never) => { $props: unknown }> = InstanceType<T>['$props']

describe('public types', () => {
  it('types primitive props', () => {
    expectTypeOf<PropsOf<typeof Slide>>().toHaveProperty('layout')
    expectTypeOf<PropsOf<typeof Text>>().toHaveProperty('fontSize')
    expectTypeOf<PropsOf<typeof Chart>>().toHaveProperty('type')

    // @ts-expect-error layout is absolute or flex
    expectTypeOf<PropsOf<typeof Slide>['layout']>().toEqualTypeOf<'grid' | undefined>()
  })

  it('types the render functions, the composables and the importer', () => {
    expectTypeOf(renderToBuffer).returns.toEqualTypeOf<Promise<Uint8Array>>()
    expectTypeOf(renderToBlob).returns.toEqualTypeOf<Promise<Blob>>()
    expectTypeOf(renderToModel).returns.toEqualTypeOf<Promise<DeckModel>>()
    expectTypeOf(pptx).returns.toEqualTypeOf<PptxInstance>()
    expectTypeOf(usePptx).returns.toEqualTypeOf<UsePptxReturn>()
    expectTypeOf<UsePptxReturn['url']>().toEqualTypeOf<Readonly<Ref<string | null>>>()
    expectTypeOf(useSlides).returns.toHaveProperty('model')
    expectTypeOf(pptxToVue).returns.toEqualTypeOf<Promise<PptxToVueResult>>()
    expectTypeOf(renderToBuffer).toBeCallableWith(Deck, { title: 'T', points: [] })
  })

  it('types document props from the component', () => {
    // Deck.vue declares `title: string` and `points: string[]`, both required.
    expectTypeOf(renderToBuffer).toBeCallableWith(Deck, { title: 'T', points: [] }, { strict: true })
    const Runtime = defineComponent({
      props: { count: { type: Number, required: true } },
      setup: () => () => h(Presentation),
    })
    // Never called: these lines only have to fail the type check.
    const wrong = () => {
      // @ts-expect-error title is a string
      void renderToBuffer(Deck, { title: 1, points: [] })
      // @ts-expect-error points is required
      void renderToModel(Deck, { title: 'T' })
      // @ts-expect-error the props are required when the deck has required props
      void renderToBlob(Deck)
      // @ts-expect-error the same for the composables and pptx()
      usePptx(Deck)
      // @ts-expect-error count is a number
      void renderToModel(Runtime, { count: 'two' })
      // @ts-expect-error showLegnd is not a chart option
      const typo: ChartProps = { type: 'bar', showLegnd: true }
      return typo
    }
    expectTypeOf(wrong).toBeFunction()
    expectTypeOf({
      type: 'bar',
      showLegend: true,
      valAxisTitle: 'Revenue',
      catAxisLabelRotate: 45,
    } as const).toExtend<ChartProps>()
    expectTypeOf(useSlides).toBeCallableWith(Deck, () => ({ title: 'T', points: ['a'] }), { debounce: 0 })
    expectTypeOf(pptx).toBeCallableWith(Deck, { title: 'T', points: [] })

    expectTypeOf(renderToModel).toBeCallableWith(Runtime, { count: 2 })

    // Without declared props, any object (or nothing) goes.
    const Bare = defineComponent(() => () => h(Presentation))
    expectTypeOf(renderToModel).toBeCallableWith(Bare)
    expectTypeOf(renderToModel).toBeCallableWith(Bare, { anything: true })
  })
})
