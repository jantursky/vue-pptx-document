// The assertions here run at compile time: `bun run typecheck` fails when the public
// types drift. At runtime the file only checks that the exports exist.
import { describe, expectTypeOf, it } from 'vitest'
import type { Ref } from 'vue'

import {
  pptx,
  renderToBlob,
  renderToBuffer,
  renderToModel,
  usePptx,
  useSlides,
  type Chart,
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
})
