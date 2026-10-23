import { describe, expect, it } from 'vitest'

import { isPercent, round, toInches } from '../../src/core/units'
import { cssColor, normalizeColors, opacity, toHex } from '../../src/core/colors'

describe('toInches', () => {
  it('reads numbers as inches and converts units', () => {
    expect(toInches(2)).toBe(2)
    expect(toInches('1.5')).toBe(1.5)
    expect(toInches('2.54cm')).toBeCloseTo(1)
    expect(toInches('25.4 mm')).toBeCloseTo(1)
    expect(toInches('72pt')).toBe(1)
    expect(toInches('96px')).toBe(1)
    expect(toInches('914400emu')).toBe(1)
    expect(toInches('3in')).toBe(3)
  })

  it('takes percentages of the base', () => {
    expect(toInches('50%', 10)).toBe(5)
    expect(toInches('50%')).toBe(0)
    expect(isPercent('25%')).toBe(true)
    expect(isPercent('25')).toBe(false)
  })

  it('rejects what is not a length', () => {
    expect(toInches('auto')).toBeUndefined()
    expect(toInches(Number.NaN)).toBeUndefined()
    expect(toInches(undefined)).toBeUndefined()
  })

  it('rounds to a ten-thousandth', () => {
    expect(round(1 / 3)).toBe(0.3333)
  })
})

describe('colors', () => {
  it('normalizes CSS colors to PptxGenJS hex', () => {
    expect(toHex('#f35815')).toBe('F35815')
    expect(toHex('#fff')).toBe('FFFFFF')
    expect(toHex('11223380')).toBe('112233')
    expect(toHex('white')).toBe('FFFFFF')
    expect(toHex('accent1')).toBe('accent1')
    expect(toHex(3)).toBeUndefined()
  })

  it('walks nested options', () => {
    expect(normalizeColors({ color: '#fff', line: { color: 'black' }, chartColors: ['#f00'], fontSize: 3 })).toEqual({
      color: 'FFFFFF',
      line: { color: '000000' },
      chartColors: ['FF0000'],
      fontSize: 3,
    })
  })

  it('maps colors and transparency for the preview', () => {
    expect(cssColor('fff')).toBe('#FFFFFF')
    expect(cssColor(undefined)).toBe('transparent')
    expect(cssColor('accent1')).toBe('#4472C4')
    expect(opacity(25)).toBe(0.75)
    expect(opacity(undefined)).toBe(1)
  })
})
