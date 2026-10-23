export interface Kpi {
  label: string
  value: string
  /** Change against the previous period, in percent. */
  delta: number
}

export interface Series {
  name: string
  values: number[]
}

export interface Region {
  name: string
  revenue: number
  target: number
}

export type ChartKind = 'bar' | 'line' | 'area'
export type DeckLayout = '16x9' | '4x3' | 'wide'

export interface Report {
  layout: DeckLayout
  title: string
  subtitle: string
  company: string
  author: string
  date: string
  accent: string
  /** A PNG or JPEG data URI. */
  logo: string | null
  highlights: string[]
  kpis: Kpi[]
  months: string[]
  series: Series[]
  chart: ChartKind
  regions: Region[]
  closing: string
  notes: boolean
}

export const STORAGE_KEY = 'vue-pptx-document:report'
export const MAX_KPIS = 4
export const MAX_SERIES = 3
export const MAX_REGIONS = 8

const today = (): string => new Date().toISOString().slice(0, 10)

export const sampleReport = (): Report => ({
  layout: '16x9',
  title: 'Q3 business review',
  subtitle: 'Revenue, customers and what comes next',
  company: 'Northwind Studio',
  author: 'Ada Novak',
  date: today(),
  accent: '#F35815',
  logo: null,
  highlights: [
    'Revenue grew 18% over last quarter',
    'The new customer portal launched in September',
    'Churn dropped below 2%',
    'The expansion to Austria is on track',
  ],
  kpis: [
    { label: 'Revenue', value: '€2.4M', delta: 18 },
    { label: 'Customers', value: '1,284', delta: 9 },
    { label: 'NPS', value: '61', delta: 4 },
    { label: 'Churn', value: '1.8%', delta: -12 },
  ],
  months: ['Jul', 'Aug', 'Sep'],
  series: [
    { name: 'This year', values: [720, 790, 890] },
    { name: 'Last year', values: [610, 640, 700] },
  ],
  chart: 'bar',
  regions: [
    { name: 'Slovakia', revenue: 940, target: 900 },
    { name: 'Czechia', revenue: 760, target: 800 },
    { name: 'Austria', revenue: 420, target: 380 },
    { name: 'Poland', revenue: 280, target: 350 },
  ],
  closing: 'Next quarter: a partner portal and an office in Vienna.',
  notes: true,
})

export const emptyReport = (): Report => ({
  ...sampleReport(),
  title: '',
  subtitle: '',
  company: '',
  author: '',
  logo: null,
  highlights: [''],
  kpis: [{ label: '', value: '', delta: 0 }],
  series: [{ name: '', values: [0, 0, 0] }],
  regions: [{ name: '', revenue: 0, target: 0 }],
  closing: '',
})

const isString = (value: unknown): value is string => typeof value === 'string'
const isNumber = (value: unknown): value is number => typeof value === 'number' && Number.isFinite(value)

/** Checks a value read from storage or an imported JSON file. */
export const isReport = (value: unknown): value is Report => {
  if (!value || typeof value !== 'object') return false
  const report = value as Record<string, unknown>
  return (
    ['16x9', '4x3', 'wide'].includes(report.layout as string) &&
    ['title', 'subtitle', 'company', 'author', 'date', 'accent', 'closing'].every((key) => isString(report[key])) &&
    (report.logo === null || (isString(report.logo) && report.logo.startsWith('data:image/'))) &&
    Array.isArray(report.highlights) &&
    report.highlights.every(isString) &&
    Array.isArray(report.kpis) &&
    report.kpis.every((kpi: unknown) => {
      const item = kpi as Record<string, unknown>
      return isString(item.label) && isString(item.value) && isNumber(item.delta)
    }) &&
    Array.isArray(report.months) &&
    report.months.every(isString) &&
    Array.isArray(report.series) &&
    report.series.every((series: unknown) => {
      const item = series as Record<string, unknown>
      return isString(item.name) && Array.isArray(item.values) && item.values.every(isNumber)
    }) &&
    ['bar', 'line', 'area'].includes(report.chart as string) &&
    Array.isArray(report.regions) &&
    report.regions.every((region: unknown) => {
      const item = region as Record<string, unknown>
      return isString(item.name) && isNumber(item.revenue) && isNumber(item.target)
    }) &&
    typeof report.notes === 'boolean'
  )
}

export const loadDraft = (): Report | null => {
  try {
    const parsed: unknown = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? 'null')
    return isReport(parsed) ? parsed : null
  } catch {
    return null
  }
}

export const saveDraft = (report: Report): void => {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(report))
  } catch {
    // Private windows and full storage: the draft just isn't kept.
  }
}
