import { XMLParser } from 'fast-xml-parser'

/** An element as fast-xml-parser's `preserveOrder` mode returns it: `{ [tag]: children, ':@': attributes }`. */
export type XmlNode = Record<string, unknown>

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  preserveOrder: true,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  processEntities: true,
  htmlEntities: true,
})

/** Parses a part and returns its root element. */
export const parseXml = (source: string): XmlNode | undefined =>
  (parser.parse(source) as XmlNode[]).find((node) => {
    const name = tagOf(node)
    return name !== undefined && !name.startsWith('?')
  })

export const tagOf = (node: XmlNode): string | undefined => Object.keys(node).find((key) => key !== ':@')

export const childrenOf = (node: XmlNode | undefined): XmlNode[] => {
  if (!node) return []
  const name = tagOf(node)
  const value = name ? node[name] : undefined
  return Array.isArray(value) ? (value as XmlNode[]) : []
}

export const attr = (node: XmlNode | undefined, name: string): string | undefined => {
  const attributes = node?.[':@'] as Record<string, string> | undefined
  return attributes?.[name]
}

export const numAttr = (node: XmlNode | undefined, name: string): number | undefined => {
  const value = attr(node, name)
  if (value === undefined) return undefined
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : undefined
}

export const child = (node: XmlNode | undefined, name: string): XmlNode | undefined =>
  childrenOf(node).find((item) => tagOf(item) === name)

export const all = (node: XmlNode | undefined, name: string): XmlNode[] =>
  childrenOf(node).filter((item) => tagOf(item) === name)

/** Follows a path of child tags. */
export const path = (node: XmlNode | undefined, ...names: string[]): XmlNode | undefined =>
  names.reduce<XmlNode | undefined>((current, name) => child(current, name), node)

/** All descendants with the tag, depth first. */
export const descendants = (node: XmlNode | undefined, name: string): XmlNode[] => {
  const out: XmlNode[] = []
  const walk = (current: XmlNode): void => {
    for (const item of childrenOf(current)) {
      if (tagOf(item) === name) out.push(item)
      walk(item)
    }
  }
  if (node) walk(node)
  return out
}

/** The text content of an element. */
export const textOf = (node: XmlNode | undefined): string => {
  if (!node) return ''
  let out = ''
  for (const item of childrenOf(node)) {
    if (tagOf(item) === '#text') out += String(item['#text'])
    else out += textOf(item)
  }
  return out
}
