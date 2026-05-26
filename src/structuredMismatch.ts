export type StructuredMismatchRow = {
  direction: 'DMS_TO_TALLY' | 'TALLY_TO_DMS'
  billNo: string
  valueName: string
  dmsValue: string
  tallyValue: string
}

export function normalizeCell(v: any): string {
  if (v === null || v === undefined) return ''
  let text = typeof v === 'number' ? String(v) : String(v)
  text = text.replace(/\u00A0/g, ' ')
  text = text.replace(/\u2018|\u2019|\u201C|\u201D/g, "'")
  text = text.replace(/\u2013|\u2014/g, '-')
  return text.trim()
}

