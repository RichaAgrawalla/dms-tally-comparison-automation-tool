import React, { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { exportReportPdf } from './DownloadPdf'

type Row = Record<string, unknown>

type MismatchRow = { billNo: string; valueName: string; dmsValue: string; tallyValue: string }

const DEFAULT_MAPPING_PAIRS: Array<{ dms: string; tally: string }> = [
  { dms: 'Bill Num', tally: 'Voucher No.' },
  { dms: 'Bill Date', tally: 'Date' },
  { dms: 'ins deductant', tally: 'Deductible' },
  { dms: 'Gross taxable Part', tally: 'Workshop Spare Sale @18%' },
  { dms: 'Gross labor', tally: 'Workshop Labour Service' },
  { dms: 'CGST 9% total', tally: 'Output CGST @9%' },
  { dms: 'SGST 9% total', tally: 'Output SGST @9%' },
  { dms: 'IGST 18% total', tally: 'Output IGST @18%' },
  { dms: 'TCS total', tally: 'TCS Payable @ 0.1%' },
  { dms: 'Net bill amt total', tally: 'Gross Total' }
]

function DifferencesTable({ mismatchRows }: { mismatchRows: MismatchRow[] }) {
  if (!mismatchRows || mismatchRows.length === 0) {
    return <p style={{ color: '#9ca3af' }}>No mismatches found!</p>
  }

  return (
    <div style={{ overflowX: 'auto', marginBottom: 16 }}>
      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
        <thead>
          <tr style={{ backgroundColor: '#1e293b', borderBottom: '2px solid #3b82f6' }}>
            <th style={{ padding: 8, textAlign: 'left', color: '#fff', fontWeight: 600 }}>Bill No</th>
            <th style={{ padding: 8, textAlign: 'left', color: '#fff', fontWeight: 600 }}>Value Name</th>
            <th style={{ padding: 8, textAlign: 'left', color: '#fff', fontWeight: 600 }}>DMS Value</th>
            <th style={{ padding: 8, textAlign: 'left', color: '#fff', fontWeight: 600 }}>Tally Value</th>
          </tr>
        </thead>
        <tbody>
          {mismatchRows.map((row, idx) => (
            <tr
              key={idx}
              style={{
                backgroundColor: idx % 2 === 0 ? '#0f1730' : '#1a2244',
                borderBottom: '1px solid #2d3748'
              }}
            >
              <td style={{ padding: 8 }}>{row.billNo}</td>
              <td style={{ padding: 8 }}>{row.valueName}</td>
              <td style={{ padding: 8, color: '#fca5a5' }}>{row.dmsValue}</td>
              <td style={{ padding: 8, color: '#86efac' }}>{row.tallyValue}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}



type Dataset = {
  name: 'DMS' | 'Tally'
  raw: Row[]
  columns: string[]
}

const ZERO_COMPARE = '__ZERO_COMPARE__'

type Mapping = {
  dmsCol: string
  tallyCol: string
}

function normalizeCell(v: any): string {
  if (v === null || v === undefined) return ''
  let text = typeof v === 'number' ? String(v) : String(v)
  text = text.replace(/\u00A0/g, ' ')
  text = text.replace(/\u2018|\u2019|\u201C|\u201D/g, "'")
  text = text.replace(/\u2013|\u2014/g, '-')
  return text.trim()
}

function areValuesEquivalent(a: any, b: any): boolean {
  const left = normalizeCell(a)
  const right = normalizeCell(b)
  if (!left && !right) return true

  const numericLeft = Number(left.replace(/,/g, ''))
  const numericRight = Number(right.replace(/,/g, ''))
  if (!Number.isNaN(numericLeft) && !Number.isNaN(numericRight)) {
    return Math.abs(numericLeft - numericRight) < 1
  }

  return left.toLowerCase() === right.toLowerCase()
}

function numericDifference(a: any, b: any): string | null {
  const left = normalizeCell(a)
  const right = normalizeCell(b)
  const numericLeft = Number(left.replace(/,/g, ''))
  const numericRight = Number(right.replace(/,/g, ''))
  if (Number.isNaN(numericLeft) || Number.isNaN(numericRight)) return null

  const diff = numericLeft - numericRight
  if (Math.abs(diff) < 1) return null
  return diff >= 0 ? `+${diff}` : `${diff}`
}

function isBlankCell(v: any): boolean {
  const s = normalizeCell(v)
  if (!s) return true
  // treat common placeholders as blank
  return ['na', 'n/a', 'null', 'undefined', '-'].includes(s.toLowerCase())
}

function shouldKeepColumn(v: any): boolean {
  // Keep a column if it has at least one cell that is either:
  // - non-blank, AND
  // - NOT zero-like (0 / 0.0 / "0")
  //
  // So we drop columns that are *entirely* blank or entirely zero-like.
  if (isBlankCell(v)) return false
  if (isZeroLike(v)) return false
  return true
}


function isZeroLike(v: any): boolean {
  const s = normalizeCell(v)
  if (!s) return true
  const num = Number(s)
  // If it's numeric and equals 0 -> zero-like.
  if (!Number.isNaN(num)) return num === 0
  // Non-numeric strings are not considered zero.
  return false
}


function computeNonEmptyColumns(rows: Row[], allCols: string[], preservedCols: Set<string> = new Set()): { keep: string[]; removed: string[] } {
  const keep: string[] = []
  const removed: string[] = []

  for (const col of allCols) {
    let hasKeepValue = preservedCols.has(col)
    if (!hasKeepValue) {
      for (const r of rows) {
        const v = (r as any)[col]
        if (shouldKeepColumn(v)) {
          hasKeepValue = true
          break
        }
      }
    }

    if (hasKeepValue) keep.push(col)
    else removed.push(col)
  }

  if (removed.length > 0) {
    // eslint-disable-next-line no-console
    console.info('[dmstally] Removed columns:', removed)
  }

  return { keep, removed }
}

function createComparisonTitleRow(dmsColumns: string[], tallyColumns: string[]): string[] {
  return [
    'DMS',
    ...new Array(Math.max(0, dmsColumns.length - 1)).fill(''),
    '',
    'TALLY',
    ...new Array(Math.max(0, tallyColumns.length - 1)).fill(''),
    '',
    'DMS TO TALLY',
    ...new Array(Math.max(0, dmsColumns.length - 1)).fill(''),
    '',
    'TALLY TO DMS',
    ...new Array(Math.max(0, dmsColumns.length - 1)).fill('')
  ]
}

function buildComparisonExcelSheetRows(
  dms: Dataset,
  tally: Dataset,
  mappings: Mapping[]
): Array<Array<string>> {
  const baseKeyMapping = mappings[0]
  const dmsToTallyResult = compareDirection(dms, tally, baseKeyMapping, mappings, false)

  const dmsColumns = dms.columns
  const tallyColumns = tally.columns

  const titleRow = createComparisonTitleRow(dmsColumns, tallyColumns)
  const headerRow = [
    ...dmsColumns,
    '',
    ...tallyColumns,
    '',
    ...dmsColumns,
    '',
    ...dmsColumns
  ]

  const rows: Array<Array<string>> = [titleRow, headerRow]

  const numericDiff = (a: string, b: string): string => {
    const left = Number(a.toString().replace(/,/g, ''))
    const right = Number(b.toString().replace(/,/g, ''))
    if (Number.isNaN(left) || Number.isNaN(right)) return ''
    const diff = left - right
    return diff === 0 ? '' : diff.toString()
  }

  const toDiffRow = (fromRow: Row, toRow: Row, sign: 1 | -1): string[] => {
    return dmsColumns.map((col) => {
      if (col === baseKeyMapping.dmsCol) {
        return getRowValue(fromRow, col)
      }

      const mapping = mappings.find((m) => m.dmsCol === col)
      if (!mapping || mapping.dmsCol === ZERO_COMPARE || mapping.tallyCol === ZERO_COMPARE) return ''

      const fromVal = getRowValue(fromRow, mapping.dmsCol)
      const toVal = getRowValue(toRow, mapping.tallyCol)
      const diff = numericDiff(fromVal, toVal)
      if (!diff) return ''
      const numeric = Number(diff)
      return Number.isNaN(numeric) ? diff : (numeric * sign).toString()
    })
  }

  for (const item of dmsToTallyResult.matched) {
    const dmsRow = dmsColumns.map((col) => getRowValue(item.fromRow, col))
    const tallyRow = tallyColumns.map((col) => getRowValue(item.toRow, col))
    const dmsToTallyDiffRow = toDiffRow(item.fromRow, item.toRow, 1)
    const tallyToDmsDiffRow = toDiffRow(item.fromRow, item.toRow, -1)

    rows.push([
      ...dmsRow,
      '',
      ...tallyRow,
      '',
      ...dmsToTallyDiffRow,
      '',
      ...tallyToDmsDiffRow
    ])
  }

  for (const item of dmsToTallyResult.unmatched) {
    if (item.sourceDataset !== 'DMS') continue
    const dmsRow = dmsColumns.map((col) => getRowValue(item.row, col))
    rows.push([
      ...dmsRow,
      '',
      ...new Array(tallyColumns.length).fill(''),
      '',
      ...new Array(dmsColumns.length).fill(''),
      '',
      ...new Array(dmsColumns.length).fill('')
    ])
  }

  const tallyUnmatched = compareDirection(tally, dms, baseKeyMapping, mappings, true).unmatched
  for (const item of tallyUnmatched) {
    if (item.sourceDataset !== 'DMS') continue
    const tallyRow = tallyColumns.map((col) => getRowValue(item.row, col))
    rows.push([
      ...new Array(dmsColumns.length).fill(''),
      '',
      ...tallyRow,
      '',
      ...new Array(dmsColumns.length).fill(''),
      '',
      ...new Array(dmsColumns.length).fill('')
    ])
  }

  return rows
}

function buildFilteredExcelSheetRows(
  dms: Dataset,
  tally: Dataset,
  mappings: Mapping[]
): Array<Array<string>> {
  const baseKeyMapping = mappings[0]
  const dmsToTallyResult = compareDirection(dms, tally, baseKeyMapping, mappings, false)

  const dmsColumns = dms.columns
  const tallyColumns = tally.columns
  const titleRow = createComparisonTitleRow(dmsColumns, tallyColumns)
  const headerRow = [
    ...dmsColumns,
    '',
    ...tallyColumns,
    '',
    ...dmsColumns,
    '',
    ...dmsColumns
  ]

  const rows: Array<Array<string>> = [titleRow, headerRow]

  const numericDiff = (a: string, b: string): string => {
    const left = Number(a.toString().replace(/,/g, ''))
    const right = Number(b.toString().replace(/,/g, ''))
    if (Number.isNaN(left) || Number.isNaN(right)) return ''
    const diff = left - right
    return diff === 0 ? '' : diff.toString()
  }

  const significantValue = (value: string): boolean => {
    const parsed = Number(value.toString().replace(/,/g, ''))
    return !Number.isNaN(parsed) && Math.abs(parsed) >= 1
  }

  const toDiffRow = (fromRow: Row, toRow: Row, sign: 1 | -1): string[] => {
    return dmsColumns.map((col) => {
      if (col === baseKeyMapping.dmsCol) {
        return getRowValue(fromRow, col)
      }

      const mapping = mappings.find((m) => m.dmsCol === col)
      if (!mapping || mapping.dmsCol === ZERO_COMPARE || mapping.tallyCol === ZERO_COMPARE) return ''

      const fromVal = getRowValue(fromRow, mapping.dmsCol)
      const toVal = getRowValue(toRow, mapping.tallyCol)
      const diff = numericDiff(fromVal, toVal)
      if (!diff) return ''
      const numeric = Number(diff)
      return Number.isNaN(numeric) ? diff : (numeric * sign).toString()
    })
  }

  for (const item of dmsToTallyResult.matched) {
    const dmsRow = dmsColumns.map((col) => getRowValue(item.fromRow, col))
    const tallyRow = tallyColumns.map((col) => getRowValue(item.toRow, col))
    const dmsToTallyDiffRow = toDiffRow(item.fromRow, item.toRow, 1)
    const tallyToDmsDiffRow = toDiffRow(item.fromRow, item.toRow, -1)

    const hasSignificantValue = dmsToTallyDiffRow.some((val) => significantValue(val)) || tallyToDmsDiffRow.some((val) => significantValue(val))
    if (!hasSignificantValue) continue

    rows.push([
      ...dmsRow,
      '',
      ...tallyRow,
      '',
      ...dmsToTallyDiffRow,
      '',
      ...tallyToDmsDiffRow
    ])
  }

  return rows
}

function downloadComparisonExcel(
  dms: Dataset,
  tally: Dataset,
  mappings: Mapping[]
) {
  const comparisonRows = buildComparisonExcelSheetRows(dms, tally, mappings)
  const filteredRows = buildFilteredExcelSheetRows(dms, tally, mappings)
  const wsComparison = XLSX.utils.aoa_to_sheet(comparisonRows)
  const wsFiltered = XLSX.utils.aoa_to_sheet(filteredRows)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, wsComparison, 'Comparison')
  XLSX.utils.book_append_sheet(wb, wsFiltered, 'Filtered >=1')
  XLSX.writeFile(wb, `dmstally-comparison-${new Date().toISOString().slice(0, 10)}.xlsx`)
}


function loadExcel(file: File, sheetIndex = 0): Promise<{ rows: Row[]; cols: string[] }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onerror = () => reject(reader.error)
    reader.onload = () => {
      try {
        const data = new Uint8Array(reader.result as ArrayBuffer)
        const wb = XLSX.read(data, { type: 'array' })
        const sheetName = wb.SheetNames[sheetIndex]
        const ws = wb.Sheets[sheetName]
        // defval ensures empty cells appear as ''
        const json = XLSX.utils.sheet_to_json(ws, { defval: '' }) as Row[]
        const allCols = Array.from(
          new Set(
            json.flatMap((r) => Object.keys(r))
          )
        )
        resolve({ rows: json, cols: allCols })
      } catch (e) {
        reject(e)
      }
    }
    reader.readAsArrayBuffer(file)
  })
}

function buildKey(row: Row, cols: string[]): string {
  return cols.map((c) => normalizeCell(row[c])).join('||').toLowerCase()
}

type ComparisonResult = {
  unmatched: Array<{ row: Row; source: 'from' | 'to'; sourceDataset: 'DMS' | 'Tally' }>
  matched: Array<{
    fromRow: Row
    toRow: Row
    mismatches: Array<{ field: string; fromValue: string; toValue: string; difference: string }>
    mappedMatches: Array<{ field: string; value: string }>
  }>
}

function formatMappingColumn(col: string, side: 'DMS' | 'Tally'): string {
  if (col === ZERO_COMPARE) return '0 (no match)'
  return `${side}[${col}]`
}

function getRowValue(row: Row, col: string): string {
  if (col === ZERO_COMPARE) return '0'
  return normalizeCell(row[col])
}

function compareDirection(
  from: Dataset,
  to: Dataset,
  baseKeyMapping: Mapping,
  allMappings: Mapping[],
  reverse: boolean
): ComparisonResult {
  const fromKeyCol = reverse ? baseKeyMapping.tallyCol : baseKeyMapping.dmsCol
  const toKeyCol = reverse ? baseKeyMapping.dmsCol : baseKeyMapping.tallyCol

  const fromToKey = (row: Row) => normalizeCell(row[fromKeyCol])
  const toToKey = (row: Row) => normalizeCell(row[toKeyCol])

  const toIndex = new Map<string, Row[]>()
  for (const r of to.raw) {
    const k = toToKey(r)
    if (!k) continue
    const bucket = toIndex.get(k) ?? []
    bucket.push(r)
    toIndex.set(k, bucket)
  }

  const unmatched: Array<{ row: Row; source: 'from' | 'to'; sourceDataset: 'DMS' | 'Tally' }> = []
  const matched: Array<{
    fromRow: Row
    toRow: Row
    mismatches: Array<{ field: string; fromValue: string; toValue: string; difference: string }>
    mappedMatches: Array<{ field: string; value: string }>
  }> = []

  const matchedToRows = new Set<Row>()

  for (const fr of from.raw) {
    const k = fromToKey(fr)
    if (!k) {
      unmatched.push({ row: fr, source: 'from', sourceDataset: reverse ? 'Tally' : 'DMS' })
      continue
    }

    const bucket = toIndex.get(k) ?? []
    const tr = bucket.find((row) => !matchedToRows.has(row))
    if (!tr) {
      unmatched.push({ row: fr, source: 'from', sourceDataset: reverse ? 'Tally' : 'DMS' })
      continue
    }

    matchedToRows.add(tr)

    const mismatches: Array<{ field: string; fromValue: string; toValue: string; difference: string }> = []
    const mappedMatches: Array<{ field: string; value: string }> = []

    for (const m of allMappings) {
      const fromCol = reverse ? m.tallyCol : m.dmsCol
      const toCol = reverse ? m.dmsCol : m.tallyCol
      const fromVal = getRowValue(fr, fromCol)
      const toVal = getRowValue(tr, toCol)

      if (!areValuesEquivalent(fromVal, toVal)) {
        const difference = numericDifference(fromVal, toVal)
        mismatches.push({
          field: `${formatMappingColumn(fromCol, reverse ? 'Tally' : 'DMS')} ↔ ${formatMappingColumn(toCol, reverse ? 'DMS' : 'Tally')}`,
          fromValue: fromVal,
          toValue: toVal,
          difference: difference ?? ''
        })
      } else if (fromVal) {
        mappedMatches.push({
          field: `${formatMappingColumn(fromCol, reverse ? 'Tally' : 'DMS')} ↔ ${formatMappingColumn(toCol, reverse ? 'DMS' : 'Tally')}`,
          value: fromVal
        })
      }
    }

    matched.push({ fromRow: fr, toRow: tr, mismatches, mappedMatches })
  }

  for (const tr of to.raw) {
    const k = toToKey(tr)
    if (!k || !matchedToRows.has(tr)) {
      unmatched.push({ row: tr, source: 'to', sourceDataset: reverse ? 'DMS' : 'Tally' })
    }
  }

  return { unmatched, matched }
}

export default function App() {
  const [dmsFile, setDmsFile] = useState<File | null>(null)
  const [tallyFile, setTallyFile] = useState<File | null>(null)

  const [dms, setDms] = useState<Dataset | null>(null)
  const [tally, setTally] = useState<Dataset | null>(null)

  const [mappings, setMappings] = useState<Mapping[]>([])

  const [reportText, setReportText] = useState<string>('')
  const [mismatchTables, setMismatchTables] = useState<Array<{
    title: string
    rows: Array<{ billNo: string; valueName: string; dmsValue: string; tallyValue: string }>
  }>>([])
  const [dmsRemovedCols, setDmsRemovedCols] = useState<string[]>([])
  const [tallyRemovedCols, setTallyRemovedCols] = useState<string[]>([])

  const [busy, setBusy] = useState(false)


  const dmsCols = dms?.columns ?? []
  const tallyCols = tally?.columns ?? []

  const usedDmsCols = new Set(mappings.filter((m) => m.dmsCol !== ZERO_COMPARE).map((m) => m.dmsCol))
  const usedTallyCols = new Set(mappings.filter((m) => m.tallyCol !== ZERO_COMPARE).map((m) => m.tallyCol))


  const suggestedMappingCount = 3

  const allMappedUnique = useMemo(() => {
    const set = new Set(mappings.map((m) => `${m.dmsCol}::${m.tallyCol}`))
    return set.size === mappings.length
  }, [mappings])

  async function handleLoad(which: 'DMS' | 'Tally') {
    const file = which === 'DMS' ? dmsFile : tallyFile
    if (!file) return

    setBusy(true)
    try {
      const loaded = await loadExcel(file)
      const preservedCols = new Set<string>()
      const preservedKeys = new Set<string>()
      if (which === 'DMS') {
        for (const pair of DEFAULT_MAPPING_PAIRS) {
          preservedKeys.add(normalizeCell(pair.dms).toLowerCase())
        }
      } else {
        for (const pair of DEFAULT_MAPPING_PAIRS) {
          preservedKeys.add(normalizeCell(pair.tally).toLowerCase())
        }
      }
      for (const col of loaded.cols) {
        if (preservedKeys.has(normalizeCell(col).toLowerCase())) {
          preservedCols.add(col)
        }
      }

      const { keep, removed } = computeNonEmptyColumns(loaded.rows, loaded.cols, preservedCols)
      const normalizedRows = loaded.rows.map((r) => {
        const out: Row = {}
        for (const c of keep) out[c] = r[c]
        return out
      })

      const dataset: Dataset = {
        name: which,
        raw: normalizedRows,
        columns: keep
      }

      if (which === 'DMS') {
        setDmsRemovedCols(removed)
      } else {
        setTallyRemovedCols(removed)
      }

      if (which === 'DMS') setDms(dataset)
      else setTally(dataset)
    } finally {
      setBusy(false)
    }
  }

  function ensureDefaultMappings() {
    if (!dms || !tally) return
    if (mappings.length > 0) return

    const tallyIndexByNorm = new Map<string, string>()
    for (const tCol of tally.columns) {
      tallyIndexByNorm.set(normalizeCell(tCol).toLowerCase(), tCol)
    }

    const dmsIndexByNorm = new Map<string, string>()
    for (const dCol of dms.columns) {
      dmsIndexByNorm.set(normalizeCell(dCol).toLowerCase(), dCol)
    }

    const explicit: Mapping[] = []
    for (const p of DEFAULT_MAPPING_PAIRS) {
      const dKey = normalizeCell(p.dms).toLowerCase()
      const tKey = normalizeCell(p.tally).toLowerCase()
      const dReal = dmsIndexByNorm.get(dKey)
      const tReal = tallyIndexByNorm.get(tKey)
      if (dReal && tReal) explicit.push({ dmsCol: dReal, tallyCol: tReal })
    }

    // Fill any remaining slots using identical-name heuristic.
    const exact: Mapping[] = []
    const usedD = new Set(explicit.map((m) => m.dmsCol))
    const usedT = new Set(explicit.map((m) => m.tallyCol))

    for (const dCol of dms.columns) {
      if (exact.length + explicit.length >= suggestedMappingCount) break
      if (usedD.has(dCol)) continue
      const idx = tally.columns.findIndex(
        (tCol) => normalizeCell(tCol).toLowerCase() === normalizeCell(dCol).toLowerCase()
      )
      if (idx >= 0) {
        const tCol = tally.columns[idx]
        if (usedT.has(tCol)) continue
        exact.push({ dmsCol: dCol, tallyCol: tCol })
        usedD.add(dCol)
        usedT.add(tCol)
      }
    }

    setMappings([...explicit, ...exact])
  }


  function buildReportText(
    dmsToTallyResult: ComparisonResult,
    tallyToDmsResult: ComparisonResult,
    baseKeyMapping: Mapping
  ) {
    const lines: string[] = []

    lines.push('=' .repeat(80))
    lines.push('DMS ↔ TALLY COMPARISON REPORT (Frontend-only, Privacy-first)')
    lines.push('=' .repeat(80))
    lines.push('')
    lines.push(`Report Generated: ${new Date().toLocaleString()}`)
    lines.push('')

    // Base key mapping
    lines.push('BASE KEY MAPPING (used to match records):')
    lines.push(`  DMS[${baseKeyMapping.dmsCol}] ↔ Tally[${baseKeyMapping.tallyCol}]`)
    lines.push('')

    // Other mapped columns
    if (mappings.length > 1) {
      lines.push('OTHER MAPPED COLUMNS (checked for mismatches):')
      for (let i = 1; i < mappings.length; i++) {
        const m = mappings[i]
        lines.push(`  ${i}. DMS[${m.dmsCol}] ↔ Tally[${m.tallyCol}]`)
      }
      lines.push('')
    }

    // Summary statistics
    lines.push('-'.repeat(80))
    lines.push('SUMMARY STATISTICS')
    lines.push('-'.repeat(80))

    const totalDmsRows = dms?.raw.length ?? 0
    const totalTallyRows = tally?.raw.length ?? 0

    lines.push(`Total DMS records:          ${totalDmsRows}`)
    lines.push(`Total Tally records:        ${totalTallyRows}`)
    lines.push('')

    const dmsMatched = dmsToTallyResult.matched.length
    const dmsUnmatched = dmsToTallyResult.unmatched.length
    const tallyMatched = tallyToDmsResult.matched.length
    const tallyUnmatched = tallyToDmsResult.unmatched.length

    lines.push('[DMS → TALLY]')
    lines.push(`  ✓ Matched records:         ${dmsMatched}`)
    lines.push(`  ✗ Unmatched records:       ${dmsUnmatched}`)
    
    const dmsMismatchCount = dmsToTallyResult.matched.filter(m => m.mismatches.length > 0).length
    lines.push(`  ⚠ Matched with mismatches: ${dmsMismatchCount}`)
    lines.push('')

    lines.push('[TALLY → DMS]')
    lines.push(`  ✓ Matched records:         ${tallyMatched}`)
    lines.push(`  ✗ Unmatched records:       ${tallyUnmatched}`)
    const tallyMismatchCount = tallyToDmsResult.matched.filter(m => m.mismatches.length > 0).length
    lines.push(`  ⚠ Matched with mismatches: ${tallyMismatchCount}`)
    lines.push('')

    lines.push('TABLE: MISMATCHED VALUES (structured table is shown in UI + PDF)')
    lines.push('-'.repeat(80))
    lines.push('| (pipe-text table removed; broken formatting fixed) | | | |')
    lines.push('|---------------------------------------------------|---|---|---|')


    const sanitizeReportCell = (value: string) =>
      value.replace(/\|/g, '/').replace(/\r?\n/g, ' ').trim()

    // Kept only for backward-compatible text report; actual structured table is shown in UI + PDF.
    lines.push('| (structured mismatches table below) | | | |')
    lines.push('')


    lines.push('TABLE: SUMMARY')
    lines.push('-'.repeat(80))
    lines.push('| Metric                          | DMS → TALLY | TALLY → DMS |')
    lines.push('|---------------------------------|-------------|-------------|')
    lines.push(`| Total records                   | ${totalDmsRows.toString().padEnd(11)} | ${totalTallyRows.toString().padEnd(11)} |`)
    lines.push(`| Matched                         | ${dmsMatched.toString().padEnd(11)} | ${tallyMatched.toString().padEnd(11)} |`)
    lines.push(`| Unmatched                       | ${dmsUnmatched.toString().padEnd(11)} | ${tallyUnmatched.toString().padEnd(11)} |`)
    lines.push(`| Matched with mismatches         | ${dmsMismatchCount.toString().padEnd(11)} | ${tallyMismatchCount.toString().padEnd(11)} |`)
    lines.push('')

    // DMS → Tally section
    lines.push('='.repeat(80))
    lines.push('DETAIL: DMS → TALLY')
    lines.push('='.repeat(80))
    lines.push('')

    if (dmsUnmatched > 0) {
      lines.push('UNMATCHED RECORDS IN DMS (no matching Tally record found):')
      lines.push('-'.repeat(80))
      for (let idx = 0; idx < dmsToTallyResult.unmatched.length; idx++) {
        const item = dmsToTallyResult.unmatched[idx]
        const baseKeyVal = normalizeCell(item.row[baseKeyMapping.dmsCol])
        lines.push(``)
        lines.push(`Record #${idx + 1}: ${baseKeyMapping.dmsCol} = "${baseKeyVal}"`)
        for (const m of mappings) {
          if (m.dmsCol !== baseKeyMapping.dmsCol) {
            const val = normalizeCell(item.row[m.dmsCol])
            lines.push(`  ${m.dmsCol}: ${val || '(empty)'}`)
          }
        }
      }
      lines.push('')
      lines.push('')
    }

    if (dmsMismatchCount > 0) {
      lines.push('MATCHED BUT MISMATCHED RECORDS (DMS vs Tally):')
      lines.push('-'.repeat(80))
      let mismatchIdx = 0
      for (const item of dmsToTallyResult.matched) {
        if (item.mismatches.length === 0) continue

        const baseKeyVal = normalizeCell(item.fromRow[baseKeyMapping.dmsCol])
        lines.push(``)
        lines.push(`Mismatch #${++mismatchIdx}: ${baseKeyMapping.dmsCol} = "${baseKeyVal}"`)
        lines.push('')
        for (const mismatch of item.mismatches) {
          lines.push(`  Field: ${mismatch.field}`)
          lines.push(`    DMS:   "${mismatch.fromValue}"`)
          lines.push(`    Tally: "${mismatch.toValue}"`)
          if (mismatch.difference) {
            lines.push(`    Difference: ${mismatch.difference}`)
          }
          lines.push('')
        }
      }
      lines.push('')
    }

    // Tally → DMS section
    lines.push('='.repeat(80))
    lines.push('DETAIL: TALLY → DMS')
    lines.push('='.repeat(80))
    lines.push('')

    if (tallyUnmatched > 0) {
      lines.push('UNMATCHED RECORDS IN TALLY (no matching DMS record found):')
      lines.push('-'.repeat(80))
      for (let idx = 0; idx < tallyToDmsResult.unmatched.length; idx++) {
        const item = tallyToDmsResult.unmatched[idx]
        const baseKeyVal = normalizeCell(item.row[baseKeyMapping.tallyCol])
        lines.push(``)
        lines.push(`Record #${idx + 1}: ${baseKeyMapping.tallyCol} = "${baseKeyVal}"`)
        for (const m of mappings) {
          if (m.tallyCol !== baseKeyMapping.tallyCol) {
            const val = normalizeCell(item.row[m.tallyCol])
            lines.push(`  ${m.tallyCol}: ${val || '(empty)'}`)
          }
        }
      }
      lines.push('')
      lines.push('')
    }

    if (tallyMismatchCount > 0) {
      lines.push('MATCHED BUT MISMATCHED RECORDS (Tally vs DMS):')
      lines.push('-'.repeat(80))
      let mismatchIdx = 0
      for (const item of tallyToDmsResult.matched) {
        if (item.mismatches.length === 0) continue

        const baseKeyVal = normalizeCell(item.fromRow[baseKeyMapping.tallyCol])
        lines.push(``)
        lines.push(`Mismatch #${++mismatchIdx}: ${baseKeyMapping.tallyCol} = "${baseKeyVal}"`)
        lines.push('')
        for (const mismatch of item.mismatches) {
          lines.push(`  Field: ${mismatch.field}`)
          lines.push(`    Tally: "${mismatch.fromValue}"`)
          lines.push(`    DMS:   "${mismatch.toValue}"`)
          if (mismatch.difference) {
            lines.push(`    Difference: ${mismatch.difference}`)
          }
          lines.push('')
        }
      }
      lines.push('')
    }

    lines.push('='.repeat(80))
    lines.push('END OF REPORT')
    lines.push('='.repeat(80))

    return lines.join('\n')
  }

  async function handleRun() {
    if (!dms || !tally) return
    if (mappings.length === 0) return

    const baseKeyMapping = mappings[0]
    if (baseKeyMapping.dmsCol === ZERO_COMPARE || baseKeyMapping.tallyCol === ZERO_COMPARE) {
      alert('The first mapping must be a real key field on both sides. Use zero-compare only for non-key columns.')
      return
    }

    setBusy(true)
    try {
      // Use the first mapping (Bill Num -> Voucher ID) as the base key for matching
      const dmsToTally = compareDirection(dms, tally, baseKeyMapping, mappings, false)
      const tallyToDms = compareDirection(tally, dms, baseKeyMapping, mappings, true)
      
      setReportText(buildReportText(dmsToTally, tallyToDms, baseKeyMapping))

      const buildMismatchTable = (res: ComparisonResult, directionTitle: string) => {
        const rows: Array<{ billNo: string; valueName: string; dmsValue: string; tallyValue: string }> = []
        for (const item of res.matched) {
          const billNo = normalizeCell(item.fromRow[baseKeyMapping.dmsCol]) || normalizeCell(item.toRow[baseKeyMapping.tallyCol])
          for (const m of item.mismatches) {
            const parts = m.field.split('↔')
            const valueName = parts[0]?.trim() || m.field
            // In this UI table we follow the request format: DMS Value + Tally Value
            // dmsValue is always the “fromValue” and tallyValue is always the “toValue” from comparison direction.
            rows.push({
              billNo,
              valueName,
              dmsValue: m.fromValue,
              tallyValue: m.toValue
            })
          }
        }
        return { title: directionTitle, rows }
      }

      setMismatchTables([
        buildMismatchTable(dmsToTally, 'DMS → TALLY MISMATCHED VALUES'),
        buildMismatchTable(tallyToDms, 'TALLY → DMS MISMATCHED VALUES')
      ])
    } finally {
      setBusy(false)
    }
  }

  function getMismatchTables() {
    // Extract structured mismatch data for table display
    if (!dms || !tally || mappings.length === 0) return []
    
    const baseKeyMapping = mappings[0]
    const dmsToTallyResult = compareDirection(dms, tally, baseKeyMapping, mappings, false)
    
    const mismatchRows = dmsToTallyResult.matched
      .flatMap((item) => {
        const billNo = normalizeCell(item.fromRow[baseKeyMapping.dmsCol]) || normalizeCell(item.toRow[baseKeyMapping.tallyCol])
        return item.mismatches.map((mismatch) => ({
          billNo,
          valueName: mismatch.field,
          dmsValue: mismatch.fromValue,
          tallyValue: mismatch.toValue
        }))
      })
    
    return mismatchRows.length > 0 ? [{ title: 'Mismatched Values', rows: mismatchRows }] : []
  }

  function downloadReportPdf() {
    exportReportPdf({
      reportText: reportText || '',
      mismatchTables: getMismatchTables(),
      filename: `dmstally-report-${new Date().toISOString().slice(0, 10)}.pdf`
    }).catch(err => {
      // eslint-disable-next-line no-console
      console.error('PDF export failed:', err)
      alert('Failed to generate PDF. Check console for details.')
    })
  }

  useEffect(() => {
    if (dms && tally && mappings.length === 0) {
      ensureDefaultMappings()
    }
  }, [dms, tally, mappings.length])

  // NOTE: downloadReportPdf currently uses the older inline layout.



  return (
    <div className="container">
      <div className="card">
        <h1 style={{ marginTop: 0 }}>DMS/Tally Comparator</h1>
        <p style={{ color: '#9ca3af', marginBottom: 18 }}>
          Privacy-first: files are processed in your browser only. Nothing is stored or uploaded.
        </p>

        <div className="grid grid-2">
          <div className="card" style={{ background: '#0f1730' }}>
            <h3 style={{ marginTop: 0 }}>1) Upload Excel files</h3>
            <div className="grid">
              <div>
                <label>DMS Excel</label>
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  onChange={(e) => setDmsFile(e.target.files?.[0] ?? null)}
                />
                <div style={{ height: 8 }} />
                <button className="btn-secondary" disabled={!dmsFile || busy} onClick={() => handleLoad('DMS')}>
                  Load DMS
                </button>
              </div>
              <div>
                <label>Tally Excel</label>
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  onChange={(e) => setTallyFile(e.target.files?.[0] ?? null)}
                />
                <div style={{ height: 8 }} />
                <button className="btn-secondary" disabled={!tallyFile || busy} onClick={() => handleLoad('Tally')}>
                  Load Tally
                </button>
              </div>
            </div>
          </div>

          <div className="card" style={{ background: '#0f1730' }}>
            <h3 style={{ marginTop: 0 }}>2) Column cleanup preview</h3>
            <div className="row">
              <span className="badge">DMS columns kept: {dmsCols.length}</span>
              <span className="badge">Tally columns kept: {tallyCols.length}</span>
            </div>
            <div style={{ height: 10 }} />
            <div className="grid grid-2">
              <div>
                <label>Kept DMS columns</label>
                <div style={{ maxHeight: 220, overflow: 'auto', paddingRight: 6 }}>
                  {dmsCols.map((c) => (
                    <div key={c} style={{ fontSize: 13, padding: '3px 0' }}>{c}</div>
                  ))}
                  {dmsCols.length === 0 && <small>No file loaded.</small>}
                </div>
              </div>
              <div>
                <label>Kept Tally columns</label>
                <div style={{ maxHeight: 220, overflow: 'auto', paddingRight: 6 }}>
                  {tallyCols.map((c) => (
                    <div key={c} style={{ fontSize: 13, padding: '3px 0' }}>{c}</div>
                  ))}
                  {tallyCols.length === 0 && <small>No file loaded.</small>}
                </div>
              </div>
            </div>
            <div style={{ height: 14 }} />
            <div className="grid grid-2">
              <div>
                <label>Removed DMS columns</label>
                <div style={{ maxHeight: 220, overflow: 'auto', paddingRight: 6 }}>
                  {dmsRemovedCols.map((c) => (
                    <div key={c} style={{ fontSize: 13, padding: '3px 0' }}>{c}</div>
                  ))}
                  {dmsRemovedCols.length === 0 && <small>None removed.</small>}
                </div>
              </div>
              <div>
                <label>Removed Tally columns</label>
                <div style={{ maxHeight: 220, overflow: 'auto', paddingRight: 6 }}>
                  {tallyRemovedCols.map((c) => (
                    <div key={c} style={{ fontSize: 13, padding: '3px 0' }}>{c}</div>
                  ))}
                  {tallyRemovedCols.length === 0 && <small>None removed.</small>}
                </div>
              </div>
            </div>
            <div style={{ height: 14 }} />
            <div className="row">
              <button
                className="btn-secondary"
                disabled={!dms || !tally || busy}
                onClick={ensureDefaultMappings}
              >
                Suggest mappings
              </button>
              <small>
                Columns with no entries or all 0s are removed before mapping.
              </small>
            </div>
          </div>
        </div>

        <div style={{ height: 16 }} />

        <div className="card" style={{ background: '#0f1730' }}>
          <h3 style={{ marginTop: 0 }}>3) Map columns and compare</h3>

          <div className="grid" style={{ marginBottom: 12 }}>
            {mappings.length === 0 && (
              <small>
                Click <b>Suggest mappings</b> or add mappings below.
              </small>
            )}

            {mappings.map((m, idx) => (
              <div key={`${m.dmsCol}-${m.tallyCol}-${idx}`} className="row">
                <div style={{ flex: 1, minWidth: 320 }}>
                  <label>DMS column</label>
                  <select
                    value={m.dmsCol}
                    onChange={(e) => {
                      const next = [...mappings]
                      next[idx] = { ...next[idx], dmsCol: e.target.value }
                      setMappings(next)
                    }}
                  >
                    <option value={ZERO_COMPARE}>No DMS match (compare Tally column to 0)</option>
                    {dmsCols.map((c) => (
                      <option value={c} key={c} disabled={usedDmsCols.has(c) && c !== m.dmsCol}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>
                <div style={{ flex: 1, minWidth: 320 }}>
                  <label>Tally column</label>
                  <select
                    value={m.tallyCol}
                    onChange={(e) => {
                      const next = [...mappings]
                      next[idx] = { ...next[idx], tallyCol: e.target.value }
                      setMappings(next)
                    }}
                  >
                    <option value={ZERO_COMPARE}>No Tally match (compare DMS column to 0)</option>
                    {tallyCols.map((c) => (
                      <option value={c} key={c} disabled={usedTallyCols.has(c) && c !== m.tallyCol}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>
                <button
                  className="btn-secondary"
                  disabled={busy}
                  onClick={() => setMappings((prev) => prev.filter((_, i) => i !== idx))}
                >
                  Remove
                </button>
              </div>
            ))}

            <div className="row">
              <button
                className="btn-secondary"
                disabled={!dms || !tally || busy}
                onClick={() => {
                  if (!dms || !tally) return
                  const dmsCol = dms.columns[0]
                  const tallyCol = tally.columns[0]
                  setMappings((prev) => {
                    const next = [...prev, { dmsCol, tallyCol }]
                    return next
                  })
                }}
              >
                Add mapping
              </button>
              {!allMappedUnique && <small style={{ color: '#fca5a5' }}>Duplicate mappings detected.</small>}
            </div>

            <div className="row" style={{ marginTop: 8 }}>
              <button disabled={!dms || !tally || mappings.length === 0 || busy} onClick={handleRun}>
                {busy ? 'Comparing...' : 'Generate report'}
              </button>
              <button className="btn-secondary" disabled={!reportText} onClick={downloadReportPdf}>
                Download PDF
              </button>
              <button className="btn-secondary" disabled={!reportText} onClick={() => {
                if (dms && tally) downloadComparisonExcel(dms, tally, mappings)
              }} style={{ marginLeft: 10 }}>
                Download Excel
              </button>

              <small>
                Keying strategy: the first mapping is used to match records; the rest are compared for differences.
              </small>
            </div>

            {reportText && (
              <div style={{ marginTop: 16 }}>
                <label>Differences Table</label>
                <DifferencesTable mismatchRows={getMismatchTables()[0]?.rows || []} />
              </div>
            )}
          </div>

          <div>
            <label>Report output</label>
            <textarea value={reportText} readOnly placeholder="Load files, map columns, then generate report." />


          </div>
        </div>
      </div>
    </div>
  )
}

