import type { jsPDF as JsPDFType } from 'jspdf'

type JsPDFCtor = new (options?: any) => JsPDFType

function sanitizeText(text: string): string {
  return text
    .replace(/↔/g, '<->')
    .replace(/✓/g, '[OK]')
    .replace(/✗/g, '[X]')
    .replace(/⚠/g, '[!]')
    .replace(/[']/g, "'")
    .replace(/["]/g, '"')
}

type MismatchRow = { billNo: string; valueName: string; dmsValue: string; tallyValue: string }

type MismatchTable = {
  title: string
  rows: MismatchRow[]
}

export function exportReportPdf(params: {
  reportText: string
  filename: string
  mismatchTables?: MismatchTable[]
}) {
  const { reportText, filename, mismatchTables = [] } = params

  return import('jspdf').then(({ jsPDF }: any) => {
    const doc = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4' })

    const pageWidth = doc.internal.pageSize.getWidth()
    const pageHeight = doc.internal.pageSize.getHeight()

    const marginX = 15
    const marginTop = 25
    const marginBottom = 15
    const lineHeight = 5
    const maxWidth = pageWidth - marginX * 2

    let y = marginTop

    // Header
    doc.setFillColor(59, 130, 246)
    doc.rect(0, 0, pageWidth, 20, 'F')

    doc.setTextColor(255, 255, 255)
    doc.setFont('helvetica', 'bold')
    doc.setFontSize(16)
    doc.text('DMS / Tally Comparison Report', marginX, 12)

    doc.setFontSize(9)
    doc.setFont('helvetica', 'normal')
    doc.setTextColor(0, 0, 0)

    y = 28

    // Structured mismatch tables FIRST - grouped by value name
    const tableColWidths = [35, 55, 55]  // Bill No, DMS Value, Tally Value - increased widths
    const cellHeight = 6
    const headers = ['Bill No', 'DMS Value', 'Tally Value']

    for (const t of mismatchTables) {
      if (!t.rows.length) continue

      // Group rows by value name
      const rowsByValueName = new Map<string, MismatchRow[]>()
      for (const row of t.rows) {
        if (!rowsByValueName.has(row.valueName)) {
          rowsByValueName.set(row.valueName, [])
        }
        rowsByValueName.get(row.valueName)!.push(row)
      }

      // Create a table for each value name
      for (const [valueName, rows] of rowsByValueName) {
        if (y > pageHeight - marginBottom - 30) {
          doc.addPage()
          y = marginTop
        }

        doc.setFont('helvetica', 'bold')
        doc.setFontSize(11)
        doc.setTextColor(30, 60, 130)
        doc.text(sanitizeText(valueName), marginX, y)
        y += 6
        doc.setTextColor(0, 0, 0)
        doc.setFont('helvetica', 'normal')
        doc.setFontSize(9)

        let x = marginX
        for (let i = 0; i < headers.length; i++) {
          doc.setDrawColor(200, 200, 200)
          doc.setLineWidth(0.2)
          doc.rect(x, y, tableColWidths[i], cellHeight, 'S')
          doc.setFont('helvetica', 'bold')
          doc.text(headers[i], x + 1, y + 4)
          doc.setFont('helvetica', 'normal')
          x += tableColWidths[i]
        }
        y += cellHeight

        for (const r of rows) {
          const rowTexts = [r.billNo, r.dmsValue, r.tallyValue].map((s) => sanitizeText(s || ''))

          // Pre-calculate wrapped lines for all columns to determine row height
          const wrappedLines: string[][] = []
          let maxLines = 1
          for (let i = 0; i < rowTexts.length; i++) {
            const wrapped = doc.splitTextToSize(rowTexts[i], tableColWidths[i] - 2)
            wrappedLines.push(wrapped)
            maxLines = Math.max(maxLines, wrapped.length)
          }

          const rowHeight = cellHeight * maxLines + 2

          if (y > pageHeight - marginBottom - rowHeight) {
            doc.addPage()
            y = marginTop
          }

          x = marginX

          for (let i = 0; i < rowTexts.length; i++) {
            doc.setDrawColor(200, 200, 200)
            doc.setLineWidth(0.2)
            doc.rect(x, y, tableColWidths[i], rowHeight, 'S')

            // Render all wrapped lines
            const wrapped = wrappedLines[i]
            for (let lineIdx = 0; lineIdx < wrapped.length; lineIdx++) {
              doc.text(wrapped[lineIdx], x + 1, y + 4 + lineIdx * cellHeight)
            }

            x += tableColWidths[i]
          }

          y += rowHeight
        }

        y += 6  // Space between tables
      }
    }

    // Add page break before detailed report if we added tables
    if (mismatchTables.length > 0 && mismatchTables.some(t => t.rows.length > 0)) {
      doc.addPage()
      y = marginTop

      doc.setFont('helvetica', 'bold')
      doc.setFontSize(12)
      doc.setTextColor(59, 130, 246)
      doc.text('DETAILED REPORT', marginX, y)
      y += lineHeight + 1

      doc.setDrawColor(59, 130, 246)
      doc.setLineWidth(0.5)
      doc.line(marginX, y, pageWidth - marginX, y)
      y += lineHeight + 1

      doc.setTextColor(0, 0, 0)
      doc.setFont('helvetica', 'normal')
      doc.setFontSize(9)
    }

    // Report text SECOND
    const lines = reportText.split('\n')
    for (let i = 0; i < lines.length; i++) {
      const rawLine = lines[i]
      const line = sanitizeText(rawLine)

      if (y > pageHeight - marginBottom) {
        doc.addPage()
        y = marginTop
        doc.setFontSize(8)
        doc.setTextColor(180, 180, 180)
        doc.text(`Page ${doc.internal.pages.length - 1}`, pageWidth - marginX - 10, pageHeight - marginBottom + 5)
        doc.setTextColor(0, 0, 0)
        doc.setFontSize(9)
      }

      if (line.startsWith('===')) {
        if (y > marginTop + 5) y += 2
        doc.setFont('helvetica', 'bold')
        doc.setFontSize(12)
        doc.setTextColor(59, 130, 246)
        const headerText = line.replace(/=/g, '').trim()
        doc.text(headerText, marginX, y)
        y += lineHeight + 1

        doc.setDrawColor(59, 130, 246)
        doc.setLineWidth(0.5)
        doc.line(marginX, y, pageWidth - marginX, y)
        y += lineHeight

        doc.setTextColor(0, 0, 0)
        doc.setFont('helvetica', 'normal')
        doc.setFontSize(9)
      } else if (line.startsWith('---')) {
        y += 1
        doc.setDrawColor(200, 200, 200)
        doc.setLineWidth(0.3)
        doc.line(marginX, y, pageWidth - marginX, y)
        y += lineHeight + 1
      } else if (line.match(/^(UNMATCHED|MATCHED|SUMMARY|DETAIL|BASE|OTHER)/i)) {
        doc.setFont('helvetica', 'bold')
        doc.setFontSize(10)
        doc.setTextColor(80, 120, 180)
        doc.text(line, marginX, y)
        y += lineHeight
        doc.setTextColor(0, 0, 0)
        doc.setFont('helvetica', 'normal')
        doc.setFontSize(9)
      } else if (line.match(/^\[(OK|X|!)/)) {
        doc.setFont('helvetica', 'normal')
        doc.setFontSize(9)
        doc.setTextColor(60, 120, 60)
        doc.text(line, marginX + 5, y)
        y += lineHeight
        doc.setTextColor(0, 0, 0)
      } else if (line.match(/^(Total|Record #|Mismatch #|Field:)/)) {
        doc.setFont('helvetica', 'bold')
        doc.setFontSize(9)
        doc.setTextColor(30, 60, 130)
        doc.text(line, marginX, y)
        y += lineHeight
        doc.setTextColor(0, 0, 0)
        doc.setFont('helvetica', 'normal')
      } else if (line.trim() === '') {
        y += 1
      } else {
        doc.setFont('helvetica', 'normal')
        doc.setFontSize(9)
        doc.setTextColor(0, 0, 0)

        const indent = line.startsWith('  ') ? 5 : 0
        const textToWrap = line.trimLeft()

        const wrapped = doc.splitTextToSize(textToWrap, maxWidth - indent)
        for (const wrappedLine of wrapped) {
          if (y > pageHeight - marginBottom) {
            doc.addPage()
            y = marginTop
            doc.setFontSize(8)
            doc.setTextColor(180, 180, 180)
            doc.text(`Page ${doc.internal.pages.length - 1}`, pageWidth - marginX - 10, pageHeight - marginBottom + 5)
            doc.setTextColor(0, 0, 0)
            doc.setFontSize(9)
          }
          doc.text(wrappedLine, marginX + indent, y)
          y += lineHeight
        }
      }
    }

    // Footer
    doc.setFontSize(8)
    doc.setTextColor(180, 180, 180)
    doc.text('Privacy-first: Data processed locally in your browser only', marginX, pageHeight - 8)

    doc.save(filename)
  })
}

