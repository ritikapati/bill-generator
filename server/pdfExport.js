import { access } from 'node:fs/promises'
import { join } from 'node:path'
import PDFDocument from 'pdfkit'
import ExcelJS from 'exceljs'

const fontDirectory = process.env.WINDIR ? join(process.env.WINDIR, 'Fonts') : 'C:\\Windows\\Fonts'
const fontFiles = {
  regular: join(fontDirectory, 'GOTHIC.TTF'),
  bold: join(fontDirectory, 'GOTHICB.TTF'),
  italic: join(fontDirectory, 'GOTHICI.TTF'),
  boldItalic: join(fontDirectory, 'GOTHICBI.TTF'),
  kannada: join(fontDirectory, 'arialuni.ttf'),
  fallback: join(fontDirectory, 'arial.ttf'),
}

function colNumber(letters) {
  return [...letters].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0)
}

function cellText(cell) {
  let value = cell.value
  if (value && typeof value === 'object') {
    if (Object.hasOwn(value, 'result')) value = value.result
    else if (Array.isArray(value.richText)) value = value.richText.map((part) => part.text).join('')
    else if (value.text !== undefined) value = value.text
    else if (value.error) value = ''
  }
  if (value && typeof value === 'object' && value.error) value = ''
  if (value === null || value === undefined) return ''
  if (typeof value === 'number') {
    const format = cell.numFmt || ''
    if (format.includes('%')) return `${new Intl.NumberFormat('en-IN', { maximumFractionDigits: 2 }).format(value * 100)}%`
    const decimals = format.match(/\.([0#]+)/)?.[1].length || 0
    if (value === 0 && format.includes('"-"')) return '-'
    return new Intl.NumberFormat('en-IN', { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(value)
  }
  if (value instanceof Date) return new Intl.DateTimeFormat('en-IN').format(value)
  return String(value)
}

function colorOf(color) {
  if (color?.argb) return `#${color.argb.slice(-6)}`
  if (color?.rgb) return `#${color.rgb.slice(-6)}`
  if (color?.indexed !== undefined) return color.indexed === 9 ? '#ffffff' : '#000000'
  return '#000000'
}

function textFontName(text, font) {
  if (/[\u0c80-\u0cff]/u.test(text)) return 'Kannada'
  if (font?.bold && font?.italic) return 'CenturyGothicBoldItalic'
  if (font?.bold) return 'CenturyGothicBold'
  if (font?.italic) return 'CenturyGothicItalic'
  return 'CenturyGothic'
}

function edgeStyle(border) {
  if (!border?.style) return null
  const widths = { hair: 0.35, thin: 0.5, medium: 1.15, thick: 1.7, double: 1.3 }
  return { width: widths[border.style] || 0.5, color: colorOf(border.color), style: border.style }
}

function drawBorder(doc, side, x1, y1, x2, y2, border, scale) {
  const style = edgeStyle(border)
  if (!style) return
  doc.save().lineWidth(style.width * scale).strokeColor(style.color)
  if (style.style === 'dotted' || style.style === 'hair') doc.dash(1.2 * scale, 1.5 * scale)
  else if (style.style.toLowerCase().includes('dash')) doc.dash(3 * scale, 2 * scale)
  else doc.undash()
  doc.moveTo(x1, y1).lineTo(x2, y2).stroke().restore()
}

function drawCellText(doc, cell, text, x, y, width, height, scale) {
  if (!text) return
  const style = cell.style || {}
  const fontStyle = style.font || {}
  const padding = 2.5 * scale
  const boxWidth = Math.max(1, width - padding * 2)
  const boxHeight = Math.max(1, height - padding * 2)
  const fontName = textFontName(text, fontStyle)
  let size = (fontStyle.size || 11) * scale
  doc.font(fontName).fontSize(size)
  if (!style.alignment?.wrapText && text.includes('\n')) text = text.replace(/\r?\n/g, ' ')
  if (!style.alignment?.wrapText) {
    while (size > 5 * scale && doc.widthOfString(text) > boxWidth) {
      size -= 0.5 * scale
      doc.fontSize(size)
    }
  }
  const alignment = style.alignment || {}
  const options = {
    width: boxWidth,
    height: boxHeight,
    align: alignment.horizontal === 'center' ? 'center' : alignment.horizontal === 'right' ? 'right' : 'left',
    valign: alignment.vertical === 'middle' ? 'center' : alignment.vertical === 'bottom' ? 'bottom' : 'top',
    lineBreak: Boolean(alignment.wrapText),
    paragraphGap: 0,
    characterSpacing: 0,
  }
  doc.fillColor(colorOf(fontStyle.color)).text(text, x + padding, y + padding, options)
}

function drawWorksheet(doc, sheet, paper) {
  const pageWidth = paper === 'A4' ? 595.28 : 612
  const pageHeight = paper === 'A4' ? 841.89 : 792
  const margins = sheet.pageSetup.margins || { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 }
  const left = (margins.left ?? 0.7) * 72
  const right = (margins.right ?? 0.7) * 72
  const top = (margins.top ?? 0.75) * 72
  const bottom = (margins.bottom ?? 0.75) * 72
  let lastColumn = 1
  sheet.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
    if (cellText(cell).trim()) lastColumn = Math.max(lastColumn, cell.col)
  }))
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    for (let column = 1; column <= sheet.columnCount; column += 1) {
      if (sheet.getCell(row, column).border?.right?.style) lastColumn = Math.max(lastColumn, column)
    }
  }

  const widths = Array.from({ length: lastColumn }, (_, index) => {
    const column = sheet.getColumn(index + 1)
    return column.hidden ? 0 : ((column.width || 9) * 7 + 5) * 0.75
  })
  const heights = Array.from({ length: sheet.rowCount }, (_, index) => sheet.getRow(index + 1).hidden ? 0 : sheet.getRow(index + 1).height || 15)
  const totalWidth = widths.reduce((sum, value) => sum + value, 0)
  const totalHeight = heights.reduce((sum, value) => sum + value, 0)
  const availableWidth = pageWidth - left - right
  const availableHeight = pageHeight - top - bottom
  const configuredScale = sheet.pageSetup.scale ? sheet.pageSetup.scale / 100 : 1
  const scale = sheet.pageSetup.fitToPage
    ? Math.min(availableWidth / totalWidth, availableHeight / totalHeight)
    : Math.min(configuredScale, availableWidth / totalWidth, availableHeight / totalHeight)
  const renderedWidth = totalWidth * scale
  let xOrigin = left
  if (sheet.pageSetup.horizontalCentered) xOrigin = left + (availableWidth - renderedWidth) / 2
  const xPositions = [xOrigin]
  const yPositions = [top]
  widths.forEach((width) => xPositions.push(xPositions.at(-1) + width * scale))
  heights.forEach((height) => yPositions.push(yPositions.at(-1) + height * scale))

  const merges = sheet.model.merges.map((range) => {
    const [start, end] = range.split(':')
    const startMatch = start.match(/([A-Z]+)(\d+)/)
    const endMatch = end.match(/([A-Z]+)(\d+)/)
    return { startRow: Number(startMatch[2]), startColumn: colNumber(startMatch[1]), endRow: Number(endMatch[2]), endColumn: colNumber(endMatch[1]) }
  })
  const mergeStarts = new Map(merges.map((merge) => [`${merge.startRow}:${merge.startColumn}`, merge]))
  const mergeChildren = new Set()
  for (const merge of merges) {
    for (let row = merge.startRow; row <= merge.endRow; row += 1) {
      for (let column = merge.startColumn; column <= merge.endColumn; column += 1) {
        if (row !== merge.startRow || column !== merge.startColumn) mergeChildren.add(`${row}:${column}`)
      }
    }
  }

  for (let row = 1; row <= sheet.rowCount; row += 1) {
    if (!heights[row - 1]) continue
    for (let column = 1; column <= lastColumn; column += 1) {
      if (widths[column - 1] === 0 || mergeChildren.has(`${row}:${column}`)) continue
      const cell = sheet.getCell(row, column)
      const merge = mergeStarts.get(`${row}:${column}`)
      const endRow = merge?.endRow || row
      const endColumn = Math.min(merge?.endColumn || column, lastColumn)
      const x = xPositions[column - 1]
      const y = yPositions[row - 1]
      const cellWidth = xPositions[endColumn] - x
      const cellHeight = yPositions[endRow] - y
      const fill = cell.style?.fill
      if (fill?.type === 'pattern' && fill.pattern !== 'none') {
        doc.save().fillColor(colorOf(fill.fgColor)).rect(x, y, cellWidth, cellHeight).fill().restore()
      }
      drawCellText(doc, cell, cellText(cell), x, y, cellWidth, cellHeight, scale)

      for (let edgeColumn = column; edgeColumn <= endColumn; edgeColumn += 1) {
        const topCell = sheet.getCell(row, edgeColumn)
        const bottomCell = sheet.getCell(endRow, edgeColumn)
        drawBorder(doc, 'top', xPositions[edgeColumn - 1], y, xPositions[edgeColumn], y, topCell.border?.top, scale)
        drawBorder(doc, 'bottom', xPositions[edgeColumn - 1], yPositions[endRow], xPositions[edgeColumn], yPositions[endRow], bottomCell.border?.bottom, scale)
      }
      for (let edgeRow = row; edgeRow <= endRow; edgeRow += 1) {
        const leftCell = sheet.getCell(edgeRow, column)
        const rightCell = sheet.getCell(edgeRow, endColumn)
        drawBorder(doc, 'left', x, yPositions[edgeRow - 1], x, yPositions[edgeRow], leftCell.border?.left, scale)
        drawBorder(doc, 'right', xPositions[endColumn], yPositions[edgeRow - 1], xPositions[endColumn], yPositions[edgeRow], rightCell.border?.right, scale)
      }
    }
  }
}

export async function exportWorkbookPdf(workbookBuffer) {
  if (process.platform !== 'win32') throw new Error('PDF export currently requires Windows.')
  for (const path of [fontFiles.regular, fontFiles.bold, fontFiles.italic, fontFiles.boldItalic, fontFiles.fallback]) {
    try { await access(path) } catch { throw new Error(`PDF export font file is missing: ${path}`) }
  }
  try { await access(fontFiles.kannada) } catch {
    throw new Error(`PDF export needs a Kannada-capable font at ${fontFiles.kannada}.`)
  }

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(workbookBuffer)
  const sheets = workbook.worksheets
  if (!sheets.length) throw new Error('The workbook has no printable worksheets.')

  return new Promise((resolve, reject) => {
    const document = new PDFDocument({ autoFirstPage: false, bufferPages: true, compress: true })
    const chunks = []
    document.registerFont('CenturyGothic', fontFiles.regular)
    document.registerFont('CenturyGothicBold', fontFiles.bold)
    document.registerFont('CenturyGothicItalic', fontFiles.italic)
    document.registerFont('CenturyGothicBoldItalic', fontFiles.boldItalic)
    document.registerFont('Kannada', fontFiles.kannada)
    document.registerFont('ArialFallback', fontFiles.fallback)
    document.on('data', (chunk) => chunks.push(chunk))
    document.on('error', reject)
    document.on('end', () => resolve(Buffer.concat(chunks)))
    try {
      for (const sheet of sheets) {
        const paper = sheet.pageSetup.paperSize === 9 ? 'A4' : 'Letter'
        document.addPage({ size: paper, margin: 0, layout: 'portrait' })
        drawWorksheet(document, sheet, paper)
      }
      document.end()
    } catch (error) {
      reject(error)
    }
  })
}
