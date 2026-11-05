import Decimal from 'decimal.js'
import ExcelJS from 'exceljs'
import { fileURLToPath } from 'node:url'
import { validateRows } from './validation.js'

Decimal.set({ precision: 24, rounding: Decimal.ROUND_HALF_UP })

export function calculateTax(quantity, rate, gstRate, taxMode) {
  const qty = new Decimal(quantity)
  const unitRate = new Decimal(rate)
  const taxPercent = new Decimal(gstRate)
  const grossOrBase = qty.times(unitRate)
  const roundedGross = grossOrBase.toDecimalPlaces(2)
  const taxable = (taxMode === 'inclusive'
    ? roundedGross.dividedBy(new Decimal(1).plus(taxPercent.dividedBy(100)))
    : grossOrBase).toDecimalPlaces(2)
  const tax = (taxMode === 'inclusive'
    ? roundedGross.minus(taxable)
    : taxable.times(taxPercent).dividedBy(100)).toDecimalPlaces(2)
  const gross = taxMode === 'inclusive' ? roundedGross : taxable.plus(tax).toDecimalPlaces(2)
  const cgst = tax.dividedBy(2).toDecimalPlaces(2)

  return {
    taxable,
    cgst,
    sgst: tax.minus(cgst),
    tax,
    gross,
  }
}

function amountInWords(amount) {
  const belowTwenty = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen']
  const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']
  const underThousand = (value) => {
    if (value < 20) return belowTwenty[value]
    if (value < 100) return `${tens[Math.floor(value / 10)]}${value % 10 ? ` ${belowTwenty[value % 10]}` : ''}`
    return `${belowTwenty[Math.floor(value / 100)]} Hundred${value % 100 ? ` ${underThousand(value % 100)}` : ''}`
  }
  const chunk = (value, unit) => value ? `${underThousand(value)}${unit ? ` ${unit}` : ''}` : ''
  let remaining = Math.floor(Number(amount))
  const rupees = [
    chunk(Math.floor(remaining / 10000000), 'Crore'),
    chunk(Math.floor((remaining %= 10000000) / 100000), 'Lakh'),
    chunk(Math.floor((remaining %= 100000) / 1000), 'Thousand'),
    chunk(remaining % 1000, ''),
  ].filter(Boolean).join(' ')
  const paise = Math.round((Number(amount) - Math.floor(Number(amount))) * 100)
  return `${rupees || 'Zero'} Rupees${paise ? ` and ${underThousand(paise)} Paise` : ''} Only`
}

function safeSheetName(name, index) {
  const cleaned = name.replace(/[\\/?*:[\]]/g, ' ').trim().slice(0, 24) || 'Bill'
  return `${cleaned} ${index + 1}`.slice(0, 31)
}

const templates = {
  credit: fileURLToPath(new URL('../Credit bill Format.xlsm', import.meta.url)),
  challan: fileURLToPath(new URL('../DC Format.xlsm', import.meta.url)),
}

const templateSheets = {
  credit: { General: 'Sheet1', SCP: 'Sheet2', TSP: 'Sheet3' },
  challan: { General: '341_Koppal_GEN', SCP: '342_Koppal_SCP', TSP: '343_Koppal_TSP' },
}

function cloneWorksheet(workbook, source, name) {
  const clone = workbook.addWorksheet(name)
  clone.properties = structuredClone(source.properties)
  clone.pageSetup = structuredClone(source.pageSetup)
  clone.headerFooter = structuredClone(source.headerFooter)
  clone.views = structuredClone(source.views)
  clone.state = source.state
  clone.autoFilter = structuredClone(source.autoFilter)
  clone.columns = source.columns.map((column) => ({
    key: column.key,
    width: column.width,
    hidden: column.hidden,
    outlineLevel: column.outlineLevel,
    collapsed: column.collapsed,
    header: structuredClone(column.header),
    style: structuredClone(column.style),
  }))

  for (let rowNumber = 1; rowNumber <= source.rowCount; rowNumber += 1) {
    const sourceRow = source.getRow(rowNumber)
    const targetRow = clone.getRow(rowNumber)
    targetRow.height = sourceRow.height
    targetRow.hidden = sourceRow.hidden
    targetRow.outlineLevel = sourceRow.outlineLevel
    targetRow.style = structuredClone(sourceRow.style)
    for (let columnNumber = 1; columnNumber <= source.columnCount; columnNumber += 1) {
      const sourceCell = sourceRow.getCell(columnNumber)
      const targetCell = targetRow.getCell(columnNumber)
      targetCell.value = structuredClone(sourceCell.value)
      targetCell.style = structuredClone(sourceCell.style)
      if (sourceCell.dataValidation) targetCell.dataValidation = structuredClone(sourceCell.dataValidation)
    }
  }

  source.model.merges.forEach((range) => clone.mergeCells(range))
  return clone
}

function groupBills(rows, documentType) {
  const groups = new Map()
  rows.forEach((row) => {
    const keyParts = [row.ada.toLocaleLowerCase('en-IN'), row.scheme]
    if (documentType === 'challan') keyParts.splice(1, 0, row.rsk.toLocaleLowerCase('en-IN'))
    const key = keyParts.join('|')
    if (!groups.has(key)) groups.set(key, { ada: row.ada, rsk: row.rsk, scheme: row.scheme, rows: [] })
    groups.get(key).rows.push(row)
  })

  for (const group of groups.values()) {
    const crops = new Set(group.rows.map((row) => row.crop.toLocaleLowerCase('en-IN')))
    if (crops.size > 1) {
      throw new Error(`Multiple products map to ${group.ada} / ${group.scheme}. Resolve the product mapping before generating bills.`)
    }
    group.quantity = group.rows.reduce((total, row) => total.plus(row.quantity), new Decimal(0))
    group.crop = group.rows[0].crop
  }
  return [...groups.values()]
}

export function countBillDocuments(inputRows, documentType) {
  const rows = validateRows(inputRows, { documentType })
  return groupBills(rows, documentType).length
}

export async function createBillWorkbook({ rows: inputRows, documentType, rate, gstRate, taxMode, firstBillNumber = 1 }) {
  const rows = validateRows(inputRows, { documentType })
  if (!Number.isSafeInteger(Number(firstBillNumber)) || Number(firstBillNumber) < 1) throw new Error('Starting bill number must be a positive whole number.')
  const unitRate = Number(rate)
  const taxPercent = Number(gstRate)
  if (!Number.isFinite(unitRate) || unitRate < 0) throw new Error('Enter a valid non-negative unit rate.')
  if (!Number.isFinite(taxPercent) || taxPercent < 0 || taxPercent > 100) throw new Error('GST rate must be between 0 and 100.')
  if (!['inclusive', 'exclusive'].includes(taxMode)) throw new Error('Choose an approved GST rate mode.')
  if (!['credit', 'challan'].includes(documentType)) throw new Error('Choose Credit Bill or Delivery Challan.')

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(templates[documentType])
  workbook.creator = 'Shrishti Enterprises Chart-to-Bill'
  workbook.subject = 'Reviewed chart data sample bill'
  workbook.created = new Date()
  workbook.calcProperties.fullCalcOnLoad = true
  const groups = groupBills(rows, documentType)
  const seedByScheme = new Map(Object.entries(templateSheets[documentType]).map(([scheme, sheetName]) => [scheme, workbook.getWorksheet(sheetName)]))
  const usedSeeds = new Set()
  const generatedSheets = []

  groups.forEach((group, index) => {
    const seed = seedByScheme.get(group.scheme)
    let worksheet
    if (!usedSeeds.has(seed.id)) {
      worksheet = seed
      usedSeeds.add(seed.id)
    } else {
      worksheet = cloneWorksheet(workbook, seed, safeSheetName(`${group.ada} ${group.scheme}`, index))
    }
    worksheet.name = safeSheetName(`${group.ada} ${group.scheme}`, index)
    generatedSheets.push(worksheet)

    const numberWidth = documentType === 'credit' ? 2 : 3
    const billNumber = String(firstBillNumber + index).padStart(numberWidth, '0')
    const amounts = calculateTax(group.quantity.toNumber(), unitRate, taxPercent, taxMode)
    const displayedRate = taxMode === 'exclusive'
      ? new Decimal(unitRate).times(new Decimal(1).plus(new Decimal(taxPercent).dividedBy(100))).toDecimalPlaces(2)
      : new Decimal(unitRate)
    worksheet.getCell('A9').value = `ASST DIRECTOR AGRICULTURE OF ${group.ada.toLocaleUpperCase('en-IN')}`
    worksheet.getCell('G8').value = `  BILL NO :-  ${billNumber}`
    worksheet.getCell('H8').value = 'Date :-     '
    worksheet.getCell('H10').value = group.scheme === 'General' ? 'GENERAL' : group.scheme
    worksheet.getCell('H11').value = group.crop
    if (documentType === 'challan') worksheet.getCell('A10').value = `[RSK ${group.rsk}]`

    worksheet.getCell('G13').value = group.quantity.toNumber()
    worksheet.getCell('H13').value = displayedRate.toNumber()
    worksheet.getCell('J13').value = taxPercent / 100
    const amountFormula = documentType === 'credit' ? '+G13*H13' : 'ROUND(G13*H13,2)'
    worksheet.getCell('I13').value = { formula: amountFormula, result: amounts.gross.toNumber() }
    worksheet.getCell('K13').value = { formula: 'SUM(H13*G13)', result: amounts.gross.toNumber() }
    worksheet.getCell('K27').value = { formula: '+SUM(K13:K26)', result: amounts.gross.toNumber() }
    worksheet.getCell('B28').value = amounts.taxable.toNumber()
    worksheet.getCell('C28').value = amounts.tax.toNumber()
    worksheet.getCell('D28').value = amounts.sgst.toNumber()
    worksheet.getCell('E28').value = amounts.cgst.toNumber()

    const wordsCell = documentType === 'credit' ? 'G31' : 'F31'
    worksheet.getCell(wordsCell).value = amountInWords(amounts.gross.toNumber()).toLocaleLowerCase('en-IN')
  })

  workbook.worksheets.slice().forEach((worksheet) => {
    if (!generatedSheets.includes(worksheet)) workbook.removeWorksheet(worksheet.id)
  })

  return workbook.xlsx.writeBuffer()
}
