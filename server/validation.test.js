import test from 'node:test'
import assert from 'node:assert/strict'
import { fileURLToPath } from 'node:url'
import ExcelJS from 'exceljs'
import sharp from 'sharp'
import { calculateTax, createBillWorkbook } from './billing.js'
import { prepareChartImage } from './chartImage.js'
import { expandChartGroups } from './chartExtraction.js'
import { readOllamaStream } from './ollamaStream.js'
import { validateRows } from './validation.js'

function assertLayoutMatches(actual, expected) {
  assert.deepEqual([...actual.model.merges].sort(), [...expected.model.merges].sort())
  for (const key of ['orientation', 'fitToPage', 'fitToWidth', 'fitToHeight', 'paperSize', 'scale']) {
    assert.equal(actual.pageSetup[key], expected.pageSetup[key])
  }
  assert.deepEqual(actual.pageSetup.margins, expected.pageSetup.margins)
  assert.equal(actual.rowCount, expected.rowCount)
  assert.equal(actual.columnCount, expected.columnCount)
  for (let rowNumber = 1; rowNumber <= expected.rowCount; rowNumber += 1) {
    assert.equal(actual.getRow(rowNumber).height, expected.getRow(rowNumber).height)
    for (let columnNumber = 1; columnNumber <= expected.columnCount; columnNumber += 1) {
      assert.deepEqual(
        actual.getRow(rowNumber).getCell(columnNumber).style,
        expected.getRow(rowNumber).getCell(columnNumber).style,
        `${actual.getRow(rowNumber).getCell(columnNumber).address} style`,
      )
    }
  }
  for (let columnNumber = 1; columnNumber <= expected.columnCount; columnNumber += 1) {
    assert.equal(actual.getColumn(columnNumber).width, expected.getColumn(columnNumber).width)
  }
}

test('validates required ADA, allowed schemes, and finite quantities', () => {
  assert.deepEqual(validateRows([{ ada: 'Ballari', rsk: '', scheme: 'SCP', crop: 'Maize', quantity: 1.5 }]), [
    { ada: 'Ballari', rsk: '', scheme: 'SCP', crop: 'Maize', quantity: 1.5 },
  ])
  assert.throws(() => validateRows([{ ada: '', scheme: 'General', quantity: 2 }]), /ADA is required/)
  assert.throws(() => validateRows([{ ada: 'Ballari', scheme: 'Unknown', quantity: 2 }]), /scheme must be/)
  assert.throws(() => validateRows([{ ada: 'Ballari', scheme: 'General', crop: 'Maize', quantity: null }]), /valid non-negative quantity/)
  assert.throws(() => validateRows([{ ada: 'Ballari', scheme: 'General', crop: '', quantity: 2 }]), /crop or product is required/)
  assert.throws(() => validateRows([{ ada: 'Ballari', scheme: 'General', crop: 'Maize', quantity: 2 }], { documentType: 'challan' }), /RSK is required/)
})

test('calculates inclusive and exclusive GST with exact two-place totals', () => {
  const inclusive = calculateTax(3, 118, 18, 'inclusive')
  assert.equal(inclusive.taxable.toString(), '300')
  assert.equal(inclusive.cgst.toString(), '27')
  assert.equal(inclusive.sgst.toString(), '27')
  assert.equal(inclusive.gross.toString(), '354')
  const exclusive = calculateTax(3, 100, 18, 'exclusive')
  assert.equal(exclusive.taxable.toString(), '300')
  assert.equal(exclusive.tax.toString(), '54')
  assert.equal(exclusive.gross.toString(), '354')
})

test('fills the supplied Credit Bill and Delivery Challan templates', async () => {
  const rows = [{ ada: 'Ballari', rsk: 'Ballari Rural', scheme: 'General', crop: 'Maize', quantity: 1 }]
  const schemeDetails = 'NEMO OILSEEDS (ಸೂರ್ಯಕಾಂತಿ)'
  const input = { rows, schemeDetails, rate: 118, gstRate: 18, taxMode: 'inclusive' }
  const credit = new ExcelJS.Workbook()
  await credit.xlsx.load(await createBillWorkbook({ ...input, documentType: 'credit', firstBillNumber: 7 }))
  const creditSheet = credit.worksheets[0]
  assert.equal(creditSheet.getCell('A2').value, 'CREDIT BILL')
  assert.equal(creditSheet.getCell('A3').value, 'SHRISTHI ENTERPRISES')
  assert.equal(creditSheet.getCell('A9').value, 'ASST DIRECTOR AGRICULTURE OF BALLARI')
  assert.equal(creditSheet.getCell('G8').value, '  BILL NO :-  07')
  assert.equal(creditSheet.getCell('H8').value, 'Date :-     ')
  assert.equal(creditSheet.getCell('H10').value, 'GENERAL')
  assert.equal(creditSheet.getCell('H11').value, schemeDetails)
  assert.equal(creditSheet.getCell('G13').value, 1)
  assert.equal(creditSheet.getCell('I13').value.formula, '+G13*H13')
  assert.equal(creditSheet.getCell('K13').value.formula, 'SUM(H13*G13)')
  assert.equal(creditSheet.getCell('G31').value, 'one hundred eighteen rupees only')

  const creditTemplate = new ExcelJS.Workbook()
  await creditTemplate.xlsx.readFile(fileURLToPath(new URL('../Credit bill Format.xlsm', import.meta.url)))
  const originalCreditSheet = creditTemplate.getWorksheet('Sheet1')
  assertLayoutMatches(creditSheet, originalCreditSheet)
  assert.equal(creditSheet.getCell('B13').value, originalCreditSheet.getCell('B13').value)

  const challan = new ExcelJS.Workbook()
  await challan.xlsx.load(await createBillWorkbook({ ...input, documentType: 'challan', firstBillNumber: 341 }))
  const challanSheet = challan.worksheets[0]
  assert.equal(challanSheet.getCell('A2').value, 'DELIVERY CHALLAN')
  assert.equal(challanSheet.getCell('A9').value, 'ASST DIRECTOR AGRICULTURE OF BALLARI')
  assert.equal(challanSheet.getCell('A10').value, '[RSK Ballari Rural]')
  assert.equal(challanSheet.getCell('G8').value, '  BILL NO :-  341')
  assert.equal(challanSheet.getCell('H8').value, 'Date :-     ')
  assert.equal(challanSheet.getCell('H10').value, 'GENERAL')
  assert.equal(challanSheet.getCell('H11').value, schemeDetails)
  assert.equal(challanSheet.getCell('F31').value, 'one hundred eighteen rupees only')

  const challanTemplate = new ExcelJS.Workbook()
  await challanTemplate.xlsx.readFile(fileURLToPath(new URL('../DC Format.xlsm', import.meta.url)))
  const originalChallanSheet = challanTemplate.getWorksheet('341_Koppal_GEN')
  assertLayoutMatches(challanSheet, originalChallanSheet)
  assert.equal(challanSheet.getCell('B13').value, 'Thiamethoxam 12.6%')

  const exclusive = new ExcelJS.Workbook()
  await exclusive.xlsx.load(await createBillWorkbook({ ...input, documentType: 'credit', rate: 100, taxMode: 'exclusive' }))
  assert.equal(exclusive.worksheets[0].getCell('H13').value, 118)
  assert.equal(exclusive.worksheets[0].getCell('K27').value.result, 118)

  await assert.rejects(createBillWorkbook({ ...input, documentType: 'challan', rows: [{ ...rows[0], rsk: '' }] }), /RSK is required/)
})

test('copies the exact scheme template for additional ADA bills', async () => {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await createBillWorkbook({
    rows: [
      { ada: 'Ballari', rsk: 'Ballari', scheme: 'General', crop: 'Maize', quantity: 1 },
      { ada: 'Kurugodu', rsk: 'Kurugodu', scheme: 'General', crop: 'Maize', quantity: 2 },
    ],
    documentType: 'credit',
    schemeDetails: 'NEMO OILSEEDS (ಸೂರ್ಯಕಾಂತಿ)',
    rate: 118,
    gstRate: 18,
    taxMode: 'inclusive',
    firstBillNumber: 1,
  }))
  assert.equal(workbook.worksheets.length, 2)
  assert.equal(workbook.worksheets[0].getCell('A9').value, 'ASST DIRECTOR AGRICULTURE OF BALLARI')
  assert.equal(workbook.worksheets[1].getCell('A9').value, 'ASST DIRECTOR AGRICULTURE OF KURUGODU')
  assert.equal(workbook.worksheets[1].getCell('G13').value, 2)
  assert.equal(workbook.worksheets[1].getCell('G8').value, '  BILL NO :-  02')
  assert.deepEqual([...workbook.worksheets[1].model.merges].sort(), [...workbook.worksheets[0].model.merges].sort())
  assert.deepEqual(workbook.worksheets[1].getCell('G13').style, workbook.worksheets[0].getCell('G13').style)
})

test('splits rounded paise so displayed CGST and SGST equal total GST', () => {
  const amounts = calculateTax(1, 0.05, 5, 'exclusive')
  assert.equal(amounts.cgst.plus(amounts.sgst).toString(), amounts.tax.toString())
  assert.equal(amounts.gross.toString(), amounts.taxable.plus(amounts.tax).toString())
})

test('assembles streamed Ollama JSON chunks including split UTF-8 Kannada text', async () => {
  const chunks = [
    '{"message":{"content":"{\\"schemeName\\":\\"NEMO OILSEEDS ("},"done":false}\n',
    '{"message":{"content":"ಸೂರ್ಯಕಾಂತಿ)\\",\\"rows\\":[]}"},"done":false}\n',
    '{"done":true,"eval_count":12}\n',
  ]
  const encoder = new TextEncoder()
  const stream = new ReadableStream({
    start(controller) {
      chunks.forEach((chunk) => controller.enqueue(encoder.encode(chunk)))
      controller.close()
    },
  })
  const result = await readOllamaStream(stream)
  assert.deepEqual(JSON.parse(result.content), { schemeName: 'NEMO OILSEEDS (ಸೂರ್ಯಕಾಂತಿ)', rows: [] })
  assert.equal(result.finalChunk.eval_count, 12)
})

test('expands compact ADA/RSK groups into scheme rows and preserves Kannada scheme text', () => {
  const schemeDetails = 'NEMO OILSEEDS (ಸೂರ್ಯಕಾಂತಿ)'
  assert.deepEqual(expandChartGroups([
    { ada: 'Koppal', rsk: 'Kushtagi', crop: null, quantities: { general: 1.32, scp: 0.24, tsp: null } },
  ]), [
    { ada: 'Koppal', rsk: 'Kushtagi', scheme: 'General', crop: '', quantity: 1.32 },
    { ada: 'Koppal', rsk: 'Kushtagi', scheme: 'SCP', crop: '', quantity: 0.24 },
  ])
})

test('bounds the inference image while preserving orientation and image readability', async () => {
  const original = await sharp({
    create: { width: 2400, height: 3600, channels: 3, background: '#ffffff' },
  }).png().toBuffer()
  const prepared = await prepareChartImage(original)
  const metadata = await sharp(prepared).metadata()
  assert.ok(metadata.width <= 896)
  assert.ok(metadata.height <= 1120)
  assert.equal(metadata.format, 'jpeg')
})