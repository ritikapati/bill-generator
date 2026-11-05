import 'dotenv/config'
import { randomUUID } from 'node:crypto'
import express from 'express'
import multer from 'multer'
import { fileTypeFromBuffer } from 'file-type'
import { exportWorkbookPdf } from './pdfExport.js'
import { countBillDocuments, createBillWorkbook } from './billing.js'
import { prepareChartImage } from './chartImage.js'
import { expandChartGroups } from './chartExtraction.js'
import { readOllamaStream } from './ollamaStream.js'

const app = express()
const port = Number(process.env.PORT || 3001)
const ollamaUrl = (process.env.OLLAMA_URL || 'http://127.0.0.1:11434').replace(/\/$/, '')
const ollamaModel = process.env.OLLAMA_MODEL || 'qwen3-vl:4b'
const jobs = new Map()
const usedBillNumbers = { credit: new Set(), challan: new Set() }

app.use(express.json({ limit: '1mb' }))
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 12 * 1024 * 1024, files: 1 } })

const chartSchema = {
  type: 'object',
  properties: {
    groups: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          ada: { type: ['string', 'null'] },
          rsk: { type: ['string', 'null'] },
          crop: { type: ['string', 'null'] },
          quantities: {
            type: 'object',
            properties: {
              general: { type: ['number', 'null'] },
              scp: { type: ['number', 'null'] },
              tsp: { type: ['number', 'null'] },
            },
            required: ['general', 'scp', 'tsp'],
          },
        },
        required: ['ada', 'rsk', 'crop', 'quantities'],
      },
    },
  },
  required: ['groups'],
}

async function getOllamaStatus() {
  try {
    const response = await fetch(`${ollamaUrl}/api/tags`, { signal: AbortSignal.timeout(3000) })
    if (!response.ok) return { connected: false, modelAvailable: false }
    const payload = await response.json()
    const modelAvailable = payload.models?.some((item) => item.name === ollamaModel || item.model === ollamaModel) || false
    return { connected: true, modelAvailable }
  } catch {
    return { connected: false, modelAvailable: false }
  }
}

app.get('/api/health', async (_request, response) => {
  const { connected, modelAvailable } = await getOllamaStatus()
  response.json({ status: 'ok', ollama: connected, modelAvailable, model: ollamaModel })
})

app.post('/api/v1/chart-jobs', upload.single('chart'), (request, response) => {
  if (!request.file) return response.status(400).json({ error: 'Choose a chart image to upload.' })
  fileTypeFromBuffer(request.file.buffer).then((detected) => {
    const supported = new Set(['image/jpeg', 'image/png', 'image/webp'])
    if (!detected || !supported.has(detected.mime)) {
      return response.status(415).json({ error: 'Only valid JPG, PNG or WebP chart images are supported in this pilot.' })
    }
    const job = { jobId: randomUUID(), status: 'QUEUED', createdAt: new Date().toISOString(), rows: null, error: null }
    jobs.set(job.jobId, job)
    response.status(202).json({ jobId: job.jobId, status: job.status })
    prepareChartImage(request.file.buffer)
      .then((imageBuffer) => runExtraction(job, imageBuffer.toString('base64')))
      .catch((error) => {
        job.status = 'FAILED'
        job.elapsedMs = Date.now() - (job.startedAt || Date.now())
        job.error = error.name === 'TimeoutError'
          ? 'Chart reading exceeded the 10-minute local inference limit. Try a clearer or cropped chart image.'
          : error.message || 'Local chart extraction failed.'
        console.error(`Chart job ${job.jobId} failed after ${Math.round(job.elapsedMs / 1000)}s: ${job.error}`)
      })
  }).catch(() => {
    if (!response.headersSent) response.status(400).json({ error: 'Chart upload could not be read.' })
  })
})

async function runExtraction(job, image) {
  job.startedAt = Date.now()
  job.status = 'EXTRACTING_CHART'
  const ollamaStatus = await getOllamaStatus()
  if (!ollamaStatus.connected) throw new Error('Ollama is not responding at the configured local address. Start Ollama and try again.')
  if (!ollamaStatus.modelAvailable) throw new Error(`Model ${ollamaModel} is not installed. In PowerShell run: ollama pull ${ollamaModel}`)

  const message = {
    role: 'user',
    content: 'Read this agricultural chart carefully and return the required JSON only. For each ADA and RSK combination, copy the crop/product name exactly into crop, including Kannada characters in Kannada script. Do not translate Kannada, replace it with Latin transliteration, normalize spelling, or infer a likely crop name. Preserve all visible marks, vowel signs, and punctuation. If any part of a name is not legible, return null for crop so it can be corrected during review; do not guess. Put the printed General, SCP, and TSP quantities into quantities.general, quantities.scp, and quantities.tsp. Use null for blank or unclear cells. Copy names and quantities as written. Keep the response concise.',
    images: [image],
  }
  let payload
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const result = await fetch(`${ollamaUrl}/api/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      signal: AbortSignal.timeout(600000),
      body: JSON.stringify({
        model: ollamaModel,
        stream: true,
        think: false,
        keep_alive: '30m',
        options: {
          num_ctx: 8192,
          num_predict: 4096,
          temperature: attempt === 0 ? 0.1 : 0.2,
          top_p: 0.9,
          repeat_penalty: attempt === 0 ? 1.1 : 1.2,
          repeat_last_n: 128,
        },
        format: chartSchema,
        messages: [message],
      }),
    })
    if (!result.ok) {
      const detail = await result.text()
      throw new Error(`Ollama returned ${result.status}: ${detail.slice(0, 240) || 'check the model service.'}`)
    }
    try {
      payload = await readOllamaStream(result.body)
      break
    } catch (error) {
      if (attempt === 0 && /repeat limit|repetitive output/i.test(error.message)) {
        console.warn(`Chart job ${job.jobId} returned repetitive output; retrying with stronger repetition controls.`)
        continue
      }
      throw error
    }
  }
  if (payload.finalChunk?.done_reason === 'length') {
    throw new Error(`Ollama reached its ${payload.finalChunk.eval_count || 'available'}-token response limit before completing chart JSON. The extraction output was truncated; try a clearer or smaller chart image.`)
  }
  let parsed
  try { parsed = JSON.parse(payload.content || '') } catch {
    throw new Error('The model returned unreadable structured data. Try a clearer chart image.')
  }
  if (!Array.isArray(parsed.groups)) throw new Error('No chart groups were returned by the model.')
  job.status = 'VALIDATING_DATA'
  job.elapsedMs = Date.now() - job.startedAt
  job.rows = expandChartGroups(parsed.groups)
  job.status = 'REVIEW_REQUIRED'
  job.generatedTokens = payload.finalChunk?.eval_count || null
  console.log(`Chart job ${job.jobId} extracted ${job.rows.length} rows in ${Math.round(job.elapsedMs / 1000)}s`)
}

app.get('/api/v1/chart-jobs/:id', (request, response) => {
  const job = jobs.get(request.params.id)
  if (!job) return response.status(404).json({ error: 'Chart job not found.' })
  response.json(job)
})

app.put('/api/v1/chart-jobs/:id/extraction', (request, response) => {
  const job = jobs.get(request.params.id)
  if (!job) return response.status(404).json({ error: 'Chart job not found.' })
  if (!Array.isArray(request.body.rows)) return response.status(400).json({ error: 'Rows are required.' })
  job.rows = request.body.rows
  response.json({ status: 'REVIEW_REQUIRED', rows: job.rows })
})

app.post('/api/v1/bill-batches', async (request, response) => {
  try {
    const { documentType, outputFormat } = request.body
    if (!Object.hasOwn(usedBillNumbers, documentType)) throw new Error('Choose Credit Bill or Delivery Challan.')
    if (!['xlsx', 'pdf'].includes(outputFormat)) throw new Error('Choose Excel or PDF output format.')
    const billCount = countBillDocuments(request.body.rows, documentType)
    const firstBillNumber = Number(request.body.startingBillNumber)
    if (!Number.isSafeInteger(firstBillNumber) || firstBillNumber < 1) throw new Error('Enter a positive whole starting bill number.')
    const numberWidth = documentType === 'credit' ? 2 : 3
    const requestedNumbers = Array.from({ length: billCount }, (_, index) => String(firstBillNumber + index).padStart(numberWidth, '0'))
    const conflict = requestedNumbers.find((number) => usedBillNumbers[documentType].has(number))
    if (conflict) return response.status(409).json({ error: `${documentType === 'credit' ? 'Credit Bill' : 'Delivery Challan'} number ${conflict} was already generated in this API session. Choose another starting number.` })
    requestedNumbers.forEach((number) => usedBillNumbers[documentType].add(number))
    try {
      const workbook = Buffer.from(await createBillWorkbook({ ...request.body, firstBillNumber }))
      const output = outputFormat === 'pdf' ? await exportWorkbookPdf(workbook) : workbook
      const billNumber = requestedNumbers[0]
      const fileType = documentType === 'credit' ? 'credit-bill' : 'delivery-challan'
      const extension = outputFormat === 'pdf' ? 'pdf' : 'xlsx'
      response.set({
        'Content-Type': outputFormat === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${fileType}-${billNumber}.${extension}"`,
        'X-Bill-Number': billNumber,
        'X-Bill-Count': String(billCount),
        'X-Generated-At': new Date().toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' }),
        'Cache-Control': 'no-store',
      })
      response.send(output)
    } catch (error) {
      requestedNumbers.forEach((number) => usedBillNumbers[documentType].delete(number))
      throw error
    }
  } catch (error) {
    response.status(400).json({ error: error.message || 'Bill generation failed.' })
  }
})

app.use((error, _request, response, _next) => {
  if (error instanceof multer.MulterError && error.code === 'LIMIT_FILE_SIZE') {
    return response.status(413).json({ error: 'Chart image must be 12 MB or smaller.' })
  }
  response.status(500).json({ error: 'The local bill service encountered an unexpected error.' })
})

app.listen(port, '0.0.0.0', () => {
  console.log(`Chart-to-Bill API listening on http://localhost:${port}`)
  console.log(`Ollama endpoint: ${ollamaUrl} (model: ${ollamaModel})`)
})
