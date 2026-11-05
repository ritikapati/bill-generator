import { useEffect, useEffectEvent, useRef, useState } from 'react'
import {
  ArrowRight,
  CheckCircle2,
  ChevronDown,
  CircleHelp,
  ClipboardList,
  FileImage,
  FileText,
  FileSpreadsheet,
  History,
  ImagePlus,
  LoaderCircle,
  Plus,
  RotateCcw,
  ShieldCheck,
  Sparkles,
  Upload,
  X,
} from 'lucide-react'
import './App.css'

const demoRows = [
  { ada: 'Ballari', rsk: 'Ballari', scheme: 'General', crop: 'Maize', quantity: 20 },
  { ada: 'Ballari', rsk: 'Ballari', scheme: 'SCP', crop: 'Maize', quantity: 37 },
  { ada: 'Ballari', rsk: 'Ballari', scheme: 'TSP', crop: 'Maize', quantity: 22.5 },
  { ada: 'Kurugodu', rsk: 'Kurugodu', scheme: 'General', crop: 'Maize', quantity: 9 },
  { ada: 'Kurugodu', rsk: 'Kurugodu', scheme: 'SCP', crop: 'Maize', quantity: 36 },
  { ada: 'Kurugodu', rsk: 'Kurugodu', scheme: 'TSP', crop: 'Maize', quantity: 15 },
  { ada: 'Siruguppa', rsk: 'Siruguppa', scheme: 'General', crop: 'Maize', quantity: 24 },
  { ada: 'Siruguppa', rsk: 'Siruguppa', scheme: 'SCP', crop: 'Maize', quantity: 91.5 },
  { ada: 'Siruguppa', rsk: 'Siruguppa', scheme: 'TSP', crop: 'Maize', quantity: 34.5 },
]

const blankConfig = { rate: '0', gstRate: '0', taxMode: 'inclusive' }
 const initialStartingNumbers = { credit: '1', challan: '341' }
const todayLabel = new Intl.DateTimeFormat('en-IN', { dateStyle: 'medium' }).format(new Date())

function formatQuantity(value) {
  if (value === '' || value === null || value === undefined) return ''
  return Number(value).toLocaleString('en-IN', { maximumFractionDigits: 2 })
}

function App() {
  const fileInput = useRef(null)
  const [file, setFile] = useState(null)
  const [previewUrl, setPreviewUrl] = useState('')
  const [job, setJob] = useState(null)
  const [rows, setRows] = useState([])
  const [documentType, setDocumentType] = useState('credit')
  const [outputFormat, setOutputFormat] = useState('xlsx')
  const [startingNumbers, setStartingNumbers] = useState(initialStartingNumbers)
  const [config, setConfig] = useState(blankConfig)
  const [isUploading, setIsUploading] = useState(false)
  const [error, setError] = useState('')
  const [generated, setGenerated] = useState(null)
  const [activeView, setActiveView] = useState('new')
  const [history, setHistory] = useState([])
  const [modelState, setModelState] = useState('checking')
  const [modelName, setModelName] = useState('qwen3-vl:4b')
  const [dragging, setDragging] = useState(false)

  const checkModel = useEffectEvent(async () => {
    try {
      const response = await fetch('/api/health')
      const data = await response.json()
      setModelName(data.model || 'qwen3-vl:4b')
      setModelState(!data.ollama ? 'offline' : data.modelAvailable ? 'ready' : 'model-missing')
    } catch {
      setModelState('offline')
    }
  })

  const refreshJob = useEffectEvent(async (jobId) => {
    try {
      const response = await fetch(`/api/v1/chart-jobs/${jobId}`)
      const data = await response.json()
      setJob(data)
      if (data.status === 'REVIEW_REQUIRED') {
        setRows(data.rows)
      }
      if (data.status === 'FAILED') setError(data.error || 'Chart extraction failed.')
    } catch {
      setError('Can’t reach the local bill service. Check that the API is running.')
    }
  })

  useEffect(() => {
    const timeout = window.setTimeout(() => checkModel(), 0)
    return () => window.clearTimeout(timeout)
  }, [])

  useEffect(() => {
    if (!job?.jobId || job.demo || ['COMPLETED', 'FAILED', 'REVIEW_REQUIRED'].includes(job.status)) return undefined
    const interval = window.setInterval(() => refreshJob(job.jobId), 1500)
    return () => window.clearInterval(interval)
  }, [job?.jobId, job?.demo, job?.status])

  useEffect(() => () => {
    if (previewUrl) URL.revokeObjectURL(previewUrl)
  }, [previewUrl])

  const validRows = rows.filter((row) => row.ada?.trim() && row.crop?.trim() && row.scheme && (documentType !== 'challan' || row.rsk?.trim()) && row.quantity !== null && row.quantity !== '' && Number.isFinite(Number(row.quantity)) && Number(row.quantity) >= 0)
  const missingRows = rows.length - validRows.length
  const quantityTotal = validRows.reduce((total, row) => total + Number(row.quantity), 0)
  const billGroupCount = new Set(validRows.map((row) => documentType === 'challan'
    ? [row.ada.trim().toLocaleLowerCase('en-IN'), row.rsk.trim().toLocaleLowerCase('en-IN'), row.scheme].join('|')
    : [row.ada.trim().toLocaleLowerCase('en-IN'), row.scheme].join('|'))).size
  const startingNumber = Number(startingNumbers[documentType])
  const isWorking = isUploading || ['QUEUED', 'EXTRACTING_CHART', 'VALIDATING_DATA'].includes(job?.status)

  function resetBatch() {
    setFile(null)
    setJob(null)
    setRows([])
    setConfig(blankConfig)
    setGenerated(null)
    setError('')
    if (fileInput.current) fileInput.current.value = ''
  }

  function chooseFile(nextFile) {
    if (!nextFile) return
    setPreviewUrl(URL.createObjectURL(nextFile))
    setError('')
    setGenerated(null)
    setRows([])
    setJob(null)
    setFile(nextFile)
  }

  async function extractChart() {
    if (!file) return
    setIsUploading(true)
    setError('')
    setRows([])
    setGenerated(null)
    try {
      const form = new FormData()
      form.append('chart', file)
      const response = await fetch('/api/v1/chart-jobs', { method: 'POST', body: form })
      const data = await response.json()
      if (!response.ok) throw new Error(data.error || 'Upload could not be processed.')
      setJob(data)
    } catch (requestError) {
      setError(requestError.message || 'Could not start chart extraction.')
    } finally {
      setIsUploading(false)
    }
  }

  function loadDemo() {
    setError('')
    setFile(null)
    setJob({ jobId: 'demo', status: 'REVIEW_REQUIRED', demo: true })
    setRows(demoRows.map((row) => ({ ...row })))
    setGenerated(null)
  }

  function updateRow(index, field, value) {
    setRows((current) => current.map((row, rowIndex) => rowIndex === index
      ? { ...row, [field]: field === 'quantity' ? (value === '' ? null : Number(value)) : value }
      : row))
  }

  function addRow() {
    setRows((current) => [...current, { ada: '', rsk: '', scheme: 'General', crop: '', quantity: null }])
  }

  async function generateBill() {
    setError('')
    setGenerated(null)
    try {
      const numberWidth = documentType === 'credit' ? 2 : 3
      const suggestedName = `${documentType === 'credit' ? 'credit-bill' : 'delivery-challan'}-${String(startingNumber).padStart(numberWidth, '0')}.${outputFormat}`
      const mimeType = outputFormat === 'pdf'
        ? 'application/pdf'
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      let saveHandle = null
      if (typeof window.showSaveFilePicker === 'function') {
        try {
          saveHandle = await window.showSaveFilePicker({
            suggestedName,
            types: [{ description: outputFormat === 'pdf' ? 'PDF document' : 'Excel workbook', accept: { [mimeType]: [`.${outputFormat}`] } }],
          })
        } catch (pickerError) {
          if (pickerError.name === 'AbortError') throw pickerError
        }
      }
      const response = await fetch('/api/v1/bill-batches', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows, documentType, outputFormat, startingBillNumber: startingNumber, ...config }),
      })
      if (!response.ok) {
        const data = await response.json()
        throw new Error(data.error || 'Bill could not be generated.')
      }
      const blob = await response.blob()
      const disposition = response.headers.get('Content-Disposition') || ''
      const fileName = disposition.match(/filename="?([^";]+)"?/)?.[1] || suggestedName
      if (saveHandle) {
        const writable = await saveHandle.createWritable()
        await writable.write(blob)
        await writable.close()
      } else {
        const url = URL.createObjectURL(blob)
        const anchor = document.createElement('a')
        anchor.href = url
        anchor.download = fileName
        anchor.style.display = 'none'
        document.body.appendChild(anchor)
        anchor.click()
        window.setTimeout(() => anchor.remove(), 1000)
        window.setTimeout(() => URL.revokeObjectURL(url), 60000)
      }
      const bill = {
        id: response.headers.get('X-Bill-Number') || 'Generated',
        documentType: documentType === 'credit' ? 'Credit Bill' : 'Delivery Challan',
        number: response.headers.get('X-Bill-Number') || 'Generated',
        time: response.headers.get('X-Generated-At') || '',
        total: quantityTotal,
      }
      const generatedCount = Number(response.headers.get('X-Bill-Count') || 1)
      setStartingNumbers((current) => ({ ...current, [documentType]: String(startingNumber + generatedCount) }))
      setGenerated(bill)
      setHistory((current) => [bill, ...current])
    } catch (requestError) {
      if (requestError.name !== 'AbortError') setError(requestError.message || 'Bill generation failed.')
    }
  }

  const fileDropProps = {
    onDragOver: (event) => { event.preventDefault(); setDragging(true) },
    onDragLeave: () => setDragging(false),
    onDrop: (event) => { event.preventDefault(); setDragging(false); chooseFile(event.dataTransfer.files?.[0]) },
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <a className="brand" href="#new" onClick={() => setActiveView('new')}>
          <span className="brand-mark">SE</span>
          <span className="brand-name">Shrishti<span>Enterprises</span></span>
        </a>
        <div className="workspace-label">WORKSPACE</div>
        <nav className="side-nav" aria-label="Main navigation">
          <button className={activeView === 'new' ? 'nav-item active' : 'nav-item'} onClick={() => setActiveView('new')}>
            <Plus size={17} strokeWidth={1.8} /> <span>New batch</span>
          </button>
          <button className={activeView === 'history' ? 'nav-item active' : 'nav-item'} onClick={() => setActiveView('history')}>
            <History size={17} strokeWidth={1.8} /> <span>Bill history</span><span className="nav-count">{history.length || ''}</span>
          </button>
        </nav>
        <div className="sidebar-bottom">
          <div className="connection-card">
            <span className={`connection-dot ${modelState}`} />
            <div><strong>Local AI service</strong><small>{modelState === 'ready' ? 'Connected · Ollama' : modelState === 'model-missing' ? `Model not installed · ${modelName}` : modelState === 'checking' ? 'Checking connection…' : 'Not connected'}</small></div>
            <CircleHelp size={15} className="help-icon" aria-label="Ollama connection status" />
          </div>
          <div className="user-profile">
            <span className="avatar">SE</span>
            <span><strong>Invoicing desk</strong><small>Local workspace</small></span>
            <ChevronDown size={15} />
          </div>
        </div>
      </aside>

      <main className="main-area">
        <header className="topbar">
          <div className="breadcrumbs"><span>Documents</span><span className="crumb-slash">/</span><strong>{activeView === 'new' ? 'New batch' : 'Bill history'}</strong></div>
          <div className="topbar-right"><span className="secure-label"><ShieldCheck size={15} /> Local processing</span><span className="date-label">{todayLabel}</span></div>
        </header>

        {activeView === 'new' ? (
          <div className="page-content">
            <div className="page-heading">
              <div><div className="eyebrow">CHART-TO-BILL</div><h1>New bill batch</h1><p>Upload a chart, check every quantity, then prepare your documents.</p></div>
              <button className="button button-quiet" onClick={resetBatch} disabled={!file && rows.length === 0}><RotateCcw size={15} /> Start over</button>
            </div>

            <div className="document-switch" role="group" aria-label="Document type">
              <button className={documentType === 'credit' ? 'selected' : ''} onClick={() => setDocumentType('credit')}><FileSpreadsheet size={16} /> Credit Bill</button>
              <button className={documentType === 'challan' ? 'selected' : ''} onClick={() => setDocumentType('challan')}><ClipboardList size={16} /> Delivery Challan</button>
              <span className="switch-note">{documentType === 'credit' ? 'ADA shown · RSK omitted' : 'ADA and RSK shown'}</span>
            </div>

            <div className="workflow-grid">
              <section className="workflow-main">
                <div className="section-head"><div className="step-number">01</div><div><h2>Source chart</h2><p>Choose a photo or image of the original quantity chart.</p></div><span className="step-state">{file ? 'FILE ADDED' : 'REQUIRED'}</span></div>
                {!file ? (
                  <div className={`upload-panel ${dragging ? 'dragging' : ''}`} {...fileDropProps}>
                    <input ref={fileInput} className="visually-hidden" type="file" accept="image/jpeg,image/png,image/webp" onChange={(event) => chooseFile(event.target.files?.[0])} />
                    <div className="upload-symbol"><ImagePlus size={22} strokeWidth={1.6} /></div>
                    <div className="upload-copy"><strong>Drop your chart here</strong><span>or choose from your device</span></div>
                    <button className="button button-outline" onClick={() => fileInput.current?.click()}><Upload size={15} /> Browse files</button>
                    <div className="upload-meta">JPG, PNG or WebP <span>·</span> Max 12 MB</div>
                    <button className="demo-link" onClick={loadDemo}>Load example data <ArrowRight size={14} /></button>
                  </div>
                ) : (
                  <div className="source-file-row">
                    {previewUrl ? <img className="source-thumb" src={previewUrl} alt="Uploaded chart preview" /> : <span className="source-thumb source-icon"><FileImage size={22} /></span>}
                    <span className="source-file-info"><strong>{file.name}</strong><small>{(file.size / (1024 * 1024)).toFixed(2)} MB · Chart source</small></span>
                    <button className="icon-button" title="Remove chart" aria-label="Remove chart" onClick={resetBatch}><X size={17} /></button>
                    {!rows.length && <button className="button button-primary" onClick={extractChart} disabled={isWorking}><Sparkles size={15} />{isWorking ? 'Reading chart…' : 'Read chart'}</button>}
                  </div>
                )}
                {isWorking && <div className="processing-note"><LoaderCircle className="spin" size={16} /><span><strong>Reading the chart</strong><small>On this CPU-only computer, chart reading can take several minutes.</small></span><span className="processing-status">{job?.status || 'UPLOADING'}</span></div>}
                {error && <div className="error-message" role="alert"><CircleHelp size={16} />{error}</div>}

                {rows.length > 0 && <section className="review-section">
                  <div className="section-head review-heading"><div className="step-number">02</div><div><h2>Review extracted quantities</h2><p>AI suggestions need your approval. Correct unclear or missing cells below.</p></div><span className="review-badge"><ShieldCheck size={14} /> HUMAN REVIEW</span></div>
                  <div className="table-toolbar"><span><strong>{rows.length}</strong> rows found <span className="toolbar-separator">·</span> {validRows.length} ready</span><button className="text-button" onClick={addRow}><Plus size={15} /> Add row</button></div>
                  <div className="table-scroll"><table className="review-table"><thead><tr><th>ADA</th><th>RSK</th><th>SCHEME</th><th>CROP / PRODUCT</th><th className="quantity-heading">QUANTITY</th><th><span className="visually-hidden">Remove row</span></th></tr></thead><tbody>
                    {rows.map((row, index) => <tr key={row.id ?? index} className={!row.ada || row.quantity === null ? 'needs-review' : ''}>
                      <td data-label="ADA"><input aria-label={`ADA row ${index + 1}`} value={row.ada || ''} onChange={(event) => updateRow(index, 'ada', event.target.value)} placeholder="Enter ADA" /></td>
                      <td data-label="RSK"><input aria-label={`RSK row ${index + 1}`} value={row.rsk || ''} onChange={(event) => updateRow(index, 'rsk', event.target.value)} placeholder="Enter RSK" /></td>
                      <td data-label="Scheme"><select aria-label={`Scheme row ${index + 1}`} value={row.scheme || 'General'} onChange={(event) => updateRow(index, 'scheme', event.target.value)}><option>General</option><option>SCP</option><option>TSP</option></select></td>
                      <td data-label="Crop / product"><input aria-label={`Crop row ${index + 1}`} value={row.crop || ''} onChange={(event) => updateRow(index, 'crop', event.target.value)} placeholder="Enter crop" /></td>
                      <td data-label="Quantity"><input className="quantity-input" aria-label={`Quantity row ${index + 1}`} type="number" min="0" step="0.01" value={row.quantity ?? ''} onChange={(event) => updateRow(index, 'quantity', event.target.value)} placeholder="Required" /></td>
                      <td><button className="row-remove" aria-label={`Remove row ${index + 1}`} onClick={() => setRows((current) => current.filter((_, rowIndex) => rowIndex !== index))}><X size={14} /></button></td>
                    </tr>)}
                  </tbody></table></div>
                  <div className="review-foot"><span className={missingRows ? 'warning-inline' : 'success-inline'}>{missingRows ? <CircleHelp size={15} /> : <CheckCircle2 size={15} />}{missingRows ? `${missingRows} row${missingRows > 1 ? 's' : ''} need attention` : 'All quantities are present'}</span><span className="total-quantity">Total quantity <strong>{formatQuantity(quantityTotal)}</strong></span></div>
                </section>}

                {rows.length > 0 && <section className="config-section">
                  <div className="section-head"><div className="step-number">03</div><div><h2>Bill configuration</h2><p>Set the approved rate and tax treatment. These values are never guessed by AI.</p></div></div>
                  <div className="config-fields">
                    <label>Starting bill number <span><input type="number" min="1" step="1" value={startingNumbers[documentType]} onChange={(event) => setStartingNumbers((current) => ({ ...current, [documentType]: event.target.value }))} /></span></label>
                    <label>Rate per unit <span className="input-with-prefix"><span>₹</span><input type="number" min="0" step="0.01" value={config.rate} onChange={(event) => setConfig((current) => ({ ...current, rate: event.target.value }))} /></span></label>
                    <label>GST rate <span className="input-with-suffix"><input type="number" min="0" max="100" step="0.01" value={config.gstRate} onChange={(event) => setConfig((current) => ({ ...current, gstRate: event.target.value }))} /><span>%</span></span></label>
                    <label>Rate mode <select value={config.taxMode} onChange={(event) => setConfig((current) => ({ ...current, taxMode: event.target.value }))}><option value="inclusive">GST inclusive</option><option value="exclusive">GST exclusive</option></select></label>
                  </div>
                  <div className="config-note"><ShieldCheck size={15} /> Confirm tax settings with your accountant before creating official bills.</div>
                </section>}
              </section>

              <aside className="workflow-aside">
                <div className="aside-title"><h3>Batch summary</h3><span className="draft-tag">DRAFT</span></div>
                <div className="summary-row"><span>Document</span><strong>{documentType === 'credit' ? 'Credit Bill' : 'Delivery Challan'}</strong></div>
                <div className="summary-row"><span>Source</span><strong>{file ? file.name : job?.demo ? 'Example data' : 'Not added'}</strong></div>
                <div className="summary-row"><span>Bill groups</span><strong>{billGroupCount || '—'}</strong></div>
                <div className="summary-total"><span>Approved quantity</span><strong>{formatQuantity(quantityTotal) || '—'}</strong></div>
                <div className="download-options" role="radiogroup" aria-label="Download format">
                  <button type="button" role="radio" aria-checked={outputFormat === 'xlsx'} className={outputFormat === 'xlsx' ? 'selected' : ''} onClick={() => setOutputFormat('xlsx')}><FileSpreadsheet size={15} /> Excel</button>
                  <button type="button" role="radio" aria-checked={outputFormat === 'pdf'} className={outputFormat === 'pdf' ? 'selected' : ''} onClick={() => setOutputFormat('pdf')}><FileText size={15} /> PDF</button>
                </div>
                <button className="button button-primary generate-button" onClick={generateBill} disabled={!validRows.length || missingRows > 0 || !Number.isSafeInteger(startingNumber) || startingNumber < 1 || !config.rate || Number(config.rate) < 0 || Number(config.gstRate) < 0 || Number(config.gstRate) > 100}>
                  {outputFormat === 'pdf' ? <FileText size={16} /> : <FileSpreadsheet size={16} />} Generate {outputFormat === 'pdf' ? 'PDF' : 'Excel'} <ArrowRight size={16} />
                </button>
                <p className="generation-disclaimer">Uses the supplied workbook format. Bill numbering resets when the local API restarts.</p>
                {generated && <div className="generated-notice"><CheckCircle2 size={17} /><span><strong>{outputFormat === 'pdf' ? 'PDF' : 'Excel workbook'} downloaded</strong><small>{generated.number} · {generated.documentType}</small></span></div>}
                <div className="aside-divider" />
                <div className="aside-note"><div className="note-icon"><ShieldCheck size={17} /></div><div><strong>Review before issuing</strong><p>Each quantity stays editable until you approve the source table. The model does not calculate tax or bill totals.</p></div></div>
              </aside>
            </div>
          </div>
        ) : (
          <div className="page-content history-page">
            <div className="page-heading"><div><div className="eyebrow">DOCUMENTS</div><h1>Bill history</h1><p>Documents generated in this browser session.</p></div><button className="button button-primary" onClick={() => setActiveView('new')}><Plus size={16} /> New batch</button></div>
            {history.length ? <div className="history-table"><div className="history-row history-header"><span>DOCUMENT</span><span>TYPE</span><span>QUANTITY</span><span>CREATED</span></div>{history.map((item) => <div className="history-row" key={item.id}><span><FileSpreadsheet size={17} /><strong>{item.number}</strong></span><span>{item.documentType}</span><span>{formatQuantity(item.total)}</span><span>{item.time}</span></div>)}</div> : <div className="empty-history"><div className="empty-history-icon"><History size={22} /></div><h2>No generated bills yet</h2><p>Completed sample workbooks will appear here during this session.</p><button className="button button-outline" onClick={() => setActiveView('new')}>Start a batch <ArrowRight size={15} /></button></div>}
          </div>
        )}

        <footer className="app-footer"><span>SHRISHTI ENTERPRISES <span className="footer-dot">·</span> LOCAL PILOT</span><span>Chart data requires human approval</span></footer>
      </main>
    </div>
  )
}

export default App
