import React, { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { parseRowsToModel, analyze } from './lib/analysis'
import { extractTable } from './lib/importer'

type UploadState = {
  filename: string
  rows: Record<string, any>[]
} | null

export default function App() {
  const [upload, setUpload] = useState<UploadState>(null)
  const [beginBal, setBeginBal] = useState<string>('0')
  const [toast, setToast] = useState<string | null>(null)
  const [showDebug, setShowDebug] = useState<boolean>(false)

  const onFile = async (f: File) => {
    console.log('[upload] file selected:', f.name, f.size)
    const buf = await f.arrayBuffer()
    const wb = XLSX.read(buf, { type: 'array' })
    const wsname = wb.SheetNames.find(n => /input/i.test(n)) || wb.SheetNames[0]
    const ws = wb.Sheets[wsname]
    console.debug('[upload] sheets:', wb.SheetNames, 'chosen:', wsname)
    const { rows, autoBeginBalance } = extractTable(ws)
    console.debug('[upload] rows parsed:', rows.length, 'autoBeginBalance:', autoBeginBalance)
    setUpload({ filename: f.name, rows })
    if (autoBeginBalance != null && beginBal === '0') {
      const v = String(autoBeginBalance)
      setBeginBal(v)
      setToast(`Opening balance detected: ${Number(v).toLocaleString()} TRY`)
      setTimeout(() => setToast(null), 4000)
    }
  }

  const result = useMemo(() => {
    if (!upload) return null
    const model = parseRowsToModel(upload.rows)
    const bb = Number(beginBal) || 0
    const start = model.firstTransactionDate || model.firstInvoiceDate
    if (!start) return { error: 'No dated rows found.' }
    return analyze(model.invoices, model.payments, start, bb)
  }, [upload, beginBal])

  useEffect(() => {
    if (!result) return
    if ('error' in result) {
      console.warn('[analysis] error:', result.error)
    } else {
      console.log('[analysis] summary:', {
        invoices: result.invoices.length,
        months: result.months.length,
        startDate: result.startDate,
        aging: result.aging,
      })
      console.debug('[analysis] metrics:', result.metrics)
    }
  }, [result])

  const exportToExcel = () => {
    if (!result || 'error' in result) return
    const wb = XLSX.utils.book_new()

    const metricsRows = result.metrics.map(m => ({ Metric: m.label, Value: m.value, Assessment: m.assess }))
    const wsMetrics = XLSX.utils.json_to_sheet(metricsRows)
    XLSX.utils.book_append_sheet(wb, wsMetrics, 'Metrics')

    const agingLabels = ['0-30 days','31-60 days','61-90 days','91+ days']
    const agingRows = agingLabels.map((lbl, i) => ({ Bucket: lbl, Outstanding: result.aging[i] }))
    const wsAging = XLSX.utils.json_to_sheet(agingRows)
    XLSX.utils.book_append_sheet(wb, wsAging, 'Aging')

    const analysisRows = result.invoices.map(inv => {
      const closing = inv.paid && inv.closingDate ? inv.closingDate : null
      const daysToPay = closing ? Math.round(((+closing) - (+inv.invoiceDate))/86400000) : ''
      const dueDate = new Date(+inv.invoiceDate + inv.term*86400000)
      const daysAfterDue = typeof daysToPay === 'number' ? (daysToPay - inv.term) : ''
      const fmtDate = (d: Date | null) => d ? new Date(d).toLocaleDateString() : ''
      return {
        'Invoice Date': fmtDate(inv.invoiceDate),
        'Invoice No': inv.invoiceNum || '',
        'Type': inv.type || '',
        'Amount': inv.amount,
        'Closing Date': fmtDate(closing as any),
        'Term (Days)': inv.term,
        'Due Date': inv.paid ? fmtDate(dueDate) : '',
        'Days to Pay': typeof daysToPay === 'number' ? daysToPay : '',
        'Days After Due': typeof daysAfterDue === 'number' ? daysAfterDue : '',
        'Remaining': inv.remaining,
      }
    })
    const wsAnalysis = XLSX.utils.json_to_sheet(analysisRows)
    XLSX.utils.book_append_sheet(wb, wsAnalysis, 'Analysis')

    const trendRows = result.months.map(m => ({ Month: new Date(m.dt).toLocaleDateString(), 'Avg Days to Pay': m.avg }))
    const wsTrend = XLSX.utils.json_to_sheet(trendRows)
    XLSX.utils.book_append_sheet(wb, wsTrend, 'Trend')

    const base = upload?.filename ? upload.filename.replace(/\.[^.]+$/, '') : 'export'
    XLSX.writeFile(wb, `${base}-ar-analysis.xlsx`)
  }

  return (
    <div style={{ fontFamily: 'system-ui, sans-serif', padding: 16 }}>
      <h1>AR Analysis Dashboard (Client-side)</h1>
      <p>Drop an Excel/CSV exported from your ERP or click to choose a file. Data stays in your browser.</p>
      <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 12 }}>
        <label style={{ border: '1px solid #ccc', padding: '8px 12px', borderRadius: 6, cursor: 'pointer', background: '#fafafa' }}>
          <input type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={(e) => {
            const f = e.target.files?.[0]
            if (f) onFile(f)
          }} />
          Upload File
        </label>
        <span>{upload ? upload.filename : 'No file selected'}</span>
        <div>|</div>
        <label>Beginning Balance (TRY): <input value={beginBal} onChange={e => setBeginBal(e.target.value)} style={{ width: 120 }} /></label>
        <div>|</div>
        <button
          onClick={exportToExcel}
          title="Export analysis to Excel"
          disabled={!result || ('error' in (result as any))}
          style={{ border: '1px solid #ccc', padding: '8px 12px', borderRadius: 6, cursor: (!result || ('error' in (result as any))) ? 'not-allowed' : 'pointer', background: '#f0f9ff', opacity: (!result || ('error' in (result as any))) ? 0.6 : 1 }}
        >
          Export Excel
        </button>
      </div>
      <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 12 }}>
        <button onClick={() => setShowDebug(v => !v)} style={{ border: '1px solid #ccc', padding: '6px 10px', borderRadius: 6, cursor: 'pointer', background: '#f6f6f6' }}>
          {showDebug ? 'Hide Debug' : 'Show Debug'}
        </button>
      </div>

      {result && 'error' in result && (
        <div style={{ color: 'crimson' }}>{result.error}</div>
      )}

      {showDebug && (
        <DebugPanel />
      )}

      {result && !('error' in result) && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 24 }}>
          <div>
            <h2>Metrics</h2>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>Metric</th>
                  <th style={{ textAlign: 'right', borderBottom: '1px solid #ddd', padding: 6 }}>Value</th>
                  <th style={{ textAlign: 'center', borderBottom: '1px solid #ddd', padding: 6 }}>Assessment</th>
                </tr>
              </thead>
              <tbody>
                {result.metrics.map((m, i) => {
                  const isPct = m.label.includes('%')
                  const isDays = m.label.toLowerCase().includes('day')
                  const fmt = (v: any) => {
                    if (v === '') return ''
                    if (isPct) return (v as number).toLocaleString(undefined, { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 })
                    if (typeof v === 'number') return v.toLocaleString(undefined, { maximumFractionDigits: 0 })
                    return String(v)
                  }
                  const color = m.assess === 'Good' ? '#c6efce' : m.assess === 'Average' ? '#ffe6cc' : m.assess === 'Poor' ? '#f4a7a7' : 'transparent'
                  return (
                    <tr key={i}>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{m.label}</td>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0', textAlign: 'right' }}>{fmt(m.value)}</td>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0', textAlign: 'center', background: color }}>{m.assess}</td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>

          <div>
            <h2>Aging Buckets</h2>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>Bucket</th>
                  <th style={{ textAlign: 'right', borderBottom: '1px solid #ddd', padding: 6 }}>Outstanding (TRY)</th>
                </tr>
              </thead>
              <tbody>
                {['0-30 days','31-60 days','61-90 days','91+ days'].map((lbl, i) => (
                  <tr key={lbl}>
                    <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{lbl}</td>
                    <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0', textAlign: 'right' }}>{(result.aging[i]).toLocaleString(undefined, { maximumFractionDigits: 0 })}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div style={{ gridColumn: '1 / span 2' }}>
            <h2>Analysis Table</h2>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  {['Invoice Date','Invoice No','Type','Amount','Closing Date','Term (Days)','Due Date','Days to Pay','Days After Due','Remaining','AR Balance'].map(h => (
                    <th key={h} style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {result.invoices.map((inv, i) => {
                  const daysToPay = (inv.paid && inv.closingDate) ? Math.round(((+inv.closingDate) - (+inv.invoiceDate))/86400000) : ''
                  const dueDate = new Date(+inv.invoiceDate + inv.term*86400000)
                  const daysAfterDue = (typeof daysToPay === 'number') ? (daysToPay - inv.term) : ''
                  const fmtDate = (d: Date | null) => d ? d.toLocaleDateString() : ''
                  const c = (typeof daysAfterDue === 'number' && daysAfterDue > 0) ? '#ffefef' : 'transparent'
                  return (
                    <tr key={i} style={{ background: c }}>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{fmtDate(inv.invoiceDate)}</td>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{inv.invoiceNum || ''}</td>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{inv.type || ''}</td>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{inv.amount.toLocaleString()}</td>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{fmtDate(inv.paid ? inv.closingDate! : null)}</td>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{inv.term}</td>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{inv.paid ? dueDate.toLocaleDateString() : ''}</td>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{typeof daysToPay === 'number' ? daysToPay.toLocaleString() : ''}</td>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{typeof daysAfterDue === 'number' ? daysAfterDue.toLocaleString() : ''}</td>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{inv.remaining.toLocaleString()}</td>
                      <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{(inv.running ?? 0).toLocaleString()}</td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
      {toast && (
        <div style={{ position: 'fixed', right: 16, bottom: 16, background: '#111', color: 'white', padding: '8px 12px', borderRadius: 8, boxShadow: '0 4px 12px rgba(0,0,0,0.2)', zIndex: 9999 }}>
          {toast}
        </div>
      )}
    </div>
  )
}

function DebugPanel() {
  // Read from global debug object populated by importer/analysis
  const dbg: any = (globalThis as any).__arDebug || {}
  const headersLower: string[] = dbg.headersLower || []
  const descColsHeaders: string[] = dbg.descColsHeaders || []
  const cPayTpIndex: number | undefined = dbg.cPayTpIndex
  const cMaturityIndex: number | undefined = dbg.cMaturityIndex
  const checkInspect: any[] = dbg.checkInspect || []
  const checkNoMatExamples: any[] = dbg.checkNoMatExamples || []
  const withMat = checkInspect.filter(x => !!x.maturity).length
  const withoutMat = checkInspect.length - withMat
  const reconcile = dbg.reconcile || {}

  return (
    <div style={{ border: '1px solid #ddd', borderRadius: 8, padding: 12, background: '#fffef8', marginBottom: 16 }}>
      <div style={{ fontWeight: 600, marginBottom: 8 }}>Debug Summary</div>
      <div style={{ display: 'grid', gridTemplateColumns: '220px 1fr', rowGap: 6, columnGap: 8 }}>
        <div>Headers (lower):</div>
        <div>{headersLower.join(', ') || 'n/a'}</div>
        <div>Description columns:</div>
        <div>{descColsHeaders.join(', ') || 'n/a'}</div>
        <div>Pay Type column index:</div>
        <div>{typeof cPayTpIndex === 'number' ? String(cPayTpIndex) : 'n/a'}</div>
        <div>Maturity column index:</div>
        <div>{typeof cMaturityIndex === 'number' ? String(cMaturityIndex) : 'n/a'}</div>
        <div>Checks parsed (with/without maturity):</div>
        <div>{withMat} / {withoutMat}</div>
        <div>Reconcile (expected vs computed):</div>
        <div>
          {typeof reconcile.expectedOutstanding === 'number' ? reconcile.expectedOutstanding.toLocaleString() : 'n/a'}
          {' '}vs{' '}
          {typeof reconcile.computedOutstanding === 'number' ? reconcile.computedOutstanding.toLocaleString() : 'n/a'}
          {typeof reconcile.delta === 'number' ? ` (delta ${reconcile.delta.toLocaleString()})` : ''}
        </div>
        {checkNoMatExamples.length > 0 && (
          <>
            <div>First no-maturity example:</div>
            <div style={{ whiteSpace: 'pre-wrap' }}>{String(checkNoMatExamples[0]?.desc || '').slice(0, 200)}</div>
          </>
        )}
      </div>
      {checkInspect.length > 0 && (
        <div style={{ marginTop: 10 }}>
          <div style={{ fontWeight: 600, marginBottom: 6 }}>First 3 Check Rows</div>
          <ol>
            {checkInspect.slice(0,3).map((x,i) => (
              <li key={i} style={{ marginBottom: 4 }}>
                <span style={{ color: '#555' }}>{String(x.desc).slice(0,120)}</span>
                {' '}| maturity: {x.maturity ? new Date(x.maturity).toLocaleDateString() : '—'}
                {' '}| raw: <span style={{ color: '#777' }}>{String(x.payTypeRaw).slice(0,60)}</span>
              </li>
            ))}
          </ol>
        </div>
      )}
      <div style={{ marginTop: 8, color: '#666' }}>
        Full details available in Console via window.__arDebug
      </div>
    </div>
  )
}
