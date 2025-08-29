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

      {result && 'error' in result && (
        <div style={{ color: 'crimson' }}>{result.error}</div>
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
                  {['Invoice Date','Invoice No','Type','Amount','Closing Date','Term (Days)','Due Date','Days to Pay','Days After Due','Remaining'].map(h => (
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
