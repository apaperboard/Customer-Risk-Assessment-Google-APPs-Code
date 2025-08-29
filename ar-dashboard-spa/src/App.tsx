import React, { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { parseRowsToModel, analyze } from './lib/analysis'

type UploadState = {
  filename: string
  rows: Record<string, any>[]
} | null

export default function App() {
  const [upload, setUpload] = useState<UploadState>(null)
  const [beginBal, setBeginBal] = useState<string>('0')

  const onFile = async (f: File) => {
    const buf = await f.arrayBuffer()
    const wb = XLSX.read(buf, { type: 'array' })
    const wsname = wb.SheetNames.find(n => /input/i.test(n)) || wb.SheetNames[0]
    const ws = wb.Sheets[wsname]
    const json = XLSX.utils.sheet_to_json(ws, { defval: '' }) as Record<string, any>[]
    setUpload({ filename: f.name, rows: json })
  }

  const result = useMemo(() => {
    if (!upload) return null
    const model = parseRowsToModel(upload.rows)
    const bb = Number(beginBal) || 0
    const start = model.firstTransactionDate || model.firstInvoiceDate
    if (!start) return { error: 'No dated rows found.' }
    return analyze(model.invoices, model.payments, start, bb)
  }, [upload, beginBal])

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
    </div>
  )
}
