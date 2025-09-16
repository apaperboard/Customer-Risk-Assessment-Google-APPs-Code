import React, { useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import { parseRowsToModel, analyze } from './lib/analysis'
import { extractTable } from './lib/importer'
import { initFirebase, isFirebaseReady, onUser, signInWithGoogle, signOutUser, saveLatestReport as fbSave, loadLatestReport as fbLoad } from './lib/firebase'

type UploadState = {
  filename: string
  rows: Record<string, any>[]
} | null

export default function App() {
  const [upload, setUpload] = useState<UploadState>(null)
  const [beginBal, setBeginBal] = useState<string>('0')
  const [beginBalAuto, setBeginBalAuto] = useState<boolean>(true)
  const [toast, setToast] = useState<string | null>(null)
  const [showDebug, setShowDebug] = useState<boolean>(false)
  const [lang, setLang] = useState<'en'|'tr'|'ar'>('en')
  // Customer key (auto from uploaded sheet B1)
  const [customerKey, setCustomerKey] = useState<string>('')
  // Firebase auth state
  const [userEmail, setUserEmail] = useState<string>('')
  // Initialize Firebase if config provided globally
  const fbEnabled = initFirebase()

  // ---- Local (browser) storage via IndexedDB (fallback when Firebase not configured) ----
  function idbOpen(): Promise<IDBDatabase> {
    return new Promise((resolve, reject) => {
      const req = window.indexedDB.open('ar-dashboard-db', 1)
      req.onupgradeneeded = () => {
        const db = req.result
        if (!db.objectStoreNames.contains('latestReports')) {
          db.createObjectStore('latestReports')
        }
      }
      req.onsuccess = () => resolve(req.result)
      req.onerror = () => reject(req.error)
    })
  }
  async function idbSet(key: string, value: any) {
    const db = await idbOpen()
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction('latestReports', 'readwrite')
      const store = tx.objectStore('latestReports')
      const req = store.put(value, key)
      req.onsuccess = () => resolve()
      req.onerror = () => reject(req.error)
    })
    db.close()
  }
  async function idbGet<T = any>(key: string): Promise<T | null> {
    const db = await idbOpen()
    const val = await new Promise<T | null>((resolve, reject) => {
      const tx = db.transaction('latestReports', 'readonly')
      const store = tx.objectStore('latestReports')
      const req = store.get(key)
      req.onsuccess = () => resolve(req.result ?? null)
      req.onerror = () => reject(req.error)
    })
    db.close()
    return val
  }
  async function idbDelete(key: string): Promise<void> {
    const db = await idbOpen()
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction('latestReports', 'readwrite')
      const store = tx.objectStore('latestReports')
      const req = store.delete(key)
      req.onsuccess = () => resolve()
      req.onerror = () => reject(req.error)
    })
    db.close()
  }
  async function idbClearAll(): Promise<void> {
    const db = await idbOpen()
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction('latestReports', 'readwrite')
      const store = tx.objectStore('latestReports')
      const req = store.clear()
      req.onsuccess = () => resolve()
      req.onerror = () => reject(req.error)
    })
    db.close()
  }
  async function idbListKeys(): Promise<string[]> {
    const db = await idbOpen()
    const keys: string[] = await new Promise((resolve, reject) => {
      const tx = db.transaction('latestReports', 'readonly')
      const store: any = tx.objectStore('latestReports')
      if (typeof store.getAllKeys === 'function') {
        const req = store.getAllKeys()
        req.onsuccess = () => resolve((req.result || []).map((k: any) => String(k)))
        req.onerror = () => reject(req.error)
      } else {
        const out: string[] = []
        const cursorReq = store.openCursor()
        cursorReq.onsuccess = (e: any) => {
          const cursor = e.target.result
          if (cursor) { out.push(String(cursor.key)); cursor.continue() } else resolve(out)
        }
        cursorReq.onerror = () => reject(cursorReq.error)
      }
    })
    db.close()
    return keys.sort((a,b) => a.localeCompare(b))
  }
  const i18n: Record<'en'|'tr'|'ar', Record<string,string>> = {
    en: { title: 'AR Analysis Dashboard (Client-side)', instructions: 'Drop an Excel/CSV exported from your ERP or click to choose a file. Data stays in your browser.', uploadFile: 'Upload File', noFile: 'No file selected', beginBal: 'Beginning Balance (TRY)', exportExcel: 'Export Excel', showDebug: 'Show Debug', hideDebug: 'Hide Debug', metrics: 'Metrics', metric: 'Metric', value: 'Value', assessment: 'Assessment', aging: 'Aging Buckets', bucket: 'Bucket', outstanding: 'Outstanding (TRY)', analysisTable: 'Analysis Table', invoiceDate: 'Invoice Date', invoiceNo: 'Invoice No', type: 'Type', amount: 'Amount', closingDate: 'Closing Date', termDays: 'Term (Days)', dueDate: 'Due Date', daysToPay: 'Days to Pay', daysAfterDue: 'Days After Due', remaining: 'Remaining', arBalance: 'AR Balance', ledger: 'Ledger', date: 'Date', description: 'Description', ref: 'Ref', debit: 'Debit', credit: 'Credit', running: 'Running Balance', language: 'Language' },
    tr: { title: 'AL Analiz Panosu (İstemci tarafı)', instructions: 'ERP’nizden dışa aktarılan Excel/CSV dosyasını bırakın veya tıklayıp seçin. Veriler tarayıcınızda kalır.', uploadFile: 'Dosya Yükle', noFile: 'Dosya seçilmedi', beginBal: 'Açılış Bakiyesi (TRY)', exportExcel: 'Excel’e Aktar', showDebug: 'Hata Ayıklamayı Göster', hideDebug: 'Hata Ayıklamayı Gizle', metrics: 'Metikler', metric: 'Metik', value: 'Değer', assessment: 'Değerlendirme', aging: 'Vade Yaşlandırma', bucket: 'Kova', outstanding: 'Bakiye (TRY)', analysisTable: 'Analiz Tablosu', invoiceDate: 'Fatura Tarihi', invoiceNo: 'Fatura No', type: 'Tür', amount: 'Tutar', closingDate: 'Kapanış Tarihi', termDays: 'Vade (Gün)', dueDate: 'Vade Tarihi', daysToPay: 'Ödeme Günleri', daysAfterDue: 'Vade Sonrası Gün', remaining: 'Kalan', arBalance: 'AR Bakiye', ledger: 'Yevmiye', date: 'Tarih', description: 'Açıklama', ref: 'Ref', debit: 'Borç', credit: 'Alacak', running: 'Bakiye', language: 'Dil' },
    ar: { title: 'لوحة تحليل الذمم (على المتصفح)', instructions: 'أسقط ملف Excel/CSV من نظام ERP أو اختر ملفاً. تبقى البيانات في المتصفح.', uploadFile: 'رفع ملف', noFile: 'لم يتم اختيار ملف', beginBal: 'الرصيد الافتتاحي (ليرة)', exportExcel: 'تصدير إلى Excel', showDebug: 'إظهار التصحيح', hideDebug: 'إخفاء التصحيح', metrics: 'المؤشرات', metric: 'المؤشر', value: 'القيمة', assessment: 'التقييم', aging: 'أعمار الديون', bucket: 'الفئة', outstanding: 'الرصيد (ليرة)', analysisTable: 'جدول التحليل', invoiceDate: 'تاريخ الفاتورة', invoiceNo: 'رقم الفاتورة', type: 'النوع', amount: 'المبلغ', closingDate: 'تاريخ الإقفال', termDays: 'المدة (أيام)', dueDate: 'تاريخ الاستحقاق', daysToPay: 'أيام السداد', daysAfterDue: 'أيام بعد الاستحقاق', remaining: 'المتبقي', arBalance: 'رصيد الذمم', ledger: 'دفتر القيود', date: 'التاريخ', description: 'الوصف', ref: 'المرجع', debit: 'مدين', credit: 'دائن', running: 'الرصيد المتراكم', language: 'اللغة' }
  }
  const t = (k: string) => i18n[lang][k] || k
  const locale = lang === 'tr' ? 'tr-TR' : lang === 'ar' ? 'ar-EG' : 'en-US'

  // Metric name translations (by original analysis label)
  const metricNames: Record<string, { en: string; tr: string; ar: string }> = {
    'Average Days to Pay (Paid Only)': {
      en: 'Average Days to Pay (Paid Only)',
      tr: 'Ödenenler İçin Ortalama Ödeme Günleri',
      ar: 'متوسط أيام السداد (المدفوعة فقط)'
    },
    'Weighted Avg Age of Unpaid Invoices (Days)': {
      en: 'Weighted Avg Age of Unpaid Invoices (Days)',
      tr: 'Ödenmemiş Faturaların Ağırlıklı Ortalama Yaşı (Gün)',
      ar: 'متوسط عمر الفواتير غير المدفوعة (مرجح) (أيام)'
    },
    '% of Unpaid Invoices Overdue': {
      en: '% of Unpaid Invoices Overdue',
      tr: '% Gecikmiş Ödenmemiş Fatura',
      ar: '% من الفواتير غير المدفوعة المتأخرة'
    },
    'Blended Average Days to Pay': {
      en: 'Blended Average Days to Pay',
      tr: 'Harmanlanmış Ortalama Ödeme Günü',
      ar: 'متوسط أيام السداد الممزوج'
    },
    'Average Check Maturity Duration (Days)': {
      en: 'Average Check Maturity Duration (Days)',
      tr: 'Çek Ortalama Vade Süresi (Gün)',
      ar: 'متوسط مدة استحقاق الشيك (أيام)'
    },
    'Avg Maturity Over By (Days)': {
      en: 'Avg Maturity Over By (Days)',
      tr: 'Ortalama Vade Aşımı (Gün)',
      ar: 'متوسط تجاوز الاستحقاق (أيام)'
    },
    '% of Payments Over Term': {
      en: '% of Payments Over Term',
      tr: '% Vade Aşımı Ödeme',
      ar: '% من المدفوعات خارج المدة'
    },
    'Customer Risk Rating': {
      en: 'Customer Risk Rating',
      tr: 'Müşteri Risk Notu',
      ar: 'تصنيف مخاطر العميل'
    }
  }
  // Extra labels added later (override when present)
  const metricNamesExtra: Record<string, { en: string; tr: string; ar: string }> = {
    '% of Checks Over Term': { en: '% of Checks Over Term', tr: '% Vadesi Aşılmış Çekler', ar: '% من الشيكات فوق الأجل' },
    '% of Payments Delivered After Term': { en: '% of Payments Delivered After Term', tr: '% Vadeden Sonra Teslim Edilen Ödemeler', ar: '% المدفوعات بعد الأجل' },
    'Average Monthly Purchases (TRY)': { en: 'Average Monthly Purchases (TRY)', tr: 'Aylık Ortalama Alımlar (TRY)', ar: 'متوسط المشتريات الشهري (TRY)' },
    'Credit Limit (TRY)': { en: 'Credit Limit (TRY)', tr: 'Kredi Limiti (TRY)', ar: 'حد الائتمان (TRY)' },
    'Customer Risk Rating': { en: 'Customer Risk Rating', tr: 'Müşteri Risk Notu', ar: 'تصنيف مخاطر العميل' },
    'Available Credit (TRY)': { en: 'Available Credit (TRY)', tr: 'Kullanılabilir Kredi (TRY)', ar: 'الائتمان المتاح (TRY)' },
    'Overdue Balance as % of Credit Limit': { en: 'Overdue Balance as % of Credit Limit', tr: 'Kredi Limitine Göre Vadesi Geçmiş Bakiye %', ar: 'الرصيد المتأخر كنسبة من حد الائتمان' },
  }
  const metricNamesAll = { ...metricNames, ...metricNamesExtra }
  const assessNames: Record<'en'|'tr'|'ar', Record<string,string>> = {
    en: { Good: 'Good', Average: 'Average', Poor: 'Poor' },
    tr: { Good: 'İyi', Average: 'Orta', Poor: 'Zayıf' },
    ar: { Good: 'جيد', Average: 'متوسط', Poor: 'ضعيف' }
  }

  // Metric notes (tooltips in UI)
  const metricNotes: Record<string, string> = {
    '% of Checks Over Term': 'Checks with maturity duration greater than expected term divided by total checks with maturity.',
    '% of Payments Delivered After Term': 'Share of paid invoices where (payment date - invoice date) exceeds the invoice term; all payment types.',
    'Customer Risk Rating': 'Composite rating (Good/Average/Poor) based on weighted metrics.',
    'Average Monthly Purchases (TRY)': 'Total invoiced in the period divided by months since start.',
    'Credit Limit (TRY)': 'Average monthly purchases x risk/term multiplier; rounded up to the nearest 10,000.',
    'Available Credit (TRY)': 'Credit limit minus current open balance; not less than zero.',
    'Overdue Balance as % of Credit Limit': 'Unpaid balance past due (by term) divided by assigned credit limit.'
  }

  const onFile = async (f: File) => {
    console.log('[upload] file selected:', f.name, f.size)
    const buf = await f.arrayBuffer()
    const wb = XLSX.read(buf, { type: 'array' })
    const wsname = wb.SheetNames.find(n => /input/i.test(n)) || wb.SheetNames[0]
    const ws = wb.Sheets[wsname]
    // Detect customer name from B1
    try {
      const b1 = (ws as any)?.['B1']?.v
      if (b1 != null && String(b1).trim() !== '') {
        const key = String(b1).trim()
        setCustomerKey(key)
        console.debug('[upload] customer (B1):', key)
      }
    } catch {}
    setLoadedResult(null); console.debug('[upload] sheets:', wb.SheetNames, 'chosen:', wsname)
    const { rows, autoBeginBalance } = extractTable(ws)
    console.debug('[upload] rows parsed:', rows.length, 'autoBeginBalance:', autoBeginBalance)
    setUpload({ filename: f.name, rows })
    if (beginBalAuto) {
      const v = autoBeginBalance != null ? String(autoBeginBalance) : '0'
      const prev = beginBal
      setBeginBal(v)
      if (autoBeginBalance != null) {
        setToast(`Opening balance detected: ${Number(v).toLocaleString()} TRY`)
      } else if (prev !== '0') {
        setToast('No opening balance found. Defaulted to 0.')
      }
      setTimeout(() => setToast(null), 4000)
    }
  }

  const [loadedResult, setLoadedResult] = useState<any | null>(null)
  const computedResult = useMemo(() => {
    if (!upload) return null
    const model = parseRowsToModel(upload.rows)
    // Parse beginning balance robustly (supports commas and different locales)
    const bb = (() => {
      const s = String(beginBal ?? '').trim()
      const onlyDigits = s.replace(/[^0-9,\.-]/g, '')
      let n: number
      if (onlyDigits.includes(',') && onlyDigits.includes('.')) {
        const lc = onlyDigits.lastIndexOf(','); const ld = onlyDigits.lastIndexOf('.')
        const norm = (lc > ld) ? onlyDigits.replace(/\./g, '').replace(',', '.') : onlyDigits.replace(/,/g, '')
        n = Number(norm)
      } else if (onlyDigits.includes(',')) {
        const parts = onlyDigits.split(',');
        const norm = (parts.length === 2 && parts[1].length > 0 && parts[1].length <= 2) ? (parts[0].replace(/,/g,'') + '.' + parts[1]) : onlyDigits.replace(/,/g,'')
        n = Number(norm)
      } else {
        n = Number(onlyDigits)
      }
      return isFinite(n) ? n : 0
    })()
    const start = model.firstTransactionDate || model.firstInvoiceDate
    if (!start) return { error: 'No dated rows found.' }
    return analyze(model.invoices, model.payments, start, bb)
  }, [upload, beginBal])
  const result: any = loadedResult || computedResult

  useEffect(() => {
    if (!result) return
    if ((result as any) && (result as any).error) {
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

  // Track Firebase auth user (if enabled)
  useEffect(() => {
    if (!fbEnabled || !isFirebaseReady()) return
    try {
      const off = onUser(u => setUserEmail(u?.email || ''))
      return () => off()
    } catch {}
  }, [fbEnabled])

    function reviveReport(rep: any): any {
    try {
      const out: any = { ...rep }
      if (out.startDate) { const d = new Date(out.startDate); if (!isNaN(+d)) out.startDate = d }
      if (Array.isArray(out.invoices)) {
        out.invoices = out.invoices.map((inv: any) => {
          const ii = { ...inv }
          if (ii.invoiceDate) { const d = new Date(ii.invoiceDate); if (!isNaN(+d)) ii.invoiceDate = d }
          if (ii.closingDate) { const d = new Date(ii.closingDate); if (!isNaN(+d)) ii.closingDate = d }
          return ii
        })
      }
      if (Array.isArray(out.months)) {
        out.months = out.months.map((m: any) => ({ ...m, dt: new Date(m.dt) }))
      }
      return out
    } catch { return rep }
  }
async function loadLatest() {
    if (!customerKey) { setToast('No customer name (B1)'); setTimeout(()=>setToast(null), 1200); return }
    try {
      if (fbEnabled && isFirebaseReady()) {
        let remote = await fbLoad<any>(customerKey)
        if (!remote) throw new Error('No saved report found for this customer')
        if (typeof remote === 'string') { try { remote = JSON.parse(remote) } catch {} }
      const revived = reviveReport(remote)
      setLoadedResult(revived)
      try { const dbg = (globalThis as any).__arDebug || ((globalThis as any).__arDebug = {}); dbg.loadedReport = revived } catch {}
      console.log('[loadLatest] report (firebase):', revived)
      } else {
        let local = await idbGet<any>(customerKey.toUpperCase().trim())
        if (!local) throw new Error('No local saved report for this customer')
        if (typeof local === 'string') { try { local = JSON.parse(local) } catch {} }
      const revivedLocal = reviveReport(local)
      setLoadedResult(revivedLocal)
      try { const dbg = (globalThis as any).__arDebug || ((globalThis as any).__arDebug = {}); dbg.loadedReport = revivedLocal } catch {}
      console.log('[loadLatest] report (local):', revivedLocal)
      }
      setToast('Loaded latest report (see console)')
      setTimeout(()=>setToast(null), 1400)
    } catch (e:any) {
      setToast('Load error: ' + (e?.message || e))
      setTimeout(()=>setToast(null), 2200)
    }
  }
  const [customerList, setCustomerList] = useState<string[]>([])
  async function refreshCustomerList() {
    try {
      if (fbEnabled && isFirebaseReady()) {
        // For now, per request, list local IndexedDB keys until Firebase setup is complete
        const localKeys = await idbListKeys()
        setCustomerList(localKeys)
      } else {
        const keys = await idbListKeys()
        setCustomerList(keys)
      }
    } catch (e) {
      console.warn('customer list refresh failed:', e)
    }
  }
  useEffect(() => { refreshCustomerList() }, [])

  async function saveLatest() {
    if (!customerKey) { setToast('No customer name (B1)'); setTimeout(()=>setToast(null), 1200); return }
    if (!result || 'error' in result) { setToast('No analysis to save'); setTimeout(()=>setToast(null), 1200); return }
    try {
      if (fbEnabled && isFirebaseReady()) {
        await fbSave(customerKey, result)
      } else {
        await idbSet(customerKey.toUpperCase().trim(), result)
      }
      setToast('Saved latest report')
      setTimeout(()=>setToast(null), 1200)
    } catch (e:any) {
      setToast('Save error: ' + (e?.message || e))
      setTimeout(()=>setToast(null), 2200)
    }
  }

  const exportToExcel = () => {
    if (!result || 'error' in result) return
    const wb = XLSX.utils.book_new()

    const metricsRows = result.metrics.map((m: any) => ({ Metric: m.label, Value: m.value, Assessment: m.assess }))
    const wsMetrics = XLSX.utils.json_to_sheet(metricsRows)
    XLSX.utils.book_append_sheet(wb, wsMetrics, 'Metrics')

    const agingLabels = ['0-30 days','31-60 days','61-90 days','91+ days']
    const agingRows = agingLabels.map((lbl, i) => ({ Bucket: lbl, Outstanding: result.aging[i] }))
    const wsAging = XLSX.utils.json_to_sheet(agingRows)
    XLSX.utils.book_append_sheet(wb, wsAging, 'Aging')

    const analysisRows = result.invoices.map((inv: any) => {
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

    const trendRows = result.months.map((m: any) => ({ Month: new Date(m.dt).toLocaleDateString(), 'Avg Days to Pay': m.avg }))
    const wsTrend = XLSX.utils.json_to_sheet(trendRows)
    XLSX.utils.book_append_sheet(wb, wsTrend, 'Trend')

    const base = upload?.filename ? upload.filename.replace(/\.[^.]+$/, '') : 'export'
    XLSX.writeFile(wb, `${base}-ar-analysis.xlsx`)
  }

  // Delete saved report for current customer (IndexedDB) and clear loaded state
  async function deleteSavedForCurrent() {
    const key = customerKey?.toUpperCase().trim()
    if (!key) { setToast('No customer name (B1)'); setTimeout(()=>setToast(null), 1200); return }
    try {
      await idbDelete(key)
      setLoadedResult(null)
      await refreshCustomerList()
      setToast(`Deleted saved report for ${customerKey}`)
      setTimeout(()=>setToast(null), 1400)
    } catch (e:any) {
      setToast('Delete failed: ' + (e?.message || e))
      setTimeout(()=>setToast(null), 2200)
    }
  }

  // Delete all saved reports on this device (IndexedDB) with confirmation
  async function deleteAllData() {
    const ok = typeof window !== 'undefined' ? window.confirm('Delete ALL saved reports on this device? This cannot be undone.') : true
    if (!ok) return
    try {
      await idbClearAll()
      setLoadedResult(null)
      setCustomerKey('')
      await refreshCustomerList()
      setToast('All saved reports deleted on this device')
      setTimeout(()=>setToast(null), 1500)
    } catch (e:any) {
      setToast('Delete all failed: ' + (e?.message || e))
      setTimeout(()=>setToast(null), 2200)
    }
  }

  return (
    <div style={{ fontFamily: 'system-ui, sans-serif', padding: 16 }} dir={lang==='ar'?'rtl':'ltr'}>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center' }}>
        <h1 style={{ margin:0 }}>{t('title')}</h1>
        <div style={{ display:'flex', gap:8, alignItems:'center' }}>
          <label style={{ fontSize:12, color:'#555' }}>{t('language')}</label>
          <select value={lang} onChange={e=>setLang(e.target.value as any)} style={{ padding:'4px 6px', borderRadius:6 }}>
            <option value='en'>English</option>
            <option value='tr'>Türkçe</option>
            <option value='ar'>العربية</option>
  
</select>
        </div>
      </div>
      <p>{t('instructions')}</p>
      <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 12 }}>
        {fbEnabled && (
          userEmail
          ? <button onClick={() => { signOutUser().catch(()=>{}); }}>{`Sign out (${userEmail})`}</button>
          : <button onClick={() => { signInWithGoogle().then(u=>setUserEmail(u.email||'')); }}>{'Sign in with Google'}</button>
        )}
        <div style={{ opacity: 0.8 }}>Customer (B1): <b>{customerKey || '(not detected yet)'}</b></div>
        <div style={{ display:'flex', gap:6, alignItems:'center' }}>
          <label style={{ opacity:0.8 }}>Or pick saved:</label>
          <select value={customerKey} onChange={e=>setCustomerKey(e.target.value)} style={{ padding:6, borderRadius:6 }}>
            <option value="">-- Select customer --</option>
            {customerList.map(k => (
              <option key={k} value={k}>{k}</option>
            ))}
          </select>
          <button onClick={refreshCustomerList} title="Refresh customer list">Refresh</button>
        </div>
        <button onClick={loadLatest}>{fbEnabled ? 'Load Latest (Firebase)' : 'Load Latest (This Browser)'}</button>
        <button onClick={saveLatest} disabled={!result || ('error' in (result as any))}>{fbEnabled ? 'Save Latest (Firebase)' : 'Save Latest (This Browser)'}</button>
        <button onClick={deleteSavedForCurrent} title="Delete saved report for this customer">Delete Saved Report</button>
        <button onClick={deleteAllData} title="Delete all saved reports on this device">Delete All Data</button>
      </div>
      <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 12 }}>
        <label style={{ border: '1px solid #ccc', padding: '8px 12px', borderRadius: 6, cursor: 'pointer', background: '#fafafa' }}>
          <input type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={(e) => {
            const f = e.target.files?.[0]
            if (f) onFile(f)
          }} />
          {t('uploadFile')}
        </label>
        <span>{upload ? upload.filename : t('noFile')}</span>
        <div>|</div>
        <label>{t('beginBal')}: <input value={beginBal} onChange={e => { setBeginBal(e.target.value); setBeginBalAuto(false) }} style={{ width: 120 }} /></label>
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
          {showDebug ? t('hideDebug') : t('showDebug')}
        </button>
      </div>

      {result && (result as any)?.error && (
        <div style={{ color: 'crimson' }}>{result.error}</div>
      )}

      {showDebug && (
        <DebugPanel />
      )}

      {result && !(result as any)?.error && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 24 }}>
          <div>
            <h2>{t('metrics')}</h2>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>{t('metric')}</th>
                  <th style={{ textAlign: 'right', borderBottom: '1px solid #ddd', padding: 6 }}>{t('value')}</th>
                  <th style={{ textAlign: 'center', borderBottom: '1px solid #ddd', padding: 6 }}>{t('assessment')}</th>
                </tr>
              </thead>
              <tbody>
                {(() => {
                  const specialLabels = new Set(['Customer Risk Rating','Average Monthly Purchases (TRY)','Credit Limit (TRY)'])
                  const metricsMain = result.metrics.filter((m: any) => !specialLabels.has(m.label))
                  return metricsMain.map((m: any, i: number) => {
                    const isPct = m.label.includes('%')
                    const fmt = (v: any) => {
                      if (v === '') return ''
                      if (isPct) return (v as number).toLocaleString(undefined, { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 })
                      if (typeof v === 'number') return v.toLocaleString(undefined, { maximumFractionDigits: 0 })
                      return String(v)
                    }
                    const color = m.assess === 'Good' ? '#c6efce' : m.assess === 'Average' ? '#ffe6cc' : m.assess === 'Poor' ? '#f4a7a7' : 'transparent'
                    const labelLocal = (metricNamesAll as any)[m.label] ? (metricNamesAll as any)[m.label][lang] : m.label
                    const assessLocal = assessNames[lang][m.assess] ?? m.assess
                    const note = metricNotes[m.label] || ''
                    return (
                      <tr key={i}>
                        <td title={note || undefined} style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{labelLocal}</td>
                        <td title={note || undefined} style={{ padding: 6, borderBottom: '1px solid #f0f0f0', textAlign: 'right' }}>{fmt(m.value)}</td>
                        <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0', textAlign: 'center', background: color }}>{assessLocal}</td>
                      </tr>
                    )
                  })
                })()}
              </tbody>
            </table>
          </div>

          <div>
            <h2>Key Account Metrics</h2>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>{t('metric')}</th>
                  <th style={{ textAlign: 'right', borderBottom: '1px solid #ddd', padding: 6 }}>{t('value')}</th>
                </tr>
              </thead>
              <tbody>
                {(() => {
                  const specialLabels = new Set(['Customer Risk Rating','Average Monthly Purchases (TRY)','Credit Limit (TRY)'])
                  const items = result.metrics.filter((m: any) => specialLabels.has(m.label))
                  const cl = items.find((it: any) => it.label === 'Credit Limit (TRY)')?.value
                  const openBal = result.invoices.length ? (result.invoices[result.invoices.length - 1].running ?? 0) : 0
                  const available: any = (typeof cl === 'number') ? Math.max(0, cl - openBal) : ''
                  const rows = items.map((m: any, i: number) => {
                    const isPct = m.label.includes('%')
                    const fmt = (v: any) => {
                      if (v === '') return ''
                      if (isPct) return (v as number).toLocaleString(undefined, { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 })
                      if (typeof v === 'number') return v.toLocaleString(undefined, { maximumFractionDigits: 0 })
                      return String(v)
                    }
                    const labelLocal = (metricNamesAll as any)[m.label] ? (metricNamesAll as any)[m.label][lang] : m.label
                    const note = metricNotes[m.label] || ''
                    const riskColor = m.label === 'Customer Risk Rating'
                      ? (m.assess === 'Good' ? '#c6efce' : m.assess === 'Average' ? '#ffe6cc' : m.assess === 'Poor' ? '#f4a7a7' : 'transparent')
                      : 'transparent'
                    return (
                      <tr key={i}>
                        <td title={note || undefined} style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{labelLocal}</td>
                        <td title={note || undefined} style={{ padding: 6, borderBottom: '1px solid #f0f0f0', textAlign: 'right', background: riskColor }}>{fmt(m.value)}</td>
                      </tr>
                    )
                  })
                  rows.push(
                    <tr key={'available-credit'}>
                      <td title={metricNotes['Available Credit (TRY)']} style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{(metricNamesAll as any)['Available Credit (TRY)'][lang]}</td>
                      <td title={metricNotes['Available Credit (TRY)']} style={{ padding: 6, borderBottom: '1px solid #f0f0f0', textAlign: 'right' }}>{
                        (typeof available === 'number') ? available.toLocaleString(undefined, { maximumFractionDigits: 0 }) : ''
                      }</td>
                    </tr>
                  )
                  return rows
                })()}
              </tbody>
            </table>
          </div>

          <div>
            <h2>{t('aging')}</h2>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <th style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>{t('bucket')}</th>
                  <th style={{ textAlign: 'right', borderBottom: '1px solid #ddd', padding: 6 }}>{t('outstanding')}</th>
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
            <details>
              <summary style={{ cursor: 'pointer', userSelect: 'none', fontWeight: 600, fontSize: 20, marginBottom: 8 }}>{t('analysisTable')}</summary>
              <div style={{ marginTop: 8 }}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  {[t('invoiceDate'),t('invoiceNo'),t('type'),t('amount'),t('closingDate'),t('termDays'),t('dueDate'),t('daysToPay'),t('daysAfterDue'),t('remaining'),t('arBalance')].map(h => (
                    <th key={h} style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {result.invoices.map((inv: any, i: number) => {
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
            </details>
          </div>

          <div style={{ gridColumn: '1 / span 2' }}>
            <details>
              <summary style={{ cursor: 'pointer', userSelect: 'none', fontWeight: 600, fontSize: 20, marginBottom: 8 }}>{t('ledger')}</summary>
              <div style={{ marginTop: 8 }}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  {[t('date'),t('type'),t('description'),t('ref'),t('debit'),t('credit'),t('running')].map(h => (
                    <th key={h} style={{ textAlign: 'left', borderBottom: '1px solid #ddd', padding: 6 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {(result.ledger || []).map((e: any, i: number) => (
                  <tr key={i}>
                    <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{new Date(e.date).toLocaleDateString(locale)}</td>
                    <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{e.kind}</td>
                    <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{e.description}</td>
                    <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0' }}>{e.ref || ''}</td>
                    <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0', textAlign: 'right' }}>{e.debit ? e.debit.toLocaleString(locale) : ''}</td>
                    <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0', textAlign: 'right' }}>{e.credit ? e.credit.toLocaleString(locale) : ''}</td>
                    <td style={{ padding: 6, borderBottom: '1px solid #f0f0f0', textAlign: 'right' }}>{e.balance.toLocaleString(locale)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
              </div>
            </details>
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
  const headerRowIdx: number | undefined = dbg.headerRowIdx
  const descColsHeaders: string[] = dbg.descColsHeaders || []
  const cPayTpIndex: number | undefined = dbg.cPayTpIndex
  const cPayTpIndexAuto: number | undefined = dbg.cPayTpIndexAuto
  const cMaturityIndex: number | undefined = dbg.cMaturityIndex
  const cDateIndex: number | undefined = dbg.cDateIndex
  const checkInspect: any[] = dbg.checkInspect || []
  const checkNoMatExamples: any[] = dbg.checkNoMatExamples || []
  const counts = dbg.checkCounts || { total: checkInspect.length, withMaturity: checkInspect.filter((x:any)=>!!x.maturity).length, withoutMaturity: checkInspect.filter((x:any)=>!x.maturity).length }
  const reconcile = dbg.reconcile || {}
  const payTypes: Record<string,number> = dbg.payTypes || {}
  const termCounts: Record<string,number> = dbg.invoiceTermCounts || {}
  const openingRowIndex: number | undefined = dbg.openingRowIndex

  return (
    <div style={{ border: '1px solid #ddd', borderRadius: 8, padding: 12, background: '#fffef8', marginBottom: 16 }}>
      <div style={{ fontWeight: 600, marginBottom: 8 }}>Debug Summary</div>
      <div style={{ display: 'grid', gridTemplateColumns: '220px 1fr', rowGap: 6, columnGap: 8 }}>
        <div>Header row index:</div>
        <div>{typeof headerRowIdx === 'number' ? String(headerRowIdx) : 'n/a'}</div>
        <div>Headers (lower):</div>
        <div>{headersLower.join(', ') || 'n/a'}</div>
        <div>Description columns:</div>
        <div>{descColsHeaders.join(', ') || 'n/a'}</div>
        <div>Date column index:</div>
        <div>{typeof cDateIndex === 'number' ? String(cDateIndex) : 'n/a'}</div>
        <div>Pay Type column index:</div>
        <div>{typeof cPayTpIndex === 'number' ? String(cPayTpIndex) : 'n/a'}{typeof cPayTpIndexAuto === 'number' ? ` (auto ${cPayTpIndexAuto})` : ''}</div>
        <div>Maturity column index:</div>
        <div>{typeof cMaturityIndex === 'number' ? String(cMaturityIndex) : 'n/a'}</div>
        <div>Opening balance row index:</div>
        <div>{typeof openingRowIndex === 'number' ? String(openingRowIndex) : 'n/a'}</div>
        <div>Checks parsed (with/without maturity):</div>
        <div>{counts.withMaturity} / {counts.withoutMaturity} (total {counts.total})</div>
        <div>Reconcile (expected vs computed):</div>
        <div>
          {typeof reconcile.expectedOutstanding === 'number' ? reconcile.expectedOutstanding.toLocaleString() : 'n/a'}
          {' '}vs{' '}
          {typeof reconcile.computedOutstanding === 'number' ? reconcile.computedOutstanding.toLocaleString() : 'n/a'}
          {typeof reconcile.delta === 'number' ? ` (delta ${reconcile.delta.toLocaleString()})` : ''}
        </div>
        <div>Pay types:</div>
        <div>{Object.keys(payTypes).length ? Object.entries(payTypes).map(([k,v])=>`${k||'∅'}=${v}`).join(', ') : 'n/a'}</div>
        <div>Invoice term counts:</div>
        <div>{Object.keys(termCounts).length ? Object.entries(termCounts).map(([k,v])=>`${k}d=${v}`).join(', ') : 'n/a'}</div>
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



