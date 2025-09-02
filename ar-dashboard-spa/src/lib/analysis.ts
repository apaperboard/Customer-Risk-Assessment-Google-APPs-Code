import { amountToNumber, extractDateFromText, mode, normalizePayType, parseDMY } from './parsers'

export type Invoice = {
  invoiceDate: Date
  invoiceNum: string
  type: string
  amount: number
  remaining: number
  running?: number
  term: number
  paid: boolean
  closingDate: Date | null
  _synthetic?: boolean
  _appliedTerms?: number[]
  _appliedChecks?: { amount: number; invoiceDate: Date; paymentDate: Date; maturityDate: Date | null }[]
}

export type Payment = {
  paymentDate: Date
  amount: number
  maturityDate: Date | null
  payType: string
  expectedTerm: number | null
  desc?: string
}

export type ParsedInput = {
  invoices: Invoice[]
  payments: Payment[]
  firstInvoiceDate: Date | null
  firstTransactionDate: Date | null
}

type RowObject = Record<string, any>

export function findCol(headers: string[], names: string[]): number {
  for (const n of names) {
    const idx = headers.indexOf(n.toLowerCase())
    if (idx >= 0) return idx + 1
  }
  return -1
}

function findByIncludes(headers: string[], names: string[]): number {
  const lower = headers.map(h => String(h ?? '').toLowerCase())
  for (let i = 0; i < lower.length; i++) {
    const h = lower[i]
    if (!h) continue
    for (const n of names) { const nn = String(n ?? '').toLowerCase(); if (nn && h.includes(nn)) return i + 1 }
  }
  return -1
}

function findAllByIncludes(headers: string[], names: string[]): number[] {
  const lower = headers.map(h => String(h ?? '').toLowerCase())
  const out: number[] = []
  for (let i = 0; i < lower.length; i++) {
    const h = lower[i]
    if (!h) continue
    for (const n of names) { const nn = String(n ?? '').toLowerCase(); if (nn && h.includes(nn)) { out.push(i + 1); break } }
  }
  return out
}

export function parseRowsToModel(rows: RowObject[]): ParsedInput {
  if (!rows.length) return { invoices: [], payments: [], firstInvoiceDate: null, firstTransactionDate: null }
  const origHeaders = Object.keys(rows[0])
  const headers = origHeaders.map(h => String(h || '').trim().toLowerCase())
  const keyMap = new Map<string, string>()
  for (let i = 0; i < headers.length; i++) {
    keyMap.set(headers[i], origHeaders[i])
  }
  // Header aliases in EN/TR/AR (lowercased)
  const creditNames = [
    'credit','alacak','alacaklar','invoice','fatura','fatura tutari','fatura miktari',
    'دائن','فاتورة','قيمة الفاتورة'
  ]
  const debitNames  = [
    'debit','borç','borclar','payment','ödeme','tahsilat','odeme',
    'مدين','دفعة','سداد','تحصيل'
  ]
  const descNames   = [
    'description','açiklama','aciklama','desc','not','memo',
    'الوصف','شرح','بيان'
  ]
  const dateNames   = [
    'date','tarih',
    'تاريخ','التاريخ'
  ]
  const payTypeNames= [
    'pay type','payment type','odeme tipi','ödeme tipi','odeme turu','ödeme türü','tahsilat tipi','paytype',
    'نوع الدفع','طريقة الدفع','نوع السداد'
  ]

  let cCredit = findCol(headers, creditNames)
  let cDebit  = findCol(headers, debitNames)
  const cDesc   = findCol(headers, descNames)
  let cDate   = findCol(headers, dateNames)
  const cPayTp  = findCol(headers, payTypeNames)
  // Fallbacks via substring includes, including Arabic tokens
  if (cCredit <= 0) cCredit = findByIncludes(headers, ['credit','alacak','invoice','fatura','دائن'])
  if (cDebit  <= 0) cDebit  = findByIncludes(headers, ['debit','borç','borc','payment','ödeme','odeme','tahsilat','مدين'])
  if (cDate   <= 0) cDate   = findByIncludes(headers, ['date','tarih','التاريخ'])
  // Heuristic: if pay type column not found, or looks wrong, auto-detect by cell values
  function looksLikePayType(val: any): boolean {
    const s = String(val ?? '').toLowerCase()
    if (!s) return false
    return /(\bkk\b|k\.k\.|kredi\s*kart|kart|credit\s*card|visa|master|peşin|pesin|nakit|cash|çek|cek|cheque|check|senet)/i.test(s)
  }
  function detectPayTypeColumn(): number {
    const n = headers.length
    let bestIdx = -1, bestScore = 0, bestRatio = 0
    const sampleCount = Math.min(rows.length, 500)
    for (let i = 1; i <= n; i++) {
      let total = 0, match = 0, numericLike = 0
      for (let r = 0; r < sampleCount; r++) {
        const v = get(rows[r] as any, i)
        if (v == null || String(v).trim() === '') continue
        total++
        if (typeof v === 'number') { numericLike++ }
        if (looksLikePayType(v)) match++
      }
      if (total === 0) continue
      if (numericLike/Math.max(total,1) > 0.7) continue
      const ratio = match/total
      if (match > bestScore || (match === bestScore && ratio > bestRatio)) { bestIdx = i; bestScore = match; bestRatio = ratio }
    }
    if (bestScore >= 3 && bestRatio >= 0.10) return bestIdx
    return -1
  }
  let cPayTpEff = cPayTp
  function payTypeScoreFor(col: number): {match:number, total:number} {
    if (col <= 0) return { match: 0, total: 0 }
    let match = 0, total = 0
    const sampleCount = Math.min(rows.length, 300)
    for (let r = 0; r < sampleCount; r++) {
      const v = get(rows[r] as any, col)
      if (v == null || String(v).trim() === '') continue
      total++
      if (looksLikePayType(v)) match++
    }
    return { match, total }
  }
  const cur = payTypeScoreFor(cPayTp)
  if (cPayTp <= 0 || cur.match < 3 || (cur.total > 0 && (cur.match/cur.total) < 0.10)) {
    const autoCol = detectPayTypeColumn()
    if (autoCol > 0) {
      try { const dbg = (globalThis as any).__arDebug || ((globalThis as any).__arDebug = {}); dbg.cPayTpIndexAuto = autoCol } catch {}
      cPayTpEff = autoCol
    }
  }
  // Additional helpers via substring match (handles mixed-language headers and composite labels)
  const cMaturity = findByIncludes(headers, ['vade', 'vade tarihi', 'maturity', 'maturity date', 'due date', 'son ödeme', 'son odeme', 'vadesi'])
  const descCols  = Array.from(new Set([
    ...findAllByIncludes(headers, ['description','açıklama','aciklama','desc','not','memo','a. açıklama'])
  ]))

  try {
    console.debug('[parseRowsToModel] header mapping:', { headers, cCredit, cDebit, cDesc, cDate, cPayTp, cMaturity })
    const dbg = (globalThis as any).__arDebug || ((globalThis as any).__arDebug = {})
    dbg.cPayTpIndex = cPayTp
    dbg.cMaturityIndex = cMaturity
    dbg.descColsHeaders = ((): string[] => {
      const out: string[] = []
      const uniq = new Set<number>(descCols.length ? descCols : (cDesc > 0 ? [cDesc] : []))
      for (const ci of uniq) out.push(headers[ci-1])
      return out
    })()
  } catch {}

  function get(row: RowObject, col: number): any {
    if (col <= 0) return ''
    const keyLower = headers[col-1]
    const orig = keyMap.get(keyLower) || keyLower
    return row[orig]
  }

  // If date column not found via headers, auto-detect by scanning cells for parseable dates
  let cDateEff = cDate
  if (cDateEff <= 0) {
    const sampleCount = Math.min(rows.length, 400)
    let bestIdx = -1, bestHits = 0, bestRatio = 0
    for (let i = 1; i <= headers.length; i++) {
      let total = 0, hits = 0
      for (let r = 0; r < sampleCount; r++) {
        const v = get(rows[r] as any, i)
        if (v == null || String(v).trim() === '') continue
        total++
        const d = parseDMY(v)
        if (d instanceof Date && !isNaN(+d)) hits++
      }
      if (total === 0) continue
      const ratio = hits / total
      if (hits > bestHits || (hits === bestHits && ratio > bestRatio)) { bestIdx = i; bestHits = hits; bestRatio = ratio }
    }
    if (bestHits >= 3 && bestRatio >= 0.15) cDateEff = bestIdx
  }
  try { const dbg = (globalThis as any).__arDebug || ((globalThis as any).__arDebug = {}); dbg.cDateIndex = cDateEff } catch {}

  const invoices: Invoice[] = []
  const payments: Payment[] = []
  let firstInvoiceDate: Date | null = null
  let firstTransactionDate: Date | null = null

  // Decide invoice vs payment mapping between debit/credit.
  const headerJoined = headers.join(' ')
  const hasBorc = /(borç|borc)/i.test(headerJoined)
  const hasAlacak = /alacak/i.test(headerJoined)
  // Baseline mapping: debit (borç/مدين/debit) are invoices; credit are payments
  let invoiceFromDebit = true
  if (hasBorc && hasAlacak) {
    invoiceFromDebit = true
  } else if (cCredit > 0 && cDebit > 0) {
    let debPos = 0, credPos = 0
    for (const r of rows) {
      const deb = amountToNumber(get(r as any, cDebit))
      const cre = amountToNumber(get(r as any, cCredit))
      if (isFinite(deb) && deb > 0) debPos++
      if (isFinite(cre) && cre > 0) credPos++
    }
    // If credit positives dominate significantly, flip mapping
    if (credPos > debPos * 1.1) invoiceFromDebit = false
    try { console.debug('[parseRowsToModel] debit vs credit positives (baseline debit=invoices):', { debPos, credPos, invoiceFromDebit }) } catch {}
  }

  for (const row of rows) {
    const credit = amountToNumber(get(row, cCredit))
    const debit  = amountToNumber(get(row, cDebit))
    const desc   = ((): any => {
      for (const c of (descCols.length ? descCols : [cDesc])) {
        const v = get(row, c)
        if (v != null && String(v).trim() !== '') return v
      }
      return get(row, cDesc)
    })()
    const date   = parseDMY(get(row, cDateEff)) || extractDateFromText(desc)

    const empty = [credit, debit, desc, date].every(v => v == null || v === '' || (typeof v === 'number' && isNaN(v)))
    if (empty) continue

    const isInvoice = invoiceFromDebit ? (isFinite(debit) && debit > 0) : (isFinite(credit) && credit > 0)
    const isPayment = invoiceFromDebit ? (isFinite(credit) && credit > 0) : (isFinite(debit) && debit > 0)

    if (isInvoice) {
      if (!date) continue
      const invNumMatch = String(desc || '').match(/(No\s*\S+|\b[A-Z0-9\-]{6,}\b)/)
      invoices.push({
        invoiceDate: date,
        invoiceNum: invNumMatch ? invNumMatch[0] : '',
        type: 'Invoice',
        amount: invoiceFromDebit ? debit : credit,
        remaining: invoiceFromDebit ? debit : credit,
        term: 30,
        paid: false,
        closingDate: null,
      })
      if (!firstInvoiceDate || +date < +firstInvoiceDate) firstInvoiceDate = date
      if (!firstTransactionDate || +date < +firstTransactionDate) firstTransactionDate = date
    }

    if (isPayment) {
      if (!date) continue
      const descAll = ((): string => {
        const parts: string[] = []
        for (const c of (descCols.length ? descCols : [cDesc])) {
          const v = get(row, c)
          if (v != null && String(v).trim() !== '') parts.push(String(v))
        }
        if (parts.length === 0) parts.push(String(desc ?? ''))
        return parts.join(' | ')
      })()
      const payTypeRaw = (() => { const pt = get(row, cPayTpEff); return pt ? (String(pt) + ' | ' + descAll) : descAll })()
      const norm = normalizePayType(payTypeRaw)
      const descDate = extractDateFromText(descAll)
      const maturity = parseDMY(get(row, cMaturity)) || descDate || null
      // Business rule: Only treat as Check if there is a date in description
      const effectiveType = (norm.type === 'Check' && !descDate) ? '' : norm.type
      payments.push({
        paymentDate: date,
        amount: invoiceFromDebit ? credit : debit,
        maturityDate: maturity,
        payType: effectiveType,
        expectedTerm: norm.termDays,
        desc: descAll,
      })
      try {
        const dbg = (globalThis as any).__arDebug || ((globalThis as any).__arDebug = {})
        if (effectiveType === 'Check') {
          const arr1 = (dbg.checkInspect ||= [])
          if (arr1.length < 10) arr1.push({ desc: descAll, maturity, payTypeRaw })
          if (!maturity) {
            const arr2 = (dbg.checkNoMatExamples ||= [])
            if (arr2.length < 10) arr2.push({ date, payTypeRaw, desc: descAll })
          }
          const all = (dbg.checkAll ||= [])
          all.push({ date, maturity, amount: invoiceFromDebit ? credit : debit, desc: descAll, payTypeRaw })
        } else if (descDate) {
          const cand = (dbg.checkCandidatesNotClassified ||= [])
          cand.push({ date, desc: descAll, payTypeRaw })
        }
        const pts = (dbg.payTypeSamples ||= [])
        if (pts.length < 15) pts.push({ payTypeRaw, norm: effectiveType })
      } catch {}
      if (!firstTransactionDate || +date < +firstTransactionDate) firstTransactionDate = date

    }
  }

  // Infer invoice term
  const paymentTerms = payments.map(p => p.expectedTerm).filter(x => x != null) as number[]
  let inferredTerm: number
  if (paymentTerms.length) {
    inferredTerm = mode(paymentTerms, 30)
  } else {
    const deltas: number[] = []
    for (const p of payments) {
      if (!p.maturityDate) continue
      const pd = Math.round((+p.maturityDate - +p.paymentDate)/86400000)
      if (isFinite(pd) && pd > 0) deltas.push(pd)
    }
    inferredTerm = mode(deltas.map(x => {
      const choices = [30,60,90]
      let best = 30, diffBest = 1e9
      for (const c of choices) {
        const d = Math.abs(x - c)
        if (d < diffBest) { diffBest = d; best = c }
      }
      return best
    }), 30)
  }
  for (const inv of invoices) inv.term = inferredTerm

  invoices.sort((a,b) => +a.invoiceDate - +b.invoiceDate)
  payments.sort((a,b) => +a.paymentDate - +b.paymentDate)
  try {
    console.debug('[parseRowsToModel] summary:', { invoices: invoices.length, payments: payments.length, firstInvoiceDate, firstTransactionDate })
  } catch {}
  return { invoices, payments, firstInvoiceDate, firstTransactionDate }
}

export function analyze(invoicesIn: Invoice[], paymentsIn: Payment[], startDate: Date, beginningBalance: number) {
  const invoices = invoicesIn.map(x => ({...x}))
  const payments = paymentsIn.map(x => ({...x}))
  const today = new Date()

  if (beginningBalance > 0 && startDate) {
    invoices.unshift({
      invoiceDate: new Date(+startDate - 86400000),
      invoiceNum: 'BEGIN BAL',
      type: 'Opening',
      amount: beginningBalance,
      remaining: beginningBalance,
      term: 30,
      paid: false,
      closingDate: null,
      _synthetic: true,
    })
  }

  // FIFO + tracking
  let allLagTotalAmt = 0, allLagWeightedSum = 0
  const unapplied: { date: Date; amount: number; remaining: number; reason: string }[] = []
  const advances: { date: Date; remaining: number; payType: string; maturityDate: Date | null; expectedTerm: number | null }[] = []
  for (const p of payments) {
    let rem = p.amount
    for (const inv of invoices) {
      if (rem <= 0) break
      if (inv.paid) continue
      if (+inv.invoiceDate > +p.paymentDate) continue
      const applied = Math.min(inv.remaining, rem)
      if (applied <= 0) continue
      inv.remaining -= applied
      rem -= applied
      const lagAll = Math.max(0, Math.round((+p.paymentDate - +inv.invoiceDate)/86400000))
      allLagTotalAmt += applied
      allLagWeightedSum += lagAll * applied
      if (p.expectedTerm != null) {
        (inv._appliedTerms ||= []).push(p.expectedTerm)
      }
      if (p.payType === 'Check') {
        (inv._appliedChecks ||= []).push({ amount: applied, invoiceDate: inv.invoiceDate, paymentDate: p.paymentDate, maturityDate: p.maturityDate || null })
      }
      if (inv.remaining === 0) { inv.paid = true; inv.closingDate = p.paymentDate }
    }
    if (rem > 0) {
      const reason = invoices.every(inv => +inv.invoiceDate > +p.paymentDate) ? 'before_first_invoice' : 'overpayment_or_future_invoice'
      unapplied.push({ date: p.paymentDate, amount: p.amount, remaining: rem, reason })
      advances.push({ date: p.paymentDate, remaining: rem, payType: p.payType, maturityDate: p.maturityDate, expectedTerm: p.expectedTerm })
    }
  }

  // Apply advances to future invoices (carry-over prepayments)
  advances.sort((a,b) => +a.date - +b.date)
  for (const adv of advances) {
    let rem = adv.remaining
    for (const inv of invoices) {
      if (rem <= 0) break
      if (+inv.invoiceDate <= +adv.date) continue
      if (inv.remaining <= 0) continue
      const applied = Math.min(inv.remaining, rem)
      if (applied <= 0) continue
      inv.remaining -= applied
      rem -= applied
      // advance has zero lag contribution (payment precedes invoice)
      allLagTotalAmt += applied
      // propagate expected term from advance to the invoice (was missing)
      if (adv.expectedTerm != null) {
        (inv._appliedTerms ||= []).push(adv.expectedTerm)
      }
      // track check-specific if advance came from a check
      if (adv.payType === 'Check') {
        (inv._appliedChecks ||= []).push({ amount: applied, invoiceDate: inv.invoiceDate, paymentDate: adv.date, maturityDate: adv.maturityDate || null })
      }
      if (inv.remaining === 0) { inv.paid = true; inv.closingDate = inv.invoiceDate }
    }
    // update leftover for debug
    adv.remaining = rem
  }
  for (const inv of invoices) {
    if (inv._synthetic) continue
    if (inv._appliedTerms && inv._appliedTerms.length) inv.term = mode(inv._appliedTerms, inv.term)
  }
  try {
    const invTermCounts: Record<string, number> = {}
    for (const inv of invoices) { const k = String(inv.term); invTermCounts[k] = (invTermCounts[k] || 0) + 1 }
    const dbg = (globalThis as any).__arDebug || ((globalThis as any).__arDebug = {})
    dbg.invoiceTermCounts = invTermCounts
  } catch {}

  const displayInvoices = invoices.filter(inv => !inv._synthetic)
  const paid = displayInvoices.filter(inv => inv.paid && inv.closingDate)
  const unpaid = displayInvoices.filter(inv => inv.remaining > 0)
  const overdueUnpaidByHandover = unpaid.filter(inv => ((+today - +inv.invoiceDate)/86400000) > 30)

  const avgPaymentLagDays = (allLagTotalAmt > 0) ? (allLagWeightedSum / allLagTotalAmt) : ''
  const sumAgeWeighted = unpaid.reduce((s,inv) => s + ((+today - +inv.invoiceDate)/86400000) * inv.remaining, 0)
  const sumRemaining = unpaid.reduce((s,inv) => s + inv.remaining, 0)
  const avgAgeUnpaid = (sumRemaining > 0) ? (sumAgeWeighted / sumRemaining) : ''
  const overdueRate = unpaid.length ? (overdueUnpaidByHandover.length/unpaid.length) : ''
  const blendedDaysToPay = displayInvoices.length ? displayInvoices.reduce((s,inv) => {
    const end = (inv.paid && inv.closingDate) ? inv.closingDate! : today
    return s + ((+end - +inv.invoiceDate)/86400000)
  }, 0)/displayInvoices.length : ''

  // Check-specific
  let checkTotalAmt = 0, checkLagWeightedSum = 0, checkWithin30Amt = 0
  let checkMatTotalAmt = 0, checkMatWeightedSum = 0
  for (const inv of invoices) {
    if (!inv._appliedChecks) continue
    for (const app of inv._appliedChecks) {
      const lagDays = Math.max(0, Math.round((+app.paymentDate - +app.invoiceDate)/86400000))
      checkTotalAmt += app.amount
      checkLagWeightedSum += lagDays * app.amount
      if (lagDays <= 30) checkWithin30Amt += app.amount
      if (app.maturityDate) {
        const matDur = Math.round((+app.maturityDate - +app.invoiceDate)/86400000)
        if (isFinite(matDur) && matDur > 0) { checkMatTotalAmt += app.amount; checkMatWeightedSum += matDur * app.amount }
      }
    }
  }
  const avgCheckHandoverLag = (checkTotalAmt > 0) ? (checkLagWeightedSum / checkTotalAmt) : ''
  const pctChecksHandedOver30 = (checkTotalAmt > 0) ? ((checkTotalAmt - checkWithin30Amt)/checkTotalAmt) : ''
  const avgCheckMaturityDuration = (checkMatTotalAmt > 0) ? (checkMatWeightedSum / checkMatTotalAmt) : ''
  const avgCheckMaturityOverBy = (avgCheckMaturityDuration !== '') ? ((avgCheckMaturityDuration as number) - 90) : ''

  // Avg Monthly Purchases (from analysis period start)
  const totalInvoicedInPeriod = displayInvoices.reduce((s, inv) => s + inv.amount, 0)
  const monthsInPeriod = startDate ? ((+today - +startDate) / (86400000 * 30.44)) : 0
  const avgMonthlyPurchases: number | '' = (monthsInPeriod > 0) ? (totalInvoicedInPeriod / monthsInPeriod) : ''

  // Score (weighted)
  function compLowerBetter(val: any, goodMax: number, avgMax: number) {
    if (val === '') return null
    return (val <= goodMax) ? 1 : (val <= avgMax) ? 0.5 : 0
  }
  let weightedSum = 0, weightTotal = 0
  function add(comp: number | null, w: number) { if (comp == null) return; weightedSum += comp*w; weightTotal += w }
  add(compLowerBetter(avgPaymentLagDays, 30, 45), 0.20)
  add(compLowerBetter(avgAgeUnpaid, 10, 20), 0.10)
  add(compLowerBetter(overdueRate, 0.10, 0.30), 0.10)
  add(compLowerBetter(blendedDaysToPay, 20, 35), 0.20)
  add(compLowerBetter(avgCheckMaturityOverBy, 30, 45), 0.20)
  add(compLowerBetter(pctChecksHandedOver30, 0.30, 0.60), 0.20)
  const normalizedScore = (weightTotal > 0) ? (weightedSum / weightTotal) : 0
  const riskBand = (normalizedScore <= 0.3333) ? 'Poor' : (normalizedScore <= 0.6667) ? 'Average' : 'Good'

  // Estimate customer credit limit based on risk band and most common term
  const maturitySamples: { days: number; expected: number }[] = []
  for (const p of payments) {
    if (p.maturityDate) {
      const d = Math.round((+p.maturityDate - +p.paymentDate) / 86400000)
      if (isFinite(d) && d > 0) maturitySamples.push({ days: d, expected: (p.expectedTerm != null ? p.expectedTerm : 30) })
    }
  }
  const mostCommonTerm = maturitySamples.length ? mode(maturitySamples.map(m => m.expected), 30) : 30
  const baseMult = (mostCommonTerm === 90)
    ? (riskBand === 'Good' ? 3.0 : (riskBand === 'Average' ? 2.75 : 2.5))
    : (riskBand === 'Good' ? 2.0 : (riskBand === 'Average' ? 1.5 : 1.0))
  let creditLimit: number | '' = (avgMonthlyPurchases !== '') ? ((avgMonthlyPurchases as number) * baseMult) : ''
  if (creditLimit !== '' && isFinite(creditLimit as number)) {
    // Round up to nearest 10,000
    creditLimit = Math.ceil((creditLimit as number) / 10000) * 10000
  }

  // Checks-only: % of checks where maturity duration exceeds expected term
  const pctChecksOverTerm: number | '' = maturitySamples.length
    ? (maturitySamples.filter(m => m.days > m.expected).length / maturitySamples.length)
    : ''

  // All payment types: % of invoices delivered (closed) after their term using handover lag
  const pctPaymentsDeliveredAfterTerm: number | '' = paid.length
    ? (paid.filter(inv => {
        const d2p = Math.round(((+inv.closingDate!) - (+inv.invoiceDate)) / 86400000)
        return d2p > inv.term
      }).length / paid.length)
    : ''

  // Metrics rows
  function assessLower(val: any, goodMax: number, avgMax: number) {
    if (val === '') return ''
    return (val <= goodMax) ? 'Good' : (val <= avgMax) ? 'Average' : 'Poor'
  }
  const metrics: { label: string; value: any; assess: string }[] = []
  const roundDays = (x: any) => (typeof x === 'number' ? Math.round(x) : x)
  metrics.push({ label: 'Average Days to Pay (Paid Only)', value: roundDays(avgPaymentLagDays), assess: assessLower(avgPaymentLagDays, 20, 40) })
  metrics.push({ label: 'Weighted Avg Age of Unpaid Invoices (Days)', value: roundDays(avgAgeUnpaid), assess: assessLower(avgAgeUnpaid, 10, 20) })
  metrics.push({ label: '% of Unpaid Invoices Overdue', value: overdueRate, assess: assessLower(overdueRate, 0.10, 0.30) })
  metrics.push({ label: 'Average Check Maturity Duration (Days)', value: roundDays(avgCheckMaturityDuration), assess: '' })
  metrics.push({ label: 'Avg Maturity Over By (Days)', value: roundDays(avgCheckMaturityOverBy), assess: assessLower(avgCheckMaturityOverBy, 0, 30) })
  metrics.push({ label: '% of Checks Over Term', value: pctChecksOverTerm, assess: assessLower(pctChecksOverTerm, 0.30, 0.60) })
  metrics.push({ label: '% of Payments Delivered After Term', value: pctPaymentsDeliveredAfterTerm, assess: assessLower(pctPaymentsDeliveredAfterTerm, 0.30, 0.60) })
  metrics.push({ label: 'Customer Risk Rating', value: riskBand, assess: riskBand })
  // New metric: overdue balance as a percentage of assigned credit limit (term-based overdue)
  const overdueOutstandingTerm = unpaid.reduce((s, inv) => {
    const ageDays = Math.floor(((+today - +inv.invoiceDate) / 86400000))
    return s + ((ageDays > inv.term) ? inv.remaining : 0)
  }, 0)
  const pctOverdueVsCreditLimit: number | '' = (creditLimit !== '' && (creditLimit as number) > 0)
    ? (overdueOutstandingTerm / (creditLimit as number))
    : ''
  metrics.push({ label: 'Overdue Balance as % of Credit Limit', value: pctOverdueVsCreditLimit, assess: assessLower(pctOverdueVsCreditLimit, 0.30, 0.60) })
  // Include purchases and credit limit like original Apps Script
  metrics.push({ label: 'Average Monthly Purchases (TRY)', value: avgMonthlyPurchases, assess: '' })
  metrics.push({ label: 'Credit Limit (TRY)', value: creditLimit, assess: '' })

  // Aging buckets
  const aging = [0,0,0,0]
  for (const inv of unpaid) {
    const age = Math.floor(((+today - +inv.invoiceDate)/86400000))
    if (age <= 30) aging[0] += inv.remaining
    else if (age <= 60) aging[1] += inv.remaining
    else if (age <= 90) aging[2] += inv.remaining
    else aging[3] += inv.remaining
  }

  // DSO-like per-month average days to pay (using closingDate)
  const monthMap: Record<string, { dt: Date; sum: number; cnt: number }> = {}
  for (const inv of paid) {
    const dt = new Date(inv.closingDate!.getFullYear(), inv.closingDate!.getMonth(), 1)
    const key = `${dt.getFullYear()}-${dt.getMonth()+1}`
    const d2p = Math.round(((+inv.closingDate! - +inv.invoiceDate)/86400000))
    if (!isFinite(d2p) || d2p < 0) continue
    if (!monthMap[key]) monthMap[key] = { dt, sum: 0, cnt: 0 }
    monthMap[key].sum += d2p; monthMap[key].cnt += 1
  }
  const months = Object.values(monthMap).sort((a,b) => +a.dt - +b.dt).map(m => ({ dt: m.dt, avg: m.sum/m.cnt }))

  // Running AR balance per invoice row (cumulative sum of remaining)
  let run = 0
  for (const inv of displayInvoices) { run += inv.remaining; inv.running = run }

  // Reconciliation figures
  const sumInvAmt = displayInvoices.reduce((s,inv) => s + inv.amount, 0)
  const sumPayAmt = payments.reduce((s,p) => s + p.amount, 0)
  const openingRemaining = (invoices.find(inv => inv._synthetic && inv.type === 'Opening')?.remaining) || 0
  let computedOutstanding = 0 // will set after building ledger to match running balance approach
  const expectedOutstanding = beginningBalance + sumInvAmt - sumPayAmt
  const reconcile = { beginningBalance, sumInvoices: sumInvAmt, sumPayments: sumPayAmt, expectedOutstanding, computedOutstanding, openingRemaining, delta: 0 }

  // Debug: pay type distribution and reconcile snapshot
  try {
    const payTypes: Record<string, number> = {}
    for (const p of payments) { const k = String(p.payType || ''); payTypes[k] = (payTypes[k] || 0) + 1 }
    const total = payments.filter(p => p.payType === 'Check').length
    const withMaturity = payments.filter(p => p.payType === 'Check' && !!p.maturityDate).length
    const withoutMaturity = total - withMaturity
    const checkCounts = { total, withMaturity, withoutMaturity }
    const unappliedAfterCarry = advances.filter(a => a.remaining > 0).map(a => ({ date: a.date, remaining: a.remaining }))
    ;(globalThis as any).__arDebug = { ...(globalThis as any).__arDebug, reconcile, payTypes, checkCounts, unapplied, unappliedAfterCarry }
  } catch {}

  // Build Ledger (date-ordered invoices + payments + prepayments + opening)
  type LedgerItem = { date: Date; kind: 'Opening'|'Invoice'|'Payment'|'Prepayment'; description: string; ref?: string; debit: number; credit: number; balance: number }
  const combined: { date: Date; kind: LedgerItem['kind']; amount: number; description: string; ref?: string }[] = []
  for (const inv of invoices) {
    const kind: LedgerItem['kind'] = inv._synthetic && inv.type === 'Opening' ? 'Opening' : (inv.type === 'Prepayment' ? 'Prepayment' : 'Invoice')
    combined.push({ date: inv.invoiceDate, kind, amount: inv.amount, description: inv.type, ref: inv.invoiceNum })
  }
  for (const p of payments) {
    combined.push({ date: p.paymentDate, kind: 'Payment', amount: p.amount, description: p.desc ? p.desc : p.payType || 'Payment' })
  }
  combined.sort((a,b) => +a.date - +b.date)
  const ledger: LedgerItem[] = []
  let bal = 0
  for (const it of combined) {
    if (it.kind === 'Payment') { bal -= it.amount; ledger.push({ date: it.date, kind: it.kind, description: it.description, debit: 0, credit: it.amount, balance: bal }) }
    else if (it.kind === 'Prepayment') { bal += it.amount; const credit = it.amount < 0 ? -it.amount : 0; const debit = it.amount > 0 ? it.amount : 0; ledger.push({ date: it.date, kind: it.kind, description: it.description, debit, credit, balance: bal }) }
    else { bal += it.amount; ledger.push({ date: it.date, kind: it.kind, description: it.description, ref: it.ref, debit: it.amount, credit: 0, balance: bal }) }
  }

  // Set computedOutstanding to the last ledger balance for 1:1 parity with expectedOutstanding
  const ledgerEnd = ledger.length ? ledger[ledger.length - 1].balance : 0
  reconcile.computedOutstanding = ledgerEnd
  reconcile.delta = ledgerEnd - reconcile.expectedOutstanding

  return { invoices: displayInvoices, metrics, aging, months, startDate, reconcile, ledger }
}
