import * as XLSX from 'xlsx'
import { amountToNumber, normalizePayType } from './parsers'

export type ImportResult = {
  rows: Record<string, any>[]
  autoBeginBalance?: number
}

const HEADER_ALIASES = {
  date: [
    'date','tarih','التاريخ','تاريخ','tarİh','tari̇h'
  ],
  desc: [
    'description','açıklama','aciklama','الوصف','شرح','desc','not','memo'
  ],
  debit: [
    'debit','borç','borc','ödeme','odeme','tahsilat','المبلغ المدين','مدين','payment'
  ],
  credit: [
    'credit','alacak','fatura','الفاتورة','المبلغ الدائن','دائن','invoice','fatura tutari','fatura miktari'
  ],
  paytype: [
    'pay type','payment type','ödeme tipi','odeme tipi','odeme turu','tahsilat tipi','نوع الدفع','طريقة الدفع','paytype'
  ]
}

const TOTAL_WORDS = [
  'total','sub total','subtotal','page total','genel toplam','toplam','ara toplam',
  'المجموع','إجمالي','الاجمالي','اجمالي'
]

const BEGIN_BAL_WORDS = [
  'opening balance','beginning balance','balance forward','opening','devir',
  'açılış bakiyesi','baslangiç bakiyesi','başlangıç bakiyesi','ilk bakiye',
  'رصيد افتتاحي','الرصيد الافتتاحي','رصيد أول المدة','رصيد اول المدة'
]

function lc(x: any): string { return String(x ?? '').toLowerCase() }

function rowHasAny(row: any[], words: string[]): boolean {
  const s = lc(row.join(' '))
  return words.some(w => s.includes(w))
}

function isHeaderCandidate(row: any[]): boolean {
  const r = row.map(lc)
  const joined = ' ' + r.join(' ') + ' '
  const has = (arr: string[]) => arr.some(k => joined.includes(' ' + k + ' '))
  const nonEmpty = r.filter(x => x && x.trim().length > 0).length
  let score = 0
  if (has(HEADER_ALIASES.date)) score += 2
  if (has(HEADER_ALIASES.desc)) score += 2
  if (has(HEADER_ALIASES.debit)) score += 1
  if (has(HEADER_ALIASES.credit)) score += 1
  if (nonEmpty >= 3) score += 1
  return score >= 3
}

function findHeaderRow(grid: any[][]): number {
  let bestIdx = -1, bestScore = -1
  for (let i = 0; i < grid.length; i++) {
    const row = grid[i].map(lc)
    // Score headers by presence of key tokens
    let score = 0
    const joined = ' ' + row.join(' ') + ' '
    const has = (arr: string[]) => arr.some(k => joined.includes(' ' + k + ' '))
    if (has(HEADER_ALIASES.date)) score += 2
    if (has(HEADER_ALIASES.desc)) score += 2
    if (has(HEADER_ALIASES.debit)) score += 1
    if (has(HEADER_ALIASES.credit)) score += 1
    if (row.filter(c => c).length >= 3) score += 1
    if (score > bestScore) { bestScore = score; bestIdx = i }
  }
  return bestIdx >= 0 ? bestIdx : 0
}

function buildObjects(headers: any[], rows: any[][]): Record<string, any>[] {
  const origHeaders = headers.map((h: any) => String(h ?? '').trim())
  const lower = origHeaders.map(h => h.toLowerCase())
  const idxOf = (aliases: string[]) => {
    for (let i = 0; i < lower.length; i++) {
      for (const a of aliases) { if (lower[i].includes(a)) return i }
    }
    return -1
  }
  const iDate = idxOf(HEADER_ALIASES.date)
  const iDesc = idxOf(HEADER_ALIASES.desc)
  const iDebit = idxOf(HEADER_ALIASES.debit)
  const iCredit= idxOf(HEADER_ALIASES.credit)
  const iPayTp = idxOf(HEADER_ALIASES.paytype)

  return rows.map(r => {
    const obj: Record<string, any> = {}
    for (let c = 0; c < origHeaders.length; c++) {
      const key = origHeaders[c] || `col_${c+1}`
      obj[key] = r[c] ?? ''
    }
    // Also provide canonical keys to make downstream matching robust
    if (iDate  >= 0) obj['date'] = r[iDate] ?? obj['date'] ?? ''
    if (iDesc  >= 0) obj['description'] = r[iDesc] ?? obj['description'] ?? ''
    if (iDebit >= 0) obj['debit'] = r[iDebit] ?? obj['debit'] ?? ''
    if (iCredit>= 0) obj['credit'] = r[iCredit] ?? obj['credit'] ?? ''
    if (iPayTp >= 0) obj['pay type'] = r[iPayTp] ?? obj['pay type'] ?? ''
    return obj
  })
}

function isTotalLike(obj: Record<string, any>): boolean {
  const values = Object.values(obj).map(v => lc(v))
  return values.some(v => TOTAL_WORDS.some(w => v.includes(w)))
}

export function extractTable(ws: XLSX.WorkSheet): ImportResult {
  const grid: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true }) as any[][]
  let headerRowIdx = 0
  if (grid.length > 1 && isHeaderCandidate(grid[1])) headerRowIdx = 1
  else headerRowIdx = findHeaderRow(grid)
  const headers = grid[headerRowIdx] || []

  // Preface (above header): try to detect opening balance
  let autoBeginBalance: number | undefined
  for (let i = 0; i < headerRowIdx; i++) {
    const row = grid[i]
    if (!row || row.length === 0) continue
    if (rowHasAny(row, BEGIN_BAL_WORDS)) {
      // Find the biggest numeric in row
      let best = NaN
      for (const cell of row) {
        const n = amountToNumber(cell)
        if (isFinite(n) && Math.abs(n) > Math.abs(best || 0)) best = n
      }
      if (isFinite(best)) { autoBeginBalance = best; break }
    }
  }

  // Detect opening/beginning balance row immediately after header
  const possibleBeginRow = grid[headerRowIdx + 1] || []
  const hasBeginKeywords = rowHasAny(possibleBeginRow, BEGIN_BAL_WORDS)
  const anyDateLike = (possibleBeginRow || []).some(cell => {
    const s = String(cell ?? '')
    return /(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{2,4})/.test(s)
  })
  const numericInRow = (possibleBeginRow || []).map(x => amountToNumber(x)).filter(n => isFinite(n))
  if (!autoBeginBalance && (hasBeginKeywords || (!anyDateLike && numericInRow.length > 0))) {
    // choose the largest magnitude number as opening balance
    autoBeginBalance = numericInRow.reduce((a,b) => Math.abs(b) > Math.abs(a) ? b : a, 0)
  }

  // Body rows start after header + possible opening balance row
  const startIdx = (autoBeginBalance != null) ? (headerRowIdx + 2) : (headerRowIdx + 1)
  const body = grid.slice(startIdx)
  // Stop at the first trailing set of 5+ consecutive empty rows
  let end = body.length
  let emptyStreak = 0
  for (let i = 0; i < body.length; i++) {
    const r = body[i]
    const nonEmpty = r.some(v => (v != null && String(v).trim() !== ''))
    if (!nonEmpty) { emptyStreak++; if (emptyStreak >= 5) { end = i - emptyStreak + 1; break } }
    else emptyStreak = 0
  }
  const trimmed = body.slice(0, end)
  const objects = buildObjects(headers, trimmed)

  // Filter out totals and fully empty rows
  const filtered = objects.filter(o => {
    const hasAny = Object.values(o).some(v => String(v ?? '').trim() !== '')
    if (!hasAny) return false
    if (isTotalLike(o)) return false
    return true
  })

  // Debug info on window for inspection
  try {
    ;(globalThis as any).__arDebug = {
      headerRowIdx,
      headers,
      autoBeginBalance,
      startIdx,
      totalRows: grid.length,
      bodyRows: body.length,
      filteredRows: filtered.length,
      sampleRow: filtered[0]
    }
    // Console logs for quick visibility
    console.debug('[importer] headerRowIdx:', headerRowIdx, 'autoBeginBalance:', autoBeginBalance, 'rows:', filtered.length)
  } catch {}

  return { rows: filtered, autoBeginBalance }
}
