import * as XLSX from 'xlsx'
import { amountToNumber, normalizePayType } from './parsers'

export type ImportResult = {
  rows: Record<string, any>[]
  autoBeginBalance?: number
}

const HEADER_ALIASES = {
  date: ['date','tarih','tarıh','التاريخ','تاريخ'],
  desc: ['açıklama','aciklama','description','desc','not','memo','البيان','الوصف','ملاحظات'],
  debit: ['debit','borç','borc','ödeme','odeme','tahsilat','payment','مدين'],
  credit: ['credit','alacak','fatura','invoice','fatura tutari','fatura miktari','دائن'],
  paytype: ['pay type','payment type','ödeme tipi','odeme tipi','odeme turu','tahsilat tipi','paytype','proje','projekt','المشروع','مشروع']
}

const TOTAL_WORDS = ['total','sub total','subtotal','page total','genel toplam','toplam','ara toplam','الإجمالي','المجموع']

const BEGIN_BAL_WORDS = ['opening balance','beginning balance','balance forward','opening','devir','ilk bakiye','رصيد اول المدة','رصيد أول المدة','الرصيد الافتتاحي','رصيد افتتاحي','رصيد بداية المدة']

// Extra canonical Arabic/Turkish forms for opening balance and combined list
const BEGIN_BAL_WORDS_EXTRA: string[] = []
const BEGIN_BAL_ALL = (BEGIN_BAL_WORDS as string[]).concat(BEGIN_BAL_WORDS_EXTRA)

function lc(x: any): string { return String(x ?? '').toLowerCase() }

function rowHasAny(row: any[], words: string[]): boolean {
  const s = lc(row.join(' '))
  return words.some(w => s.includes(w))
}

function isHeaderCandidate(row: any[]): boolean {
  const r = row.map(lc)
  const joined = ' ' + r.join(' ') + ' '
  const dateAl = ([] as string[]).concat(HEADER_ALIASES.date, ['tarih','tarıh','التاريخ','تاريخ'])
  const descAl = ([] as string[]).concat(HEADER_ALIASES.desc, ['açıklama','aciklama','البيان','الوصف','ملاحظات'])
  const debitAl= ([] as string[]).concat(HEADER_ALIASES.debit, ['borç','borc','ödeme','odeme','tahsilat','payment','مدين'])
  const creditAl=([] as string[]).concat(HEADER_ALIASES.credit, ['alacak','invoice','fatura','دائن'])
  const has = (arr: string[]) => arr.some(k => k && joined.includes(' ' + lc(k) + ' '))
  const nonEmpty = r.filter(x => x && x.trim().length > 0).length
  const numericish = r.filter(x => /^\d{4,}$/.test(x || '')).length
  let score = 0
  if (has(dateAl)) score += 2
  if (has(descAl)) score += 2
  if (has(debitAl)) score += 1
  if (has(creditAl)) score += 1
  if (nonEmpty >= 3) score += 1
  // Penalize rows dominated by numeric tokens
  if (numericish >= Math.max(3, Math.floor(r.length * 0.5))) score -= 2
  return score >= 3
}

function findHeaderRow(grid: any[][]): number {
  let bestIdx = -1, bestScore = -1
  const dateAl = ([] as string[]).concat(HEADER_ALIASES.date, ['tarih','tarıh','التاريخ','تاريخ'])
  const descAl = ([] as string[]).concat(HEADER_ALIASES.desc, ['açıklama','aciklama','البيان','الوصف','ملاحظات'])
  const debitAl= ([] as string[]).concat(HEADER_ALIASES.debit, ['borç','borc','ödeme','odeme','tahsilat','payment','مدين'])
  const creditAl=([] as string[]).concat(HEADER_ALIASES.credit, ['alacak','invoice','fatura','دائن'])
  for (let i = 0; i < Math.min(grid.length, 15); i++) {
    const row = grid[i].map(lc)
    // Score headers by presence of key tokens
    let score = 0
    const joined = ' ' + row.join(' ') + ' '
    const has = (arr: string[]) => arr.some(k => k && joined.includes(' ' + lc(k) + ' '))
    if (has(dateAl)) score += 2
    if (has(descAl)) score += 2
    if (has(debitAl)) score += 1
    if (has(creditAl)) score += 1
    if (row.filter(c => c).length >= 3) score += 1
    const numericish = row.filter(x => /^\d{4,}$/.test(x || '')).length
    if (numericish >= Math.max(3, Math.floor(row.length * 0.5))) score -= 2
    if (score > bestScore) { bestScore = score; bestIdx = i }
  }
  // Fallback: favor the first row among the first 5 that matches candidate rules
  if (bestIdx < 0) {
    for (let i = 0; i < Math.min(grid.length, 5); i++) if (isHeaderCandidate(grid[i])) return i
  }
  return bestIdx >= 0 ? bestIdx : 0
}

function buildObjects(headers: any[], rows: any[][]): Record<string, any>[] {
  const origHeaders = headers.map((h: any) => String(h ?? '').trim())
  const lower = origHeaders.map(h => String(h ?? '').toLowerCase())
  const idxOf = (aliases: string[]) => {
    for (let i = 0; i < lower.length; i++) {
      const val = String(lower[i] ?? '')
      for (const a of aliases) { const aa = String(a ?? ''); if (aa && val.includes(aa)) return i }
    }
    return -1
  }
  const dateAliases  = ([] as string[]).concat(HEADER_ALIASES.date, ['tarih','tarıh','التاريخ','تاريخ'])
  const descAliases  = ([] as string[]).concat(HEADER_ALIASES.desc, ['açıklama','aciklama','البيان','الوصف','ملاحظات'])
  const debitAliases = ([] as string[]).concat(HEADER_ALIASES.debit, ['borç','borc','ödeme','odeme','tahsilat','payment','مدين'])
  const creditAliases= ([] as string[]).concat(HEADER_ALIASES.credit, ['alacak','invoice','fatura','دائن'])
  const paytpAliases = ([] as string[]).concat(HEADER_ALIASES.paytype, ['proje','projekt','المشروع','مشروع'])
  const iDate = idxOf(dateAliases)
  const iDesc = idxOf(descAliases)
  const iDebit = idxOf(debitAliases)
  const iCredit= idxOf(creditAliases)
  const iPayTp = idxOf(paytpAliases)
  // Additional: collect all description-like columns to prioritize non-empty text
  const descCandidateIdx: number[] = []
  for (let i = 0; i < lower.length; i++) {
    const h = lower[i]
    if (h && HEADER_ALIASES.desc.some(a => h.includes(a))) descCandidateIdx.push(i)
  }

  return rows.map(r => {
    const obj: Record<string, any> = {}
    for (let c = 0; c < origHeaders.length; c++) {
      const key = origHeaders[c] || `col_${c+1}`
      obj[key] = r[c] ?? ''
    }
    // Also provide canonical keys to make downstream matching robust
    if (iDate  >= 0) obj['date'] = r[iDate] ?? obj['date'] ?? ''
    // description: prefer AÃ§Ä±klama/Ø§Ù„ÙˆØµÙ/Ø§Ù„Ø¨ÙŠØ§Ù†/Ø´Ø±Ø­ over plain 'description', then fallbacks
    let descVal = ''
    for (const idx of descCandidateIdx) { if (r[idx] != null && String(r[idx]).trim() !== '') { descVal = r[idx]; break } }
    if (!descVal && iDesc >= 0) descVal = r[iDesc] ?? ''
    if (!descVal && typeof obj['desc'] === 'string') descVal = obj['desc']
    if (!descVal && typeof obj['not'] === 'string') descVal = obj['not']
    if (!descVal && typeof obj['memo'] === 'string') descVal = obj['memo']
    if (descVal) obj['description'] = descVal
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

  // Preface (above header): try to detect opening balance cautiously
  let autoBeginBalance: number | undefined
  // Pre-compute header indices to constrain numeric detection to debit/credit columns only
  const headersLowerPreface = (grid[headerRowIdx] || []).map((h:any)=>String(h??'').toLowerCase())
  const idxOfPref = (aliases: string[]) => {
    for (let i = 0; i < headersLowerPreface.length; i++) {
      const val = String(headersLowerPreface[i] ?? '')
      for (const a of aliases) { const aa = String(a ?? ''); if (aa && val.includes(aa)) return i }
    }
    return -1
  }
  const iDatePref  = idxOfPref(HEADER_ALIASES.date)
  const iDebitPref = idxOfPref(HEADER_ALIASES.debit)
  const iCreditPref= idxOfPref(HEADER_ALIASES.credit)
  for (let i = 0; i < headerRowIdx; i++) {
    const row = grid[i]
    if (!row || row.length === 0) continue
    if (rowHasAny(row, BEGIN_BAL_ALL)) {
      // Consider only debit/credit columns for numeric
      const candidates: number[] = []
      const pushNum = (idx: number) => { if (idx >= 0 && idx < row.length) { const n = amountToNumber(row[idx]); if (isFinite(n)) candidates.push(n) } }
      pushNum(iDebitPref); pushNum(iCreditPref)
      if (candidates.length === 0) continue
      const best = candidates.reduce((a,b)=> Math.abs(b) > Math.abs(a)? b : a, 0)
      autoBeginBalance = best
      break
    }
  }

  // Detect opening/beginning balance row within next few rows after header (keywords only)
  let openingRowIndex: number | undefined
  for (let i = headerRowIdx + 1; i <= Math.min(grid.length - 1, headerRowIdx + 5); i++) {
    const row = grid[i] || []
    if (!rowHasAny(row, BEGIN_BAL_ALL)) continue
    const nums: number[] = []
    if (iDebitPref >= 0 && iDebitPref < row.length) { const n = amountToNumber(row[iDebitPref]); if (isFinite(n)) nums.push(n) }
    if (iCreditPref >= 0 && iCreditPref < row.length) { const n = amountToNumber(row[iCreditPref]); if (isFinite(n)) nums.push(n) }
    if (nums.length) { autoBeginBalance = nums.reduce((a,b)=> Math.abs(b)>Math.abs(a)?b:a, 0); openingRowIndex = i; break }
  }

  // Body rows start after header + possible opening balance row
  const startIdx = (openingRowIndex != null) ? (openingRowIndex + 1) : (headerRowIdx + 1)
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
      headersLower: headers.map((h:any)=>String(h).toLowerCase()),
      autoBeginBalance,
      openingRowIndex,
      startIdx,
      totalRows: grid.length,
      bodyRows: body.length,
      filteredRows: filtered.length,
      sampleRow: filtered[0],
      // reset runtime collections for this upload
      checkInspect: [],
      checkNoMatExamples: [],
      checkAll: [],
      payTypeSamples: []
    }
    // Console logs for quick visibility
    console.debug('[importer] headerRowIdx:', headerRowIdx, 'autoBeginBalance:', autoBeginBalance, 'rows:', filtered.length)
  } catch {}

  return { rows: filtered, autoBeginBalance }
}


