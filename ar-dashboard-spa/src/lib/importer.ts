import * as XLSX from 'xlsx'
import { amountToNumber, normalizePayType } from './parsers'

export type ImportResult = {
  rows: Record<string, any>[]
  autoBeginBalance?: number
}

const HEADER_ALIASES = {
  date: [
    'date','tarih','Ø§Ù„ØªØ§Ø±ÙŠØ®','ØªØ§Ø±ÙŠØ®','tarÄ°h','tariÌ‡h'
  ],
  desc: [
    'aÃ§Ä±klama','aciklama','Ø§Ù„ÙˆØµÙ','Ø§Ù„Ø¨ÙŠØ§Ù†','Ø´Ø±Ø­','description','desc','not','memo'
  ],
  debit: [
    'debit','borÃ§','borc','Ã¶deme','odeme','tahsilat','Ø§Ù„Ù…Ø¨Ù„Øº Ø§Ù„Ù…Ø¯ÙŠÙ†','Ù…Ø¯ÙŠÙ†','payment'
  ],
  credit: [
    'credit','alacak','fatura','Ø§Ù„ÙØ§ØªÙˆØ±Ø©','Ø§Ù„Ù…Ø¨Ù„Øº Ø§Ù„Ø¯Ø§Ø¦Ù†','Ø¯Ø§Ø¦Ù†','invoice','fatura tutari','fatura miktari'
  ],
  paytype: [
    'pay type','payment type','Ã¶deme tipi','odeme tipi','odeme turu','tahsilat tipi','paytype',
    // Some ERPs place payment method under a generic 'project' column name
    'proje','projekt',
    // Arabic labels
    'Ù†ÙˆØ¹ Ø§Ù„Ø¯ÙØ¹','Ø·Ø±ÙŠÙ‚Ø© Ø§Ù„Ø¯ÙØ¹','Ø§Ù„Ù…Ø´Ø±ÙˆØ¹'
  ]
}

const TOTAL_WORDS = [
  'total','sub total','subtotal','page total','genel toplam','toplam','ara toplam',
  'Ø§Ù„Ù…Ø¬Ù…ÙˆØ¹','Ø¥Ø¬Ù…Ø§Ù„ÙŠ','Ø§Ù„Ø§Ø¬Ù…Ø§Ù„ÙŠ','Ø§Ø¬Ù…Ø§Ù„ÙŠ'
]

const BEGIN_BAL_WORDS = [
  'opening balance','beginning balance','balance forward','opening','devir',
  'aÃ§Ä±lÄ±ÅŸ bakiyesi','baslangiÃ§ bakiyesi','baÅŸlangÄ±Ã§ bakiyesi','ilk bakiye',
  'Ø±ØµÙŠØ¯ Ø§ÙØªØªØ§Ø­ÙŠ','Ø§Ù„Ø±ØµÙŠØ¯ Ø§Ù„Ø§ÙØªØªØ§Ø­ÙŠ','Ø±ØµÙŠØ¯ Ø£ÙˆÙ„ Ø§Ù„Ù…Ø¯Ø©','Ø±ØµÙŠØ¯ Ø§ÙˆÙ„ Ø§Ù„Ù…Ø¯Ø©'
]

// Extra canonical Arabic/Turkish forms for opening balance and combined list
const BEGIN_BAL_WORDS_EXTRA = [
  'رصيد اول المدة','رصيد أول المدة','الرصيد الافتتاحي','رصيد افتتاحي','رصيد بداية المدة','ilk bakiye'
]
const BEGIN_BAL_ALL = (BEGIN_BAL_WORDS as string[]).concat(BEGIN_BAL_WORDS_EXTRA)

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

  // Detect opening/beginning balance row immediately after header (strict rules)
  const possibleBeginRow = grid[headerRowIdx + 1] || []
  const hasBeginKeywords = rowHasAny(possibleBeginRow, BEGIN_BAL_ALL)
  // Prefer debit/credit cells only; ignore date/other columns to avoid Excel date serials
  const numsAfterHeader: number[] = []
  if (iDebitPref >= 0 && iDebitPref < possibleBeginRow.length) {
    const n = amountToNumber(possibleBeginRow[iDebitPref]); if (isFinite(n)) numsAfterHeader.push(n)
  }
  if (iCreditPref >= 0 && iCreditPref < possibleBeginRow.length) {
    const n = amountToNumber(possibleBeginRow[iCreditPref]); if (isFinite(n)) numsAfterHeader.push(n)
  }
  const dateCell = (iDatePref >= 0 && iDatePref < possibleBeginRow.length) ? String(possibleBeginRow[iDatePref] ?? '') : ''
  const dateLike = /(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{2,4})/.test(dateCell)
  // Criteria: either keywords, or (no date in date column AND there is exactly one numeric in debit/credit)
  if (!autoBeginBalance && hasBeginKeywords) {
    autoBeginBalance = numsAfterHeader.length ? numsAfterHeader[0] : undefined
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
      headersLower: headers.map((h:any)=>String(h).toLowerCase()),
      autoBeginBalance,
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

