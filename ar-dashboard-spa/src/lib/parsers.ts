import dayjs from 'dayjs'

// Normalize Arabic-Indic digits and trim
export function normalizeDigits(v: any): string {
  const s = String(v ?? '')
  // Arabic-Indic (٠١٢٣٤٥٦٧٨٩) and Eastern Arabic-Indic (۰۱۲۳۴۵۶۷۸۹)
  const map: Record<string, string> = {
    '٠': '0','١': '1','٢': '2','٣': '3','٤': '4','٥': '5','٦': '6','٧': '7','٨': '8','٩': '9',
    '۰': '0','۱': '1','۲': '2','۳': '3','۴': '4','۵': '5','۶': '6','۷': '7','۸': '8','۹': '9',
  }
  return s.replace(/[٠-٩۰-۹]/g, ch => map[ch] || ch)
}

export function parseDMY(v: any): Date | null {
  if (v == null || v === '') return null
  if (v instanceof Date && !isNaN(+v)) return v
  if (typeof v === 'number' && isFinite(v)) {
    const base = new Date(Date.UTC(1899, 11, 30))
    const d = Math.floor(v), MS = 86400000
    return new Date(base.getTime() + d * MS + Math.round((v - d) * MS))
  }
  const s = normalizeDigits(v).trim()
  const head = s.includes(' ') ? s.slice(0, s.indexOf(' ')) : s
  const norm = head.replace(/[.\-]/g, '/')
  const parts = norm.split('/')
  if (parts.length !== 3) return null
  const dd = Number(parts[0]), mm = Number(parts[1])
  const yRaw = parts[2].trim()
  if (!isFinite(dd) || !isFinite(mm)) return null
  let yyyy: number
  if (yRaw.length === 2) {
    const yy = Number(yRaw); if (!isFinite(yy)) return null
    yyyy = (yy >= 30) ? (1900 + yy) : (2000 + yy)
  } else {
    yyyy = Number(yRaw)
  }
  if (!isFinite(yyyy)) return null
  const dt = new Date(yyyy, mm - 1, dd)
  if (dt.getFullYear() !== yyyy || dt.getMonth() !== (mm - 1) || dt.getDate() !== dd) return null
  return dt
}

export function daysBetween(a: any, b: any): number {
  const ad = parseDMY(a), bd = parseDMY(b)
  if (!ad || !bd) return NaN
  return Math.round((+bd - +ad) / 86400000)
}

export function amountToNumber(v: any): number {
  if (v == null || v === '') return NaN
  let s = normalizeDigits(v).trim().replace(/\u00A0/g, ' ').replace(/[^\d,.\-]/g, '')
  if (s.includes(',') && s.includes('.')) {
    const lc = s.lastIndexOf(','), ld = s.lastIndexOf('.')
    s = (lc > ld) ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '')
  } else if (s.includes(',')) {
    const p = s.split(',')
    s = (p.length === 2 && p[1].length > 0 && p[1].length <= 2) ? (p[0].replace(/,/g, '') + '.' + p[1]) : s.replace(/,/g, '')
  } else if (s.includes('.')) {
    const dp = s.split('.')
    if (!(dp.length === 2 && dp[1].length > 0 && dp[1].length <= 2)) s = s.replace(/\./g, '')
  }
  const n = Number(s)
  return isFinite(n) ? n : NaN
}

export function extractDateFromText(s: any): Date | null {
  if (!s) return null
  const str = String(s)
  const re = /(\b\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{2,4})\b/
  const m = str.match(re)
  if (!m) return null
  let y = String(m[3])
  if (y.length === 2) y = Number(y) >= 30 ? ('19' + y) : ('20' + y)
  return parseDMY(m[1] + '/' + m[2] + '/' + y)
}

export function lc(s: any): string {
  return String(s ?? '').toLowerCase()
}

export function normalizePayType(v: any): { type: string; termDays: number | null } {
  const t = lc(v).trim()
  if (!t) return { type: '', termDays: null }
  if (/(cek|çek|cheque|check|vadeli)/.test(t)) return { type: 'Check', termDays: 90 }
  if (/(kk|kredi\s*kart|credit\s*card|card)/.test(t)) return { type: 'Card', termDays: 30 }
  if (/(peşin|pesin|cash|nakit)/.test(t)) return { type: 'Cash', termDays: 30 }
  return { type: '', termDays: null }
}

export function mode(arr: Array<number | null | undefined>, def: number): number {
  const freq: Record<string, number> = {}
  let best: number | null = null, bestCount = -1
  for (const v of arr) {
    if (v == null) continue
    const k = String(v)
    freq[k] = (freq[k] || 0) + 1
    if (freq[k] > bestCount) { best = Number(v); bestCount = freq[k] }
  }
  return best ?? def
}

