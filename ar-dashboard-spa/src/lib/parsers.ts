import dayjs from 'dayjs'

// Normalize Arabic-Indic and Eastern Arabic-Indic digits to ASCII
export function normalizeDigits(v: any): string {
  const s = String(v ?? '')
  let out = ''
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i)
    if (code >= 0x0660 && code <= 0x0669) { out += String(code - 0x0660); continue }
    if (code >= 0x06F0 && code <= 0x06F9) { out += String(code - 0x06F0); continue }
    out += s[i]
  }
  return out
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
  const norm = head.replace(/[\.\-]/g, '/')
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
  const str = normalizeDigits(String(s))
  // Prefer labelled maturity/due terms close to a date
  const labelTerms = ['vade','vadesi','vade tarihi','son ödeme','son odeme','maturity','maturity date','due','due date','استحقاق','تاريخ الاستحقاق','الاستحقاق','تستحق']
  const labelPattern = new RegExp('(?:' + labelTerms.join('|') + ')[^0-9]{0,20}(\\d{1,2}[\\/\\.\\-]\\d{1,2}[\\/\\.\\-]\\d{2,4})','i')
  let m = str.match(labelPattern)
  if (m) {
    const d = m[1]
    const parts = d.split(/[\.\/-]/)
    let y = parts[2]
    if ((y as string).length === 2) y = Number(y) >= 30 ? ('19' + y) : ('20' + y)
    return parseDMY(parts[0] + '/' + parts[1] + '/' + y)
  }

  // 1) dd/mm/yyyy or dd.mm.yyyy or dd-mm-yy (take last match)
  let last: RegExpExecArray | null = null
  const re1 = /(\b\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{2,4})\b/g
  let mmatch: RegExpExecArray | null
  while ((mmatch = re1.exec(str)) !== null) last = mmatch
  if (last) {
    let y = String(last[3])
    if (y.length === 2) y = Number(y) >= 30 ? ('19' + y) : ('20' + y)
    return parseDMY(last[1] + '/' + last[2] + '/' + y)
  }
  // 2) yyyy-mm-dd or yyyy/mm/dd (take last match)
  last = null
  const re2 = /\b(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})\b/g
  while ((mmatch = re2.exec(str)) !== null) last = mmatch
  if (last) {
    const y = Number(last[1]), mm = Number(last[2]), dd = Number(last[3])
    const dt = new Date(y, mm - 1, dd)
    return isNaN(+dt) ? null : dt
  }
  // 3) dd Mon yyyy (EN) or dd Ay yyyy (TR) (take last match)
  const monthMap: Record<string, number> = {
    jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12,
    ocak:1,şubat:2,subat:2,mart:3,nisan:4,mayıs:5,mayis:5,haziran:6,temmuz:7,ağustos:8,agustos:8,eylül:9,eylul:9,ekim:10,kasım:11,kasim:11,aralık:12,aralik:12
  }
  last = null
  const re3 = /\b(\d{1,2})\s+([a-zçğıöşü]+)\s+(\d{2,4})\b/gi
  while ((mmatch = re3.exec(str.toLowerCase())) !== null) last = mmatch
  if (last) {
    const dd = Number(last[1]); const key = last[2].normalize('NFC')
    const mm = monthMap[key as keyof typeof monthMap]
    let y = last[3]; if ((y as string).length === 2) y = Number(y) >= 30 ? ('19'+y) : ('20'+y)
    const yyyy = Number(y)
    if (mm) {
      const dt = new Date(yyyy, mm - 1, dd)
      return isNaN(+dt) ? null : dt
    }
  }
  return null
}

export function lc(s: any): string {
  return String(s ?? '').toLowerCase()
}

export function normalizePayType(v: any): { type: string; termDays: number | null } {
  const t = lc(v).trim()
  if (!t) return { type: '', termDays: null }
  // Turkish + English + Arabic variants
  if (/(cek|çek|cheque|check|senet|vadeli|بولصة|شيك)/.test(t)) return { type: 'Check', termDays: 90 }
  if (/(kk|kredi\s*kart|credit\s*card|card|kart|بطاقة|فيزا|كردت)/.test(t)) return { type: 'Card', termDays: 30 }
  if (/(peşin|pesin|cash|nakit|نقد|نقدي|كاش)/.test(t)) return { type: 'Cash', termDays: 30 }
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
