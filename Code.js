/******************************
 * Customer Risk & Payments Analysis — Universal (Fixed Headers)
 * Adds Pay Type support with per-type terms.
 ******************************/

/** Day-first parser: D/M/YY or DD/MM/YYYY; supports '-', '.', or '/'; supports Sheets serials. */
function parseDMY_(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number' && isFinite(v)) { // Sheets serial
    var base = new Date(Date.UTC(1899, 11, 30));
    var d = Math.floor(v), MS = 86400000;
    return new Date(base.getTime() + d * MS + Math.round((v - d) * MS));
  }
  var s = String(v).trim();
  var head = s.indexOf(' ') > -1 ? s.slice(0, s.indexOf(' ')) : s; // drop time
  var norm = head.replace(/[.\-]/g, '/');
  var parts = norm.split('/');
  if (parts.length !== 3) return null;
  var dd = Number(parts[0]), mm = Number(parts[1]), yRaw = parts[2].trim();
  if (!isFinite(dd) || !isFinite(mm)) return null;
  var yyyy;
  if (yRaw.length === 2) {
    var yy = Number(yRaw); if (!isFinite(yy)) return null;
    yyyy = (yy >= 30) ? (1900 + yy) : (2000 + yy);
  } else {
    yyyy = Number(yRaw);
  }
  if (!isFinite(yyyy)) return null;
  var dt = new Date(yyyy, mm - 1, dd);
  if (dt.getFullYear() !== yyyy || dt.getMonth() !== (mm - 1) || dt.getDate() !== dd) return null;
  return dt;
}

/** Integer calendar-day difference: maturity - payment. Returns NaN if invalid. */
function daysBetween_(paymentVal, maturityVal) {
  var pd = parseDMY_(paymentVal), md = parseDMY_(maturityVal);
  if (!pd || !md) return NaN;
  return Math.round((md - pd) / 86400000);
}

/** Normalize TRY amount to Number (handles "437,069", "26.639,32", "26,639.32", etc.). */
function amountToNumber_(v) {
  if (v == null || v === '') return NaN;
  var s = String(v).trim()
    .replace(/\u00A0/g, ' ')       // NBSP to space
    .replace(/[^\d,.\-]/g, '');    // keep digits , . -
  if (s.indexOf(',') > -1 && s.indexOf('.') > -1) {
    var lc = s.lastIndexOf(','), ld = s.lastIndexOf('.');
    s = (lc > ld) ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  } else if (s.indexOf(',') > -1) {
    var p = s.split(',');
    s = (p.length === 2 && p[1].length > 0 && p[1].length <= 2) ? (p[0].replace(/,/g, '') + '.' + p[1]) : s.replace(/,/g, '');
  } else if (s.indexOf('.') > -1) {
    var dp = s.split('.');
    if (!(dp.length === 2 && dp[1].length > 0 && dp[1].length <= 2)) s = s.replace(/\./g, '');
  }
  var n = Number(s);
  return isFinite(n) ? n : NaN;
}

/** Mode (most frequent) of numeric array; fallback defaultVal if empty. */
function mode_(arr, defaultVal) {
  if (!arr || !arr.length) return defaultVal;
  var freq = {}, best = null, bestCount = -1;
  for (var i = 0; i < arr.length; i++) {
    var v = arr[i]; if (v == null) continue;
    freq[v] = (freq[v] || 0) + 1;
    if (freq[v] > bestCount) { best = v; bestCount = freq[v]; }
  }
  return best != null ? Number(best) : defaultVal;
}

/** Assessment helpers */
function assessLowerBetter_(val, goodMax, avgMax) {
  if (val === "" || val == null || isNaN(val)) return "";
  return (val <= goodMax) ? "Good" : (val <= avgMax) ? "Average" : "Poor";
}
function assessHigherBetter_(val, goodMin, avgMin) {
  if (val === "" || val == null || isNaN(val)) return "";
  return (val >= goodMin) ? "Good" : (val >= avgMin) ? "Average" : "Poor";
}

/** Extract a date from text like "30/03/2025" or "31.03.25" (helps maturity detection). */
function extractDateFromText_(s) {
  if (!s) return null;
  var str = String(s);
  var re = /(\b\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{2,4})\b/; // D/M/YY(YY)
  var m = str.match(re);
  if (!m) return null;
  var d = m[1], M = m[2], y = String(m[3]);
  if (y.length === 2) y = Number(y) >= 30 ? ('19' + y) : ('20' + y);
  return parseDMY_(d + '/' + M + '/' + y);
}

/** Utility: find column by any of the given names (lowercased). Returns 1-based index or -1. */
function findCol_(headers, names) {
  for (var i = 0; i < names.length; i++) {
    var idx = headers.indexOf(names[i].toLowerCase());
    if (idx >= 0) return idx + 1;
  }
  return -1;
}

/** Utility: very light Turkish diacritic strip + lowercase. */
function lc_(s) {
  return String(s || '').toLowerCase()
    .replace(/ç/g,'c').replace(/ğ/g,'g').replace(/ı/g,'i')
    .replace(/ö/g,'o').replace(/ş/g,'s').replace(/ü/g,'u');
}

/** Normalize Pay Type to {type, termDays}. */
function normalizePayType_(v) {
  var t = lc_(v).trim();
  if (!t) return { type: '', termDays: null };
  if (/(cek|çek|cheque|check|vadeli)/.test(t)) return { type: 'Check', termDays: 90 };
  if (/(kk|kredi\s*kart|credit\s*card|card)/.test(t)) return { type: 'Card', termDays: 30 };
  if (/(pesin|peşin|cash|nakit)/.test(t)) return { type: 'Cash', termDays: 30 };
  return { type: '', termDays: null };
}

// --- TR-aware overrides (appended) ---
/** Improved Turkish diacritic normalization */
function lc2_(s) {
  var str = String(s == null ? '' : s);
  var map = { 'Ç':'c','ç':'c','Ğ':'g','ğ':'g','İ':'i','I':'i','ı':'i','Ö':'o','ö':'o','Ş':'s','ş':'s','Ü':'u','ü':'u' };
  var out = '';
  for (var i = 0; i < str.length; i++) {
    var ch = str[i];
    out += (map[ch] != null ? map[ch] : ch);
  }
  return out.toLowerCase();
}

/** Improved pay type detection using lc2_ */
function normalizePayType2_(v) {
  var t = lc2_(v).trim();
  if (!t) return { type: '', termDays: null };
  if (/(cek|çek|cheque|check|senet|vadeli)/.test(t)) return { type: 'Check', termDays: 90 };
  if (/(\bkk\b|k\.k\.|kredi\s*kart|credit\s*card|card|\bkart\b)/.test(t)) return { type: 'Card', termDays: 30 };
  if (/(\bpesin\b|cash|nakit)/.test(t)) return { type: 'Cash', termDays: 30 };
  return { type: '', termDays: null };
}

/** Read Beginning Balance (same as before). */
function readBeginningBalance_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Input") || ss.getSheetByName("input") || ss.getSheetByName("BS Input") || ss.getSheetByName("BS input") || ss.getSheetByName("BSInput") || ss.getSheetByName("Bs Input") || ss.getSheetByName("Universal");
  if (!sh) throw new Error('Input sheet not found. Expected one of: BS Input, BS input, BSInput, Bs Input, Universal');

  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow >= 2) {
    var headers = sh.getRange(1,1,1,lastCol).getValues()[0]
      .map(function(h){ return String(h || '').trim().toLowerCase(); });

    var beginNames = [
      'beginning balance','opening balance','start balance',
      'açılış bakiyesi','acilis bakiyesi','başlangıç bakiyesi','baslangic bakiyesi'
    ];
    var idx = findCol_(headers, beginNames);
    if (idx > 0) {
      var vals = sh.getRange(2, idx, lastRow - 1, 1).getValues();
      for (var i = 0; i < vals.length; i++) {
        var n = amountToNumber_(vals[i][0]);
        if (isFinite(n)) return n;
      }
    }
  }

  var label = sh.getRange("G1").getValue();
  if (!label) sh.getRange("G1").setValue("Beginning Balance").setFontWeight("bold");
  var v = amountToNumber_(sh.getRange("G2").getValue());
  return isFinite(v) ? v : 0;
}

/** Read BS Input sheet and build invoices/payments. */
function readInput_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Input") || ss.getSheetByName("input") || ss.getSheetByName("BS Input") || ss.getSheetByName("BS input") || ss.getSheetByName("BSInput") || ss.getSheetByName("Bs Input") || ss.getSheetByName("Universal");
  if (!sh) throw new Error('Input sheet not found. Expected one of: Input, BS Input, BS input, BSInput, Bs Input, Universal');

  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return { invoices: [], payments: [], firstInvoiceDate: null, firstTransactionDate: null };

  var headersRaw = sh.getRange(1,1,1,lastCol).getValues()[0];
  var headers = headersRaw.map(function(h){ return String(h || '').trim().toLowerCase(); });

  var creditNames = ['credit','alacak','alacaklar','invoice','fatura','fatura tutarı','fatura miktarı'];
  var debitNames  = ['debit','borç','borclar','payment','ödeme','tahsilat','odeme'];
  var descNames   = ['description','açıklama','aciklama','desc','not','memo'];
  var dateNames   = ['date','tarih'];
  var payTypeNames= ['pay type','payment type','odeme tipi','ödeme tipi','odeme turu','ödeme türü','tahsilat tipi','paytype'];

  var cCredit = findCol_(headers, creditNames);
  var cDebit  = findCol_(headers, debitNames);
  var cDesc   = findCol_(headers, descNames);
  var cDate   = findCol_(headers, dateNames);
  var cPayTp  = findCol_(headers, payTypeNames);

  if (cCredit === -1 || cDebit === -1 || cDesc === -1 || cDate === -1) {
    throw new Error('Headers must include: Credit | Debit | Description | Date (Turkish equivalents supported).');
  }

  if (lastRow >= 2) sh.getRange(2, cDate, lastRow-1, 1).setNumberFormat("dd/MM/yyyy");

  var values = sh.getRange(2, 1, lastRow-1, lastCol).getValues();

  var invoices = [], payments = [], firstInvoiceDate = null, firstTransactionDate = null;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var credit = amountToNumber_(row[cCredit - 1]);
    var debit  = amountToNumber_(row[cDebit  - 1]);
    var desc   = row[cDesc - 1];
    var date   = parseDMY_(row[cDate - 1]) || extractDateFromText_(desc);

    var empty = [credit, debit, desc, date].every(function(v){ return v == null || v === '' || (typeof v === 'number' && isNaN(v)); });
    if (empty) continue;

    if (isFinite(credit) && credit > 0) {
      if (!date) continue;
      var invNumMatch = String(desc || '').match(/(No\s*\S+|\b[A-Z0-9\-]{6,}\b)/);
      invoices.push({
        invoiceDate: date,
        invoiceNum: invNumMatch ? invNumMatch[0] : "",
        type: "Invoice",
        amount: credit,
        remaining: credit,
        term: 30, // set later via inferred
        paid: false,
        closingDate: null
      });
      if (!firstInvoiceDate || date < firstInvoiceDate) firstInvoiceDate = date;
      if (!firstTransactionDate || date < firstTransactionDate) firstTransactionDate = date;
    }

    if (isFinite(debit) && debit > 0) {
      if (!date) continue;
      var payTypeRaw = (cPayTp > 0) ? row[cPayTp - 1] : '';
      var norm = normalizePayType2_(payTypeRaw);
      payments.push({
        paymentDate: date,
        amount: debit,
        maturityDate: extractDateFromText_(desc) || null,
        payType: norm.type,
        expectedTerm: norm.termDays // null if unknown
      });
      if (!firstTransactionDate || date < firstTransactionDate) firstTransactionDate = date;
    }
  }

  // Infer invoice term: prefer mode of payment expectedTerm; else fallback to 30/60/90 from maturities; else default 30.
  var paymentTerms = payments.map(function(p){ return p.expectedTerm; }).filter(function(x){ return x != null; });
  var inferredTerm;
  if (paymentTerms.length) {
    inferredTerm = mode_(paymentTerms, 30);
  } else {
    var deltas = [];
    for (var j = 0; j < payments.length; j++) {
      var pd = daysBetween_(payments[j].paymentDate, payments[j].maturityDate);
      if (isFinite(pd) && pd > 0) deltas.push(pd);
    }
    inferredTerm = mode_(deltas.map(function(x){
      var choices = [30, 60, 90], best = 30, bestDiff = 1e9;
      for (var k = 0; k < choices.length; k++) {
        var diff = Math.abs(x - choices[k]);
        if (diff < bestDiff) { bestDiff = diff; best = choices[k]; }
      }
      return best;
    }), 30);
  }

  for (var z = 0; z < invoices.length; z++) invoices[z].term = inferredTerm;

  invoices.sort(function(a,b){ return a.invoiceDate - b.invoiceDate; });
  payments.sort(function(a,b){ return a.paymentDate - b.paymentDate; });

  return { invoices: invoices, payments: payments, firstInvoiceDate: firstInvoiceDate, firstTransactionDate: firstTransactionDate };
}

/** Core analysis & dashboard (FIFO + metrics). beginningBalance applies first. */
function runAnalysisCore_(invoices, payments, startDate, beginningBalance) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();

  // Analysis sheet is a log of calculations (table only)
  var analysis = ss.getSheetByName("Analysis");
  if (!analysis) {
    analysis = ss.insertSheet("Analysis");
  } else {
    // Ensure prior protections do not block editors when clearing/updating
    try {
      var aps = analysis.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (var ai = 0; ai < aps.length; ai++) {
        if (aps[ai].getDescription && aps[ai].getDescription() === 'AR Tool: Sheet Lock') {
          aps[ai].remove();
        }
      }
    } catch (e) {}
    analysis.clear();
  }

  // Synthetic opening balance invoice
  if (beginningBalance > 0 && startDate) {
    invoices.unshift({
      invoiceDate: new Date(startDate.getTime() - 86400000),
      invoiceNum: "BEGIN BAL",
      type: "Opening",
      amount: beginningBalance,
      remaining: beginningBalance,
      term: 30,
      paid: false,
      closingDate: null,
      _synthetic: true
    });
  }

  // FIFO: apply payments only to invoices dated on/before the payment date
  // Also accumulate global payment handover lag (invoice -> payment) across all payment types
  var allLagTotalAmt = 0, allLagWeightedSum = 0;
  for (var pi = 0; pi < payments.length; pi++) {
    var p = payments[pi];
    var rem = p.amount;
    for (var ii = 0; ii < invoices.length; ii++) {
      var inv = invoices[ii];
      if (rem <= 0) break;
      if (inv.paid) continue;
      if (inv.invoiceDate > p.paymentDate) continue;
      var applied = Math.min(inv.remaining, rem);
      if (applied <= 0) continue;
      inv.remaining -= applied;
      rem -= applied;
      // Global handover lag across all payment types (amount-weighted)
      var lagAll = Math.max(0, Math.round((p.paymentDate - inv.invoiceDate)/86400000));
      allLagTotalAmt += applied;
      allLagWeightedSum += lagAll * applied;
      // Track applied payment term for per-invoice term inference later
      if (p.expectedTerm != null) {
        if (!inv._appliedTerms) inv._appliedTerms = [];
        inv._appliedTerms.push(p.expectedTerm);
      }
      // Track all applied payments for settlement-basis metrics
      if (!inv._appliedPays) inv._appliedPays = [];
      inv._appliedPays.push({ amount: applied, invoiceDate: inv.invoiceDate, paymentDate: p.paymentDate, maturityDate: p.maturityDate || null, payType: p.payType || '' });
      // Track applied checks for check-specific metrics (handover lag and maturity duration)
      if (p.payType === 'Check') {
        if (!inv._appliedChecks) inv._appliedChecks = [];
        inv._appliedChecks.push({
          amount: applied,
          invoiceDate: inv.invoiceDate,
          paymentDate: p.paymentDate,
          maturityDate: p.maturityDate || null
        });
      }
      if (inv.remaining === 0) { inv.paid = true; inv.closingDate = p.paymentDate; }
    }
  }

  // After applying payments, set each invoice's term to the mode of applied payment terms (if any)
  for (var iti = 0; iti < invoices.length; iti++) {
    var invt = invoices[iti];
    if (invt._synthetic) continue;
    if (invt._appliedTerms && invt._appliedTerms.length) {
      invt.term = mode_(invt._appliedTerms, invt.term);
    }
  }

  // Build Analysis table (exclude synthetic invoices)
  var headers = ["Invoice Date","Invoice No","Type","Amount","Closing Date","Term (Days)","Due Date","Days to Pay","Days After Due","Remaining"];
  analysis.getRange(1,1,1,headers.length).setValues([headers]);
  var displayInvoices = invoices.filter(function(inv){ return !inv._synthetic; });
  displayInvoices.sort(function(a,b){ return a.invoiceDate - b.invoiceDate; });
  var rows = [];
  for (var k = 0; k < displayInvoices.length; k++) {
    var inv = displayInvoices[k];
    var daysToPay = (inv.paid && inv.closingDate) ? Math.round((inv.closingDate - inv.invoiceDate)/86400000) : "";
    if (typeof daysToPay === 'number' && daysToPay < 0) daysToPay = "";
    var dueDate = new Date(inv.invoiceDate.getTime() + inv.term*86400000);
    var daysAfterDue = (inv.paid && inv.closingDate && typeof daysToPay === 'number') ? (daysToPay - inv.term) : "";
    rows.push([inv.invoiceDate, inv.invoiceNum || "", inv.type || "", inv.amount, inv.paid ? inv.closingDate : "", inv.term, inv.paid ? dueDate : "", daysToPay, daysAfterDue, inv.remaining]);
  }
  if (rows.length) analysis.getRange(2,1,rows.length,headers.length).setValues(rows);
  // Highlight late payments
  for (var r = 0; r < rows.length; r++) {
    var val = rows[r][8];
    var cell = analysis.getRange(r+2,9);
    if (typeof val === 'number' && val > 0) cell.setBackground("#ffcccc"); else cell.setBackground(null);
  }

  // Metrics
  var paid = displayInvoices.filter(function(inv){ return inv.paid && inv.closingDate; });
  var unpaid = displayInvoices.filter(function(inv){ return inv.remaining > 0; });
  // Overdue by handover rule: unpaid invoices older than 30 days from invoice date
  var overdueUnpaidByHandover = unpaid.filter(function(inv){ return ((today - inv.invoiceDate)/86400000) > 30; });

  // Average Days to Pay based on payment handover date across all payments (amount-weighted)
  var avgPaymentLagDays = (allLagTotalAmt > 0) ? (allLagWeightedSum / allLagTotalAmt) : "";
  // Amount-weighted average age of unpaid invoices
  var sumAgeWeighted = 0, sumRemaining = 0;
  for (var wau = 0; wau < unpaid.length; wau++) {
    var invu = unpaid[wau];
    var ageDays = (today - invu.invoiceDate) / 86400000;
    if (isFinite(ageDays) && isFinite(invu.remaining) && invu.remaining > 0) {
      sumAgeWeighted += ageDays * invu.remaining;
      sumRemaining += invu.remaining;
    }
  }
  var avgAgeUnpaid = (sumRemaining > 0) ? (sumAgeWeighted / sumRemaining) : "";
  // Percent as a fraction (0..1); now amount-weighted by unpaid balances (30-day handover rule)
  var overdueUnpaidAmt = 0;
  for (var ou = 0; ou < unpaid.length; ou++) {
    var invou = unpaid[ou];
    var ageou = Math.floor((today - invou.invoiceDate)/86400000);
    if (ageou > 30) overdueUnpaidAmt += invou.remaining;
  }
  var overdueRate = (sumRemaining > 0) ? (overdueUnpaidAmt / sumRemaining) : "";
  var blendedDaysToPay = displayInvoices.length ? displayInvoices.reduce(function(s,inv){ var end = (inv.paid && inv.closingDate) ? inv.closingDate : today; return s + ((end - inv.invoiceDate)/86400000); },0)/displayInvoices.length : "";

  // Per-payment expected vs maturity
  var maturitySamples = [];
  for (var cj = 0; cj < payments.length; cj++) {
    var pmt = payments[cj];
    if (pmt.maturityDate) {
      var d = daysBetween_(pmt.paymentDate, pmt.maturityDate);
      if (isFinite(d) && d > 0) maturitySamples.push({ days: d, expected: (pmt.expectedTerm != null ? pmt.expectedTerm : 30) });
    }
  }
  var avgMaturityDuration = maturitySamples.length ? maturitySamples.reduce(function(s,m){ return s + m.days; },0)/maturitySamples.length : "";
  var avgOverBy = maturitySamples.length ? maturitySamples.reduce(function(s,m){ return s + (m.days - m.expected); },0)/maturitySamples.length : "";
  // Percent as a fraction (0..1)
  var pctChecksOverTerm = maturitySamples.length ? (maturitySamples.filter(function(m){ return m.days > m.expected; }).length / maturitySamples.length) : "";

  // Avg Monthly Purchases
  var totalInvoicedInPeriod = displayInvoices.reduce(function(s,inv){ return s + inv.amount; },0);
  var monthsInPeriod = (today - startDate)/(86400000*30.44);
  var avgMonthlyPurchases = monthsInPeriod > 0 ? (totalInvoicedInPeriod / monthsInPeriod) : "";

  // Check-specific metrics: handover lag and maturity duration based on invoice date
  var checkTotalAmt = 0, checkLagWeightedSum = 0, checkWithin30Amt = 0;
  var checkMatTotalAmt = 0, checkMatWeightedSum = 0;
  for (var ci = 0; ci < invoices.length; ci++) {
    var invc = invoices[ci];
    if (!invc || !invc._appliedChecks) continue;
    for (var ac = 0; ac < invc._appliedChecks.length; ac++) {
      var app = invc._appliedChecks[ac];
      if (!(app && app.invoiceDate && app.paymentDate)) continue;
      var lagDays = Math.max(0, Math.round((app.paymentDate - app.invoiceDate)/86400000));
      checkTotalAmt += app.amount;
      checkLagWeightedSum += lagDays * app.amount;
      if (lagDays <= 30) checkWithin30Amt += app.amount;
      if (app.maturityDate) {
        var matDur = Math.round((app.maturityDate - app.invoiceDate)/86400000);
        if (isFinite(matDur) && matDur > 0) {
          checkMatTotalAmt += app.amount;
          checkMatWeightedSum += matDur * app.amount;
        }
      }
    }
  }
  var avgCheckHandoverLag = (checkTotalAmt > 0) ? (checkLagWeightedSum / checkTotalAmt) : "";
  // fraction 0..1
  var pctChecksHandedOver30 = (checkTotalAmt > 0) ? (((checkTotalAmt - checkWithin30Amt) / checkTotalAmt)) : "";
  var avgCheckMaturityDuration = (checkMatTotalAmt > 0) ? (checkMatWeightedSum / checkMatTotalAmt) : "";
  var avgCheckMaturityOverBy = (avgCheckMaturityDuration !== "") ? (avgCheckMaturityDuration - 90) : "";

  // Settlement-basis (clearing) metrics
  function _settlementDateFor(inv) {
    if (!(inv && inv.paid)) return null;
    var cands = [];
    if (inv._appliedPays && inv._appliedPays.length) {
      for (var ii = 0; ii < inv._appliedPays.length; ii++) {
        var ap = inv._appliedPays[ii];
        if (ap && ap.payType === 'Check') { cands.push(ap.maturityDate || ap.paymentDate); }
        else if (ap) { cands.push(ap.paymentDate); }
      }
    } else if (inv.closingDate) {
      cands.push(inv.closingDate);
    }
    if (!cands.length) return null;
    var max = cands[0];
    for (var jj = 1; jj < cands.length; jj++) { if (cands[jj] > max) max = cands[jj]; }
    return max;
  }
  var _settled = [];
  for (var pi = 0; pi < paid.length; pi++) {
    var invp = paid[pi];
    var sd = _settlementDateFor(invp);
    if (sd) _settled.push({ inv: invp, sd: sd });
  }
  var avgDaysToSettle = _settled.length ? (_settled.reduce(function(s,x){ return s + Math.round((x.sd - x.inv.invoiceDate)/86400000); },0) / _settled.length) : "";
  var pctInvoicesSettledAfterTerm = _settled.length ? (_settled.filter(function(x){ return Math.round((x.sd - x.inv.invoiceDate)/86400000) > x.inv.term; }).length / _settled.length) : "";

  // Dashboard Avg Days to Pay: use payment handover lag (consistent for all types)
  var avgDaysToPayForDash = avgPaymentLagDays;

  // Assessments (computed after all dependent values are defined)
  // Align assessments with scoring thresholds
  var assAvgDaysToPay        = assessLowerBetter_(avgDaysToPayForDash, 20, 40);
  var assAvgAgeUnpaid        = assessLowerBetter_(avgAgeUnpaid, 10, 20);
  // thresholds as fractions: 10% and 30%
  var assOverdueRate         = assessLowerBetter_(overdueRate, 0.10, 0.30);
  var assBlendedDaysToPay    = assessLowerBetter_(blendedDaysToPay, 15, 50);
  var assAvgMaturity         = (avgCheckMaturityOverBy === "") ? "" : assessLowerBetter_(avgCheckMaturityOverBy, 0, 30);
  var assPctChecksOverTerm   = assessLowerBetter_(pctChecksOverTerm, 0.30, 0.60);

  // Risk and Credit Limit (weighted scoring 0..1)
  function compLowerBetter(val, goodMax, avgMax) {
    if (val === "") return null;
    return (val <= goodMax) ? 1 : (val <= avgMax) ? 0.5 : 0;
  }
  var weightedSum = 0, weightTotal = 0;
  function addWeighted(comp, weight) {
    if (comp == null) return;
    weightedSum += comp * weight;
    weightTotal += weight;
  }
  // Weights
  var wAvgDays = 0.20;
  var wAgeUnpaid = 0.10;
  var wOverdue = 0.10;
  var wBlended = 0.20;
  var wMaturityOver = 0.20;
  var wOverTerm = 0.20;
  // Components
  addWeighted(compLowerBetter(avgDaysToPayForDash, 20, 40), wAvgDays);
  addWeighted(compLowerBetter(avgAgeUnpaid, 10, 20), wAgeUnpaid);
  addWeighted(compLowerBetter(overdueRate, 0.10, 0.30), wOverdue); // fraction thresholds
  addWeighted(compLowerBetter(blendedDaysToPay, 15, 50), wBlended);
  addWeighted(compLowerBetter(avgCheckMaturityOverBy, 0, 30), wMaturityOver);
  addWeighted(compLowerBetter(pctChecksOverTerm, 0.30, 0.60), wOverTerm); // fraction thresholds
  var normalizedScore = (weightTotal > 0) ? (weightedSum / weightTotal) : 0;
  // Map to bands similar to old 0-4/5-8/9-12 cutoffs (approx 0.33, 0.66)
  var riskBand = (normalizedScore <= 0.3333) ? "Poor" : (normalizedScore <= 0.6667) ? "Average" : "Good";
  var mostCommonTerm = maturitySamples.length ? mode_(maturitySamples.map(function(m){ return m.expected; }), 30) : 30;
  var baseMult = (mostCommonTerm === 90)
    ? (riskBand === "Good" ? 3.5 : riskBand === "Average" ? 3.25 : 3.0)
    : (riskBand === "Good" ? 2.0 : riskBand === "Average" ? 1.75 : 1.0);
  var creditLimit = (avgMonthlyPurchases !== "") ? (avgMonthlyPurchases * baseMult) : "";

  // Build Dashboard with metrics, aging, and DSO trend
  var dashName = "Credit Risk Dashboard";
  var dash = ss.getSheetByName(dashName);
  if (!dash) {
    dash = ss.insertSheet(dashName);
  } else {
    // Ensure prior protections do not block editors when clearing/updating
    try {
      var dps = dash.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (var di = 0; di < dps.length; di++) {
        if (dps[di].getDescription && dps[di].getDescription() === 'AR Tool: Sheet Lock') {
          dps[di].remove();
        }
      }
    } catch (e) {}
    // Remove any existing embedded charts to avoid duplicates across runs
    try {
      var existingCharts = dash.getCharts();
      for (var c = 0; c < existingCharts.length; c++) {
        dash.removeChart(existingCharts[c]);
      }
    } catch (e) {}
    dash.clear();
  }

  // Metrics table with headers at A1:C
  var metricsHeader = [["Metric","Value","Assessment"]];
  var metricsRows = [
    ["Beginning Balance (TRY)",                        beginningBalance,            ""],
    ["Average Days to Pay (Handover)",                avgDaysToPayForDash,         assAvgDaysToPay],
    ["Weighted Avg Age of Unpaid Invoices (Days)",     avgAgeUnpaid,                assAvgAgeUnpaid],
    ["% of Unpaid Invoices Overdue",                   overdueRate,                 assOverdueRate],
    ["Blended Average Days to Pay",                    blendedDaysToPay,            assBlendedDaysToPay],
    ["Average Monthly Purchases (TRY)",                avgMonthlyPurchases,         ""],
    ["Average Check Maturity Duration (Invoice→Maturity)",         avgCheckMaturityDuration,    ""],
    ["Avg Check Maturity Over Expected (Days)",                    avgCheckMaturityOverBy,      assAvgMaturity],
    ["% of Checks Over Expected Term (Handover→Maturity)",        pctChecksOverTerm,           assPctChecksOverTerm],
    ["% of Checks Handed Over >30 Days (Invoice→Handover)",       pctChecksHandedOver30,       ""],
    ["Average Days to Settle (Settlement)",                         (settledPaid.length ? Math.round(avgDaysToSettle) : ""),             ""],
    ["% of Invoices Settled After Term (Settlement)",              pctInvoicesSettledAfterTerm, ""],
    ["Customer Risk Rating",                           riskBand,                    ""],
    ["Credit Limit (TRY)",                             creditLimit,                 ""]
  ];

  // Round day-based metrics to whole days and remove thousands separators via number format later
  var dayLabels = {
    'Average Days to Pay (Handover)': true,
    'Weighted Avg Age of Unpaid Invoices (Days)': true,
    'Blended Average Days to Pay': true,
    'Average Check Maturity Duration (Invoice→Maturity)': true,
    'Avg Check Maturity Over Expected (Days)': true,
    'Average Days to Settle (Settlement)': true
  };
  for (var mridx = 0; mridx < metricsRows.length; mridx++) {
    var lbl = metricsRows[mridx][0];
    var val = metricsRows[mridx][1];
    if (dayLabels[lbl] && typeof val === 'number' && isFinite(val)) {
      metricsRows[mridx][1] = Math.round(val);
    }
  }
  dash.getRange(1,1,1,3).setValues(metricsHeader).setFontWeight("bold");
  if (metricsRows.length) dash.getRange(2,1,metricsRows.length,3).setValues(metricsRows);

  // Color assessment chips in column C
  for (var i = 0; i < metricsRows.length; i++) {
    var val = metricsRows[i][2];
    var cell = dash.getRange(2+i, 3);
    var color = null;
    if (val === 'Good') color = '#c6efce';
    else if (val === 'Average') color = '#ffe6cc';
    else if (val === 'Poor') color = '#f4a7a7'; // more red for Poor
    if (color) cell.setBackground(color); else cell.setBackground(null);
  }

  // Metric notes (Arabic + English)
  var notesMap = {
    'Beginning Balance (TRY)': 'AR: رصيد البداية لمستحقات العميل قبل فترة التحليل.\nEN: Opening receivables balance before the analysis period.',
    'Average Days to Pay (Paid Only)': 'AR: متوسط أيام الدفع = متوسط الأيام من تاريخ الفاتورة إلى تاريخ الدفع (تسليم الشيك/الدفع). مرجّح بالمبلغ.\nEN: Average days from invoice date to payment (handover). Amount-weighted across payments.',
    'Weighted Avg Age of Unpaid Invoices (Days)': 'AR: متوسط عمر الديون غير المسددة مرجّحًا بالمبالغ = مجموع (العمر بالأيام × الرصيد) ÷ مجموع الأرصدة.\nEN: Amount-weighted average age of unpaid invoices: sum(age_days × remaining) / sum(remaining).',
    '% of Unpaid Invoices Overdue': 'AR: نسبة الفواتير غير المسددة التي تجاوز عمرها 30 يومًا من تاريخ الفاتورة.\nEN: Share of unpaid invoices older than 30 days since invoice date (handover basis).',
    'Blended Average Days to Pay': 'AR: متوسط عمر الفواتير الكلي: المدفوعة حتى تاريخ الإقفال وغير المدفوعة حتى اليوم.\nEN: Average age over all invoices: paid (invoice→closing), unpaid (invoice→today).',
    'Average Monthly Purchases (TRY)': 'AR: متوسط المشتريات الشهري = إجمالي الفواتير ÷ عدد الأشهر في الفترة.\nEN: Total invoiced amount divided by months in the period.',
    'Average Check Maturity Duration (Days)': 'AR: للشيكات فقط: متوسط الأيام من تاريخ الفاتورة إلى استحقاق الشيك (مرجّح بالمبلغ).\nEN: Checks only: average days from invoice date to check maturity (amount-weighted).',
    'Avg Maturity Over By (Days)': 'AR: للشيكات فقط: متوسط (مدة الاستحقاق − 90).\nEN: Checks only: average (maturity duration − 90).',
    '% of Payments Over Term': 'AR: نسبة الدفعات التي تجاوزت مدة الاستحقاق المتوقعة (مثلاً 30/90).\nEN: Share of payments where maturity duration exceeds the expected term.',
    '% Checks Handed over 30 Days': 'AR: نسبة مبالغ الشيكات المسلَّمة بعد أكثر من 30 يومًا من تاريخ الفاتورة.\nEN: Amount share of checks handed over more than 30 days after invoice date.',
    'Customer Risk Rating': 'AR: تقييم مركب مرجّح بناءً على المقاييس الرئيسية (20/10/10/20/20/20%).\nEN: Weighted composite rating based on key metrics (20/10/10/20/20/20%).',
    'Credit Limit (TRY)': 'AR: حد الائتمان = متوسط المشتريات الشهري × معامل حسب التصنيف والمدة.\nEN: Credit limit = average monthly purchases × multiplier by risk band and term.'
  };
  for (var nr = 0; nr < metricsRows.length; nr++) {
    var labelN = metricsRows[nr][0];
    var note = notesMap[labelN];
    if (note) {
      dash.getRange(2+nr, 1).setNote(note);
      dash.getRange(2+nr, 2).setNote(note);
    }
  }

  // Currency and percent formatting in dashboard for specific rows by label
  var currencyFormat = "[$TRY] #,##0.00";
  for (var mr = 0; mr < metricsRows.length; mr++) {
    var label = metricsRows[mr][0];
    if (label === 'Beginning Balance (TRY)' || label === 'Average Monthly Purchases (TRY)' || label === 'Credit Limit (TRY)') {
      dash.getRange(2+mr, 2).setNumberFormat(currencyFormat);
    }
    if (label === '% of Unpaid Invoices Overdue' || label === '% of Checks Over Expected Term (Handover→Maturity)' || label === '% of Checks Handed Over >30 Days (Invoice→Handover)' || label === '% of Invoices Settled After Term (Settlement)') {
      dash.getRange(2+mr, 2).setNumberFormat("0.0% ");
    }
    if (dayLabels[label]) {
      dash.getRange(2+mr, 2).setNumberFormat("0"); // whole days, no thousands separator
    }
  }

  // Color the Customer Risk Rating value cell using the same scheme
  for (var rr = 0; rr < metricsRows.length; rr++) {
    if (metricsRows[rr][0] === 'Customer Risk Rating') {
      var rb = riskBand;
      var cellB = dash.getRange(2+rr, 2);
      var rbColor = null;
      if (rb === 'Good') rbColor = '#c6efce';
      else if (rb === 'Average') rbColor = '#ffe6cc';
      else if (rb === 'Poor') rbColor = '#f4a7a7';
      if (rbColor) cellB.setBackground(rbColor); else cellB.setBackground(null);
      break;
    }
  }

  // Aging buckets table at E1:F5
  var agingLabels = ["0-30 days","31-60 days","61-90 days","91+ days"], sums = [0,0,0,0];
  for (var u = 0; u < unpaid.length; u++) {
    var invu = unpaid[u], age = Math.floor((today - invu.invoiceDate)/86400000);
    if (age <= 30) sums[0] += invu.remaining; else if (age <= 60) sums[1] += invu.remaining; else if (age <= 90) sums[2] += invu.remaining; else sums[3] += invu.remaining;
  }
  dash.getRange(1,5,1,2).setValues([["Aging Bucket","Amount Outstanding"]]).setFontWeight("bold");
  if (agingLabels.length) dash.getRange(2,5,agingLabels.length,2).setValues(agingLabels.map(function(lbl,i){ return [lbl, sums[i]]; }));
  dash.getRange(2,6,agingLabels.length,1).setNumberFormat(currencyFormat);

  // DSO trend (monthly average days to pay) at H1:I*
  var monthMap = {};
  for (var p = 0; p < paid.length; p++) {
    var inv = paid[p];
    var dt = new Date(inv.closingDate.getFullYear(), inv.closingDate.getMonth(), 1);
    var key = dt.getFullYear() + '-' + (dt.getMonth()+1);
    var d2p = Math.round((inv.closingDate - inv.invoiceDate)/86400000);
    if (!isFinite(d2p) || d2p < 0) continue;
    if (!monthMap[key]) monthMap[key] = { dt: dt, sum: 0, cnt: 0 };
    monthMap[key].sum += d2p; monthMap[key].cnt += 1;
  }
  var months = Object.keys(monthMap).map(function(k){ return monthMap[k]; }).sort(function(a,b){ return a.dt - b.dt; });
  dash.getRange(1,8,1,2).setValues([["Month","Avg Days to Pay"]]).setFontWeight("bold");
  if (months.length) {
    var dsoRows = months.map(function(m){ return [m.dt, m.sum/m.cnt]; });
    dash.getRange(2,8,dsoRows.length,2).setValues(dsoRows);
    dash.getRange(2,8,dsoRows.length,1).setNumberFormat("mmm yyyy");
  }

  // Charts
  var agingChart = dash.newChart().setChartType(Charts.ChartType.BAR)
    .addRange(dash.getRange("E2:F5"))
    .setPosition(14,1,0,0)
    .setOption('title','Unpaid Invoices by Aging Bucket')
    .setOption('legend', { position: 'none' })
    .setOption('vAxis', { format: '#,##0.00' })
    .build();
  dash.insertChart(agingChart);

  if (months.length) {
    var dsoChart = dash.newChart().setChartType(Charts.ChartType.LINE)
      .addRange(dash.getRange(1,8,months.length+1,2))
      .setPosition(14,5,0,0)
      .setOption('title','DSO / Days to Pay Trend')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { format: 'MMM yyyy' })
      .build();
    dash.insertChart(dsoChart);
  }
  
}

/** ENTRY POINT */
function runAnalysisBSInput() {
  var bb = readBeginningBalance_();
  var parsed = readInput_();
  if (!parsed.invoices.length && !parsed.payments.length) {
    SpreadsheetApp.getUi().alert("No parsable rows in 'Input'. Check headers and data.");
    return;
  }
  var startDate = parsed.firstTransactionDate || parsed.firstInvoiceDate;
  if (!startDate) {
    SpreadsheetApp.getUi().alert("No dated rows (invoice or payment) found to determine the start date.");
    return;
  }
  runAnalysisCore_(parsed.invoices, parsed.payments, startDate, bb);
}

// Backwards-compat alias for older triggers/menus
function runAnalysisUniversal() {
  return runAnalysisBSInput();
}

/** Menu */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("AR Analysis Tool")
    .addItem("Run Analysis (Input)", "runAnalysisBSInput")
    .addItem("Show Input Mapping (Debug)", "showInputDebug")
    .addToUi();
}

/** DEBUG: Show detected column mapping and sample parsing */
function showInputDebug() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Input") || ss.getSheetByName("input") || ss.getSheetByName("BS Input") || ss.getSheetByName("BS input") || ss.getSheetByName("BSInput") || ss.getSheetByName("Bs Input") || ss.getSheetByName("Universal");
  if (!sh) {
    SpreadsheetApp.getUi().alert('Input sheet not found. Expected one of: Input, BS Input, BS input, BSInput, Bs Input, Universal');
    return;
  }
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data rows to debug (need at least 1 row under headers).');
    return;
  }

  var headersRaw = sh.getRange(1,1,1,lastCol).getValues()[0];
  var headers = headersRaw.map(function(h){ return String(h || '').trim().toLowerCase(); });

  var creditNames = ['credit','alacak','alacaklar','invoice','fatura','fatura tutari','fatura miktari'];
  var debitNames  = ['debit','bor�','borclar','payment','�deme','tahsilat','odeme'];
  var descNames   = ['description','a�iklama','aciklama','desc','not','memo'];
  var dateNames   = ['date','tarih'];
  var payTypeNames= ['pay type','payment type','odeme tipi','�deme tipi','odeme turu','�deme t�r�','tahsilat tipi','paytype'];

  var cCredit = findCol_(headers, creditNames);
  var cDebit  = findCol_(headers, debitNames);
  var cDesc   = findCol_(headers, descNames);
  var cDate   = findCol_(headers, dateNames);
  var cPayTp  = findCol_(headers, payTypeNames);

  var parsed = readInput_();
  var bb = readBeginningBalance_();

  var dbg = ss.getSheetByName('Debug Mapping');
  if (!dbg) dbg = ss.insertSheet('Debug Mapping'); else dbg.clear();

  var now = new Date();
  var info = [
    ['Debug Timestamp', now],
    ['Detected Input Sheet', sh.getName()],
    ['Headers (lowercased)', headers.join(' | ')],
    ['Credit col (1-based)', cCredit],
    ['Debit col (1-based)', cDebit],
    ['Description col (1-based)', cDesc],
    ['Date col (1-based)', cDate],
    ['Pay Type col (1-based)', cPayTp],
    ['Beginning Balance Used', bb],
    ['Invoices Parsed', parsed.invoices.length],
    ['Payments Parsed', parsed.payments.length],
    ['First Invoice Date', parsed.firstInvoiceDate]
  ];
  dbg.getRange(1,1,info.length,2).setValues(info);
  dbg.getRange(1,1,1,2).setFontWeight('bold');

  // Sample rows
  var sampleCount = Math.min(50, Math.max(0, lastRow - 1));
  var values = sh.getRange(2,1,sampleCount,lastCol).getValues();
  var outHdr = [['Row','Raw Date','Parsed Date','Raw Credit','Parsed Credit','Raw Debit','Parsed Debit','Description','Date in Desc','Raw Pay Type','Norm Pay Type','Expected Term']];
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rawDate = (cDate>0) ? row[cDate-1] : '';
    var parsedDate = parseDMY_(rawDate) || extractDateFromText_((cDesc>0)?row[cDesc-1]:'');
    var rawCredit = (cCredit>0) ? row[cCredit-1] : '';
    var rawDebit  = (cDebit>0)  ? row[cDebit-1]  : '';
    var pCredit = amountToNumber_(rawCredit);
    var pDebit  = amountToNumber_(rawDebit);
    var desc    = (cDesc>0) ? row[cDesc-1] : '';
    var descDate= extractDateFromText_(desc);
    var rawPay  = (cPayTp>0) ? row[cPayTp-1] : '';
  var norm    = normalizePayType2_(rawPay);
    out.push([
      i+2,
      rawDate,
      parsedDate || '',
      rawCredit,
      isFinite(pCredit)?pCredit:'',
      rawDebit,
      isFinite(pDebit)?pDebit:'',
      desc,
      descDate || '',
      rawPay,
      norm.type,
      norm.termDays
    ]);
  }
  dbg.getRange(14,1,1,outHdr[0].length).setValues(outHdr).setFontWeight('bold');
  if (out.length) dbg.getRange(15,1,out.length,out[0].length).setValues(out);
  dbg.autoResizeColumns(1, Math.max(2,outHdr[0].length));
  dbg.setFrozenRows(13);
}



