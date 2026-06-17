const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

const COLORS = {
  primary: 'FF5900', accent: '411517', secondary: 'F4EDE3',
  highlight: 'CFFF00', white: 'FFFFFF', black: '000000',
};

const MONTHS_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function excelDateToStr(serial) {
  if (!serial || typeof serial !== 'number') return serial;
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const d = new Date(utcValue * 1000);
  const frac = serial - Math.floor(serial);
  const totalSecs = Math.round(frac * 86400);
  const hrs = Math.floor(totalSecs / 3600);
  const mins = Math.floor((totalSecs % 3600) / 60);
  const dd = String(d.getUTCDate()).padStart(2, '0');
  const mmm = MONTHS_SHORT[d.getUTCMonth()];
  const yyyy = d.getUTCFullYear();
  return `${dd}-${mmm}-${yyyy} ${String(hrs).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
}

function dateOnlyDDMMMYYYY(dateTimeStr) {
  if (!dateTimeStr || typeof dateTimeStr !== 'string') return dateTimeStr || '(unknown)';
  const parts = dateTimeStr.split(' ')[0];
  return parts;
}

function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (inQuotes) {
      if (c === '"' && line[i + 1] === '"') { current += '"'; i++; }
      else if (c === '"') inQuotes = false;
      else current += c;
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ',') { result.push(current); current = ''; }
      else current += c;
    }
  }
  result.push(current);
  return result;
}

function emit(msg) { console.log(msg); }

async function readDetailedReports(baseDir) {
  const detDir = path.join(baseDir, 'Detailed Reports');
  if (!fs.existsSync(detDir)) throw new Error('Detailed Reports folder not found');
  const files = fs.readdirSync(detDir).filter(f => f.endsWith('.xlsx'));
  if (!files.length) throw new Error('No .xlsx files in Detailed Reports');

  const colMap = {
    12: 'Restaurant Id', 13: 'Restaurant', 14: 'Branch Id', 15: 'Branch',
    16: 'Order Id', 18: 'Date / Time', 19: 'Payment Method', 20: 'Status',
    21: 'Service Type', 25: 'SubTotal', 32: 'Total Sale', 64: 'Net Payment Per Order',
  };
  const colNums = Object.keys(colMap).map(Number);
  const records = [];
  const seenOrderIds = new Set();
  let totalDetailedRows = 0;

  for (const file of files) {
    emit(`  Reading: ${file}`);
    const wb = new ExcelJS.stream.xlsx.WorkbookReader(path.join(detDir, file));
    let rowNum = 0;
    for await (const ws of wb) {
      for await (const row of ws) {
        rowNum++;
        if (rowNum <= 3) continue;
        totalDetailedRows++;
        const rec = {};
        row.eachCell({ includeEmpty: true }, (cell, c) => {
          if (colNums.includes(c)) rec[colMap[c]] = cell.value;
        });
        const totalSale = parseFloat(rec['Total Sale']);
        if (isNaN(totalSale) || totalSale < 1000) continue;
        rec['Total Sale'] = totalSale;
        rec['SubTotal'] = parseFloat(rec['SubTotal']) || 0;
        rec['Net Payment Per Order'] = parseFloat(rec['Net Payment Per Order']) || 0;
        rec['Restaurant Id'] = rec['Restaurant Id'] ? String(rec['Restaurant Id']) : '';
        rec['Branch Id'] = rec['Branch Id'] ? String(rec['Branch Id']) : '';
        rec['Order Id'] = rec['Order Id'] ? String(rec['Order Id']) : '';
        rec['Date / Time'] = excelDateToStr(rec['Date / Time']);
        if (seenOrderIds.has(rec['Order Id'])) continue;
        seenOrderIds.add(rec['Order Id']);
        records.push(rec);
      }
    }
    emit(`    Accumulated ${records.length} records (Total Sale >= 1000) out of ${totalDetailedRows} total`);
  }
  return { records, totalDetailedRows };
}

function readBranchInquiry(baseDir) {
  const dir = path.join(baseDir, 'Branch Inquiry');
  if (!fs.existsSync(dir)) return {};
  const files = fs.readdirSync(dir).filter(f => f.endsWith('.xlsx'));
  if (!files.length) return {};
  emit(`  Reading Branch Inquiry: ${files[0]}`);
  const wb = XLSX.readFile(path.join(dir, files[0]));
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const lookup = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;
    const branchId = String(row[6] || '');
    if (branchId) {
      lookup[branchId] = {
        'Customer account': String(row[0] || ''),
        'Customer group': String(row[2] || ''),
      };
    }
  }
  emit(`    ${Object.keys(lookup).length} branch entries`);
  return lookup;
}

function readAX365(baseDir) {
  const dir = path.join(baseDir, 'AX365');
  if (!fs.existsSync(dir)) return {};
  const files = fs.readdirSync(dir).filter(f => f.endsWith('.xlsx'));
  if (!files.length) return {};
  emit(`  Reading AX365: ${files[0]}`);
  const wb = XLSX.readFile(path.join(dir, files[0]));
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const lookup = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;
    const branchId = String(row[4] || '');
    if (branchId) lookup[branchId] = { 'Branch AX Code': String(row[6] || '') };
  }
  emit(`    ${Object.keys(lookup).length} AX entries`);
  return lookup;
}

async function readCheckoutForOrders(baseDir, orderIds) {
  const dir = path.join(baseDir, 'Checkout');
  if (!fs.existsSync(dir)) return {};
  const files = fs.readdirSync(dir).filter(f => f.endsWith('.csv'));
  const lookup = {};
  const orderIdSet = new Set(orderIds);
  for (const file of files) {
    emit(`  Checkout (orders): ${file}`);
    let headers = null, refIdx, actionIdx, respIdx, emailIdx, amountIdx;
    let lineCount = 0, matchCount = 0;
    const rl = readline.createInterface({
      input: fs.createReadStream(path.join(dir, file), { encoding: 'utf-8' }),
      crlfDelay: Infinity,
    });
    for await (const line of rl) {
      lineCount++;
      const fields = parseCSVLine(line);
      if (!headers) {
        headers = fields;
        refIdx = headers.indexOf('Reference');
        actionIdx = headers.indexOf('Action Type');
        respIdx = headers.indexOf('Response Description');
        emailIdx = headers.indexOf('Customer Email');
        amountIdx = headers.indexOf('Amount');
        continue;
      }
      const ref = (fields[refIdx] || '').trim();
      if (!ref || !orderIdSet.has(ref)) continue;
      const actionType = (fields[actionIdx] || '').trim();
      if (actionType && actionType !== 'Authorisation') continue;
      const respDesc = (fields[respIdx] || '').trim();
      if (respDesc && respDesc !== 'Approved') continue;
      const amount = parseFloat(fields[amountIdx]) || 0;
      const email = (fields[emailIdx] || '').trim();
      if (!lookup[ref]) { lookup[ref] = { actionType, respDesc, email, amount }; matchCount++; }
      else if (actionType === 'Authorisation') { lookup[ref] = { actionType, respDesc, email, amount }; }
    }
    emit(`    ${lineCount} lines, ${matchCount} matches`);
  }
  return lookup;
}

async function readCheckoutEmailStats(baseDir) {
  const dir = path.join(baseDir, 'Checkout');
  if (!fs.existsSync(dir)) return [];
  const files = fs.readdirSync(dir).filter(f => f.endsWith('.csv'));
  const emailData = {};

  for (const file of files) {
    emit(`  Checkout (emails): ${file}`);
    let headers = null, emailIdx, dateIdx, actionIdx, amountIdx;
    let lineCount = 0;
    const rl = readline.createInterface({
      input: fs.createReadStream(path.join(dir, file), { encoding: 'utf-8' }),
      crlfDelay: Infinity,
    });
    for await (const line of rl) {
      lineCount++;
      const fields = parseCSVLine(line);
      if (!headers) {
        headers = fields;
        emailIdx = headers.indexOf('Customer Email');
        dateIdx = headers.indexOf('Action Date UTC');
        actionIdx = headers.indexOf('Action Type');
        amountIdx = headers.indexOf('Amount');
        continue;
      }
      const email = (fields[emailIdx] || '').trim();
      if (!email) continue;
      const actionType = (fields[actionIdx] || '').trim();
      if (actionType !== 'Authorisation') continue;
      const dateStr = (fields[dateIdx] || '').trim();
      const amount = parseFloat(fields[amountIdx]) || 0;
      if (!emailData[email]) emailData[email] = { count: 0, totalAmount: 0, timestamps: [] };
      emailData[email].count++;
      emailData[email].totalAmount += amount;
      if (dateStr) emailData[email].timestamps.push(dateStr);
    }
    emit(`    ${lineCount} lines scanned`);
  }

  const result = Object.entries(emailData)
    .filter(([_, d]) => d.count > 1)
    .map(([email, d]) => {
      d.timestamps.sort();
      let consecutiveCount = 0;
      let maxBurst = 1, currentBurst = 1;
      for (let i = 1; i < d.timestamps.length; i++) {
        const t1 = new Date(d.timestamps[i - 1]).getTime();
        const t2 = new Date(d.timestamps[i]).getTime();
        const diffMins = (t2 - t1) / 60000;
        if (diffMins <= 60) { currentBurst++; consecutiveCount++; }
        else { maxBurst = Math.max(maxBurst, currentBurst); currentBurst = 1; }
      }
      maxBurst = Math.max(maxBurst, currentBurst);
      const suspicionScore = d.count * 2 + consecutiveCount * 5 + maxBurst * 3;
      return {
        email, count: d.count, totalAmount: Math.round(d.totalAmount * 100) / 100,
        consecutiveCount, maxBurst,
        firstSeen: d.timestamps[0] || '', lastSeen: d.timestamps[d.timestamps.length - 1] || '',
        suspicionScore,
      };
    })
    .sort((a, b) => b.suspicionScore - a.suspicionScore);

  return result;
}

function mergeData(records, branchInq, ax365, checkout) {
  for (const rec of records) {
    const bid = rec['Branch Id'], oid = rec['Order Id'];
    const bi = branchInq[bid] || {};
    rec['Customer group'] = bi['Customer group'] || '';
    rec['Customer account'] = bi['Customer account'] || '';
    const ax = ax365[bid] || {};
    rec['Branch AX Code'] = ax['Branch AX Code'] || '';
    rec['MSD Checker'] = rec['Customer account'] && rec['Branch AX Code']
      ? (rec['Customer account'] === rec['Branch AX Code'] ? 'True' : 'False') : '';
    const co = checkout[oid] || {};
    rec['Action Type'] = co.actionType || '';
    rec['Response Description'] = co.respDesc || '';
    rec['Customer Email'] = co.email || '';
    rec['Checkout Amount'] = co.amount || null;
  }
  return records;
}

function applyFraudRules(records) {
  const dupByRestaurantId = {}, dupByBranchId = {}, dupByCustGroup = {}, dupByEmail = {}, emailCounts = {};
  for (const rec of records) {
    const ts = rec['Total Sale'], rid = rec['Restaurant Id'], bid = rec['Branch Id'];
    const cg = rec['Customer group'], em = rec['Customer Email'];
    dupByRestaurantId[`${rid}_${ts}`] = (dupByRestaurantId[`${rid}_${ts}`] || 0) + 1;
    dupByBranchId[`${bid}_${ts}`] = (dupByBranchId[`${bid}_${ts}`] || 0) + 1;
    if (cg) dupByCustGroup[`${cg}_${ts}`] = (dupByCustGroup[`${cg}_${ts}`] || 0) + 1;
    if (em) {
      dupByEmail[`${em}_${ts}`] = (dupByEmail[`${em}_${ts}`] || 0) + 1;
      emailCounts[em] = (emailCounts[em] || 0) + 1;
    }
  }

  for (const rec of records) {
    const reasons = [];
    const ts = rec['Total Sale'], rid = rec['Restaurant Id'], bid = rec['Branch Id'];
    const cg = rec['Customer group'], em = rec['Customer Email'];
    const pm = (rec['Payment Method'] || '').toLowerCase();
    const coAmt = rec['Checkout Amount'];

    if (ts >= 10000) reasons.push(`Total Sale unusually high: AED ${ts.toLocaleString()}`);
    if (dupByRestaurantId[`${rid}_${ts}`] > 1) reasons.push(`Duplicate Total Sale (AED ${ts}) under same Restaurant Id ${rid}`);
    if (dupByBranchId[`${bid}_${ts}`] > 1) reasons.push(`Duplicate Total Sale (AED ${ts}) under same Branch Id ${bid}`);
    if (cg && dupByCustGroup[`${cg}_${ts}`] > 1) reasons.push(`Duplicate Total Sale (AED ${ts}) under same Customer Group ${cg}`);
    if (em && dupByEmail[`${em}_${ts}`] > 1) reasons.push(`Duplicate Total Sale (AED ${ts}) under same Customer Email`);
    if (em && emailCounts[em] > 2) reasons.push(`Customer Email used ${emailCounts[em]} times (suspicious frequency)`);
    if (pm === 'credit card' || pm === 'talabat credit') {
      if (coAmt !== null && coAmt !== undefined) {
        const diff = coAmt - ts;
        if (diff < -200) reasons.push(`Checkout Amount (${coAmt}) minus Total Sale (${ts}) = ${diff.toFixed(2)}, below -200 threshold`);
      }
    }
    if (pm === 'cash') reasons.push('Payment Method is Cash (high-value cash order)');

    rec['Fraud Reasons'] = reasons.join('; ');
    rec['Potential Fraud'] = reasons.length > 0 ? 'Yes' : 'No';
  }
  return records;
}

function applyPotentialFraudFilter(records) {
  const dupByRidEmail = {};
  const emailCounts = {};
  for (const rec of records) {
    if (rec['Potential Fraud'] !== 'Yes') continue;
    const ts = rec['Total Sale'], rid = rec['Restaurant Id'], em = rec['Customer Email'];
    if (rid && em) {
      const key = `${rid}_${em}_${ts}`;
      dupByRidEmail[key] = (dupByRidEmail[key] || 0) + 1;
    }
    if (em) emailCounts[em] = (emailCounts[em] || 0) + 1;
  }

  const potentialFraud = [];
  for (const rec of records) {
    if (rec['Potential Fraud'] !== 'Yes') continue;
    const ts = rec['Total Sale'], rid = rec['Restaurant Id'], em = rec['Customer Email'];
    const reasons = [];

    if (ts >= 10000) reasons.push(`Unusually high amount: AED ${ts.toLocaleString()}`);
    if (em && emailCounts[em] > 1) reasons.push(`Email used ${emailCounts[em]} times across flagged orders`);
    if (rid && em && dupByRidEmail[`${rid}_${em}_${ts}`] > 1) {
      reasons.push(`Duplicate Total Sale (AED ${ts}) with same Restaurant Id ${rid} AND same email ${em}`);
    }

    if (reasons.length > 0) {
      rec['Potential Fraud Reasons'] = reasons.join('; ');
      rec['Potential Fraud Filter'] = 'Yes';
      potentialFraud.push(rec);
    } else {
      rec['Potential Fraud Filter'] = 'No';
      rec['Potential Fraud Reasons'] = '';
    }
  }
  return potentialFraud;
}

function applyHeaderStyle(ws, row, colCount) {
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    cell.font = { bold: true, color: { argb: 'FF' + COLORS.white }, size: 11 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + COLORS.primary } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF' + COLORS.accent } },
      bottom: { style: 'thin', color: { argb: 'FF' + COLORS.accent } },
      left: { style: 'thin', color: { argb: 'FF' + COLORS.accent } },
      right: { style: 'thin', color: { argb: 'FF' + COLORS.accent } },
    };
  }
}

function applyDataRowStyle(ws, row, colCount, rowIdx, isFraud) {
  const bgColor = isFraud ? 'FFFDE8E8' : (rowIdx % 2 === 0 ? 'FF' + COLORS.secondary : 'FF' + COLORS.white);
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    cell.font = { size: 10, color: { argb: 'FF' + COLORS.accent } };
    cell.border = { bottom: { style: 'hair', color: { argb: 'FFCCCCCC' } }, left: { style: 'hair', color: { argb: 'FFCCCCCC' } }, right: { style: 'hair', color: { argb: 'FFCCCCCC' } } };
    cell.alignment = { vertical: 'middle' };
  }
}

async function generateExcel(records, potentialFraud, checkoutEmailTop20, totalDetailedRows, outputPath) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Talabat Fraud Hunter';
  wb.created = new Date();
  const colLetter = (n) => { let s = ''; while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); } return s; };

  // --- Sheet 1: Main Data ---
  const mainCols = [
    'Restaurant Id', 'Branch Id', 'Restaurant', 'Branch', 'Order Id',
    'Date / Time', 'Payment Method', 'Status', 'Service Type',
    'SubTotal', 'Total Sale', 'Net Payment Per Order',
    'Customer group', 'Customer account', 'Branch AX Code', 'MSD Checker',
    'Action Type', 'Response Description', 'Customer Email', 'Checkout Amount',
    'Potential Fraud', 'Fraud Reasons',
  ];
  const ws1 = wb.addWorksheet('Main Data', { properties: { tabColor: { argb: 'FF' + COLORS.primary } } });
  ws1.columns = mainCols.map(h => ({ header: h, key: h, width: h === 'Fraud Reasons' ? 60 : h === 'Customer Email' ? 30 : 18 }));
  applyHeaderStyle(ws1, ws1.getRow(1), mainCols.length);
  ws1.getRow(1).height = 30;
  ws1.autoFilter = { from: 'A1', to: `${colLetter(mainCols.length)}1` };
  records.forEach((rec, i) => {
    const row = ws1.addRow(mainCols.map(c => rec[c] ?? ''));
    applyDataRowStyle(ws1, row, mainCols.length, i, rec['Potential Fraud'] === 'Yes');
  });
  ws1.views = [{ state: 'frozen', ySplit: 1 }];

  // --- Sheet 2: Potential Fraud (sorted: Total Sale desc, Restaurant Id, Customer Email) ---
  potentialFraud.sort((a, b) => {
    const diff = (b['Total Sale'] || 0) - (a['Total Sale'] || 0);
    if (diff !== 0) return diff;
    const ridCmp = (a['Restaurant Id'] || '').localeCompare(b['Restaurant Id'] || '');
    if (ridCmp !== 0) return ridCmp;
    return (a['Customer Email'] || '').localeCompare(b['Customer Email'] || '');
  });
  const fraudCols = [
    'Serial', 'Restaurant Id', 'Branch Id', 'Restaurant', 'Branch', 'Order Id',
    'Date / Time', 'Payment Method', 'Status', 'Service Type',
    'SubTotal', 'Total Sale', 'Net Payment Per Order',
    'Customer Group', 'Customer Account', 'Customer Email', 'Potential Fraud Reasons',
  ];
  const ws2 = wb.addWorksheet('Potential Fraud', { properties: { tabColor: { argb: 'FFFF0000' } } });
  ws2.columns = fraudCols.map(h => ({ header: h, key: h, width: h === 'Potential Fraud Reasons' ? 60 : h === 'Customer Email' ? 30 : 18 }));
  applyHeaderStyle(ws2, ws2.getRow(1), fraudCols.length);
  ws2.getRow(1).height = 30;
  ws2.autoFilter = { from: 'A1', to: `${colLetter(fraudCols.length)}1` };
  potentialFraud.forEach((rec, i) => {
    const row = ws2.addRow([
      i + 1, rec['Restaurant Id'], rec['Branch Id'], rec['Restaurant'], rec['Branch'],
      rec['Order Id'], rec['Date / Time'], rec['Payment Method'] || '', rec['Status'] || '', rec['Service Type'] || '',
      rec['SubTotal'], rec['Total Sale'], rec['Net Payment Per Order'],
      rec['Customer group'], rec['Customer account'], rec['Customer Email'], rec['Potential Fraud Reasons'],
    ]);
    applyDataRowStyle(ws2, row, fraudCols.length, i, true);
  });
  ws2.views = [{ state: 'frozen', ySplit: 1 }];

  // --- Sheet 3: Dashboard 1 ---
  const potentialFraudTotal = potentialFraud.reduce((s, r) => s + (r['Total Sale'] || 0), 0);
  const fraudPct = totalDetailedRows > 0 ? ((potentialFraud.length / totalDetailedRows) * 100).toFixed(2) : '0.00';

  const ws3 = wb.addWorksheet('Dashboard 1', { properties: { tabColor: { argb: 'FF' + COLORS.highlight } } });
  ws3.mergeCells('A1:H1');
  const t3 = ws3.getCell('A1');
  t3.value = 'FRAUD ANALYSIS DASHBOARD';
  t3.font = { bold: true, size: 18, color: { argb: 'FF' + COLORS.white } };
  t3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + COLORS.primary } };
  t3.alignment = { horizontal: 'center', vertical: 'middle' };
  ws3.getRow(1).height = 40;

  ws3.getCell('A3').value = 'SUMMARY';
  ws3.getCell('A3').font = { bold: true, size: 14, color: { argb: 'FF' + COLORS.primary } };
  ws3.mergeCells('A3:H3');
  ws3.getRow(4).values = ['Metric', 'Value'];
  applyHeaderStyle(ws3, ws3.getRow(4), 2);

  const summaryItems = [
    ['Total Orders (Detailed Report)', totalDetailedRows.toLocaleString()],
    ['Potential Fraud Count', potentialFraud.length],
    ['Potential Fraud %', `${fraudPct}%`],
    ['Potential Fraud Amount (AED)', parseFloat(potentialFraudTotal.toFixed(2))],
  ];
  summaryItems.forEach(([label, val], i) => {
    const r = ws3.getRow(5 + i);
    r.values = [label, val];
    r.getCell(1).font = { bold: true, color: { argb: i >= 1 ? 'FFFF0000' : 'FF' + COLORS.accent } };
    const bg = i >= 1 ? 'FFFDE8E8' : 'FF' + COLORS.secondary;
    [1, 2].forEach(c => { r.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }; });
    if (i === 3) r.getCell(2).numFmt = '#,##0.00';
  });
  ws3.getColumn(1).width = 35; ws3.getColumn(2).width = 25;

  // Breakdown helper with % column
  const groupByFieldWithPct = (arr, field, totalCount) => {
    const map = {};
    for (const r of arr) { const k = r[field] || '(blank)'; if (!map[k]) map[k] = { count: 0, total: 0 }; map[k].count++; map[k].total += r['Total Sale'] || 0; }
    return Object.entries(map).sort((a, b) => b[1].total - a[1].total)
      .map(([k, v]) => ({ key: k, count: v.count, pct: ((v.count / totalCount) * 100).toFixed(1), total: Math.round(v.total * 100) / 100 }));
  };

  const pCount = potentialFraud.length || 1;
  const restNameMap = {}, custAcctMap = {}, branchNameMap = {};
  for (const r of records) {
    restNameMap[r['Restaurant Id']] = r['Restaurant'];
    custAcctMap[r['Restaurant Id']] = r['Customer account'] || '';
    branchNameMap[r['Branch Id']] = r['Branch'];
  }

  let cr = 11;

  // Breakdown by Customer Group
  ws3.getCell(`A${cr}`).value = 'BREAKDOWN BY CUSTOMER GROUP';
  ws3.getCell(`A${cr}`).font = { bold: true, size: 13, color: { argb: 'FF' + COLORS.primary } };
  ws3.mergeCells(`A${cr}:H${cr}`); cr++;
  ws3.getRow(cr).values = ['Customer Group', 'Count', '%', 'Total Amount (AED)'];
  applyHeaderStyle(ws3, ws3.getRow(cr), 4); cr++;
  const byCG = groupByFieldWithPct(potentialFraud, 'Customer group', pCount);
  for (const item of byCG) {
    const row = ws3.getRow(cr);
    row.values = [item.key, item.count, `${item.pct}%`, item.total];
    row.getCell(4).numFmt = '#,##0.00';
    applyDataRowStyle(ws3, row, 4, cr, false); cr++;
  }
  cr += 2;

  // Breakdown by Date
  ws3.getCell(`A${cr}`).value = 'BREAKDOWN BY DATE';
  ws3.getCell(`A${cr}`).font = { bold: true, size: 13, color: { argb: 'FF' + COLORS.primary } };
  ws3.mergeCells(`A${cr}:H${cr}`); cr++;
  ws3.getRow(cr).values = ['Date', 'Count', '%', 'Total Amount (AED)'];
  applyHeaderStyle(ws3, ws3.getRow(cr), 4); cr++;
  const byDateMap = {};
  for (const r of potentialFraud) {
    const dt = (r['Date / Time'] || '').substring(0, 10) || '(unknown)';
    if (!byDateMap[dt]) byDateMap[dt] = { count: 0, total: 0 };
    byDateMap[dt].count++; byDateMap[dt].total += r['Total Sale'] || 0;
  }
  const byDate = Object.entries(byDateMap).sort((a, b) => b[1].total - a[1].total)
    .map(([k, v]) => ({ key: k, count: v.count, pct: ((v.count / pCount) * 100).toFixed(1), total: Math.round(v.total * 100) / 100 }));
  for (const item of byDate) {
    const row = ws3.getRow(cr);
    row.values = [item.key, item.count, `${item.pct}%`, item.total];
    row.getCell(4).numFmt = '#,##0.00';
    applyDataRowStyle(ws3, row, 4, cr, false); cr++;
  }
  cr += 2;

  // Breakdown by Restaurant ID + Customer Account + Restaurant
  ws3.getCell(`A${cr}`).value = 'BREAKDOWN BY RESTAURANT ID + CUSTOMER ACCOUNT + RESTAURANT';
  ws3.getCell(`A${cr}`).font = { bold: true, size: 13, color: { argb: 'FF' + COLORS.primary } };
  ws3.mergeCells(`A${cr}:H${cr}`); cr++;
  ws3.getRow(cr).values = ['Restaurant Id', 'Customer Account', 'Restaurant', 'Count', '%', 'Total Amount (AED)'];
  applyHeaderStyle(ws3, ws3.getRow(cr), 6); cr++;
  const byRidMap = {};
  for (const r of potentialFraud) {
    const rid = r['Restaurant Id'] || '';
    const acct = r['Customer account'] || '';
    const rname = r['Restaurant'] || '';
    const key = `${rid}||${acct}||${rname}`;
    if (!byRidMap[key]) byRidMap[key] = { rid, acct, rname, count: 0, total: 0 };
    byRidMap[key].count++; byRidMap[key].total += r['Total Sale'] || 0;
  }
  const byRid = Object.values(byRidMap).sort((a, b) => b.total - a.total);
  for (const item of byRid.slice(0, 50)) {
    const row = ws3.getRow(cr);
    row.values = [item.rid, item.acct, item.rname, item.count, `${((item.count / pCount) * 100).toFixed(1)}%`, Math.round(item.total * 100) / 100];
    row.getCell(6).numFmt = '#,##0.00';
    applyDataRowStyle(ws3, row, 6, cr, false); cr++;
  }
  cr += 2;

  // Breakdown by Branch ID + Branch
  ws3.getCell(`A${cr}`).value = 'BREAKDOWN BY BRANCH ID + BRANCH';
  ws3.getCell(`A${cr}`).font = { bold: true, size: 13, color: { argb: 'FF' + COLORS.primary } };
  ws3.mergeCells(`A${cr}:H${cr}`); cr++;
  ws3.getRow(cr).values = ['Branch Id', 'Branch', 'Count', '%', 'Total Amount (AED)'];
  applyHeaderStyle(ws3, ws3.getRow(cr), 5); cr++;
  const byBidMap = {};
  for (const r of potentialFraud) {
    const bid = r['Branch Id'] || '';
    const bname = r['Branch'] || '';
    const key = `${bid}||${bname}`;
    if (!byBidMap[key]) byBidMap[key] = { bid, bname, count: 0, total: 0 };
    byBidMap[key].count++; byBidMap[key].total += r['Total Sale'] || 0;
  }
  const byBid = Object.values(byBidMap).sort((a, b) => b.total - a.total);
  for (const item of byBid.slice(0, 50)) {
    const row = ws3.getRow(cr);
    row.values = [item.bid, item.bname, item.count, `${((item.count / pCount) * 100).toFixed(1)}%`, Math.round(item.total * 100) / 100];
    row.getCell(5).numFmt = '#,##0.00';
    applyDataRowStyle(ws3, row, 5, cr, false); cr++;
  }

  ws3.getColumn(3).width = 25; ws3.getColumn(4).width = 18; ws3.getColumn(5).width = 22; ws3.getColumn(6).width = 22;

  // --- Sheet 4: Dashboard 2 ---
  const ws4 = wb.addWorksheet('Dashboard 2', { properties: { tabColor: { argb: 'FF' + COLORS.accent } } });
  ws4.mergeCells('A1:G1');
  const t4 = ws4.getCell('A1');
  t4.value = 'TOP 20 SUSPICIOUS EMAILS FROM CHECKOUT DATA';
  t4.font = { bold: true, size: 16, color: { argb: 'FF' + COLORS.white } };
  t4.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + COLORS.primary } };
  t4.alignment = { horizontal: 'center', vertical: 'middle' };
  ws4.getRow(1).height = 40;
  ws4.getRow(3).values = ['Rank', 'Customer Email', 'Order Count', 'Total Amount (AED)', 'Max Consecutive Burst', 'First Seen', 'Last Seen'];
  applyHeaderStyle(ws4, ws4.getRow(3), 7);
  ws4.getColumn(1).width = 8; ws4.getColumn(2).width = 40; ws4.getColumn(3).width = 14;
  ws4.getColumn(4).width = 20; ws4.getColumn(5).width = 22; ws4.getColumn(6).width = 20; ws4.getColumn(7).width = 20;
  checkoutEmailTop20.forEach((e, i) => {
    const row = ws4.getRow(i + 4);
    row.values = [i + 1, e.email, e.count, e.totalAmount, e.maxBurst, e.firstSeen, e.lastSeen];
    row.getCell(4).numFmt = '#,##0.00';
    applyDataRowStyle(ws4, row, 7, i, true);
  });
  ws4.views = [{ state: 'frozen', ySplit: 3 }];

  await wb.xlsx.writeFile(outputPath);
  emit(`  Excel saved: ${outputPath}`);
}

async function runAnalysis(baseDir, outputPath) {
  emit('=== TALABAT FRAUD HUNTER ===');
  emit(`Input directory: ${baseDir}`);
  emit(`Output file: ${outputPath}`);
  const t0 = Date.now();
  const fileErrors = [];

  emit('[1/7] Reading Detailed Reports...');
  let records, totalDetailedRows;
  try {
    const result = await readDetailedReports(baseDir);
    records = result.records;
    totalDetailedRows = result.totalDetailedRows;
  } catch (err) {
    const detDir = path.join(baseDir, 'Detailed Reports');
    const files = fs.existsSync(detDir) ? fs.readdirSync(detDir).filter(f => f.endsWith('.xlsx')) : [];
    fileErrors.push({ step: 'Detailed Reports', folder: detDir, files, error: err.message, fix: 'Ensure at least one .xlsx file exists in the Detailed Reports folder with transaction data starting at row 4.' });
    throw new Error(`Detailed Reports failed: ${err.message}\n\nFile Errors:\n${JSON.stringify(fileErrors, null, 2)}`);
  }

  emit('[2/7] Reading Branch Inquiry...');
  let branchInq = {};
  try {
    branchInq = readBranchInquiry(baseDir);
  } catch (err) {
    fileErrors.push({ step: 'Branch Inquiry', folder: path.join(baseDir, 'Branch Inquiry'), error: err.message, fix: 'Check the .xlsx file in Branch Inquiry folder. Ensure column G (index 6) contains Branch IDs.' });
    emit(`  WARNING: Branch Inquiry failed (${err.message}), continuing without it.`);
  }

  emit('[3/7] Reading AX365...');
  let ax365 = {};
  try {
    ax365 = readAX365(baseDir);
  } catch (err) {
    fileErrors.push({ step: 'AX365', folder: path.join(baseDir, 'AX365'), error: err.message, fix: 'Check the .xlsx file in AX365 folder. Ensure column E (index 4) has Branch IDs and column G (index 6) has AX Codes.' });
    emit(`  WARNING: AX365 failed (${err.message}), continuing without it.`);
  }

  emit('[4/7] Reading Checkout (matching orders)...');
  const orderIds = records.map(r => r['Order Id']);
  let checkout = {};
  try {
    checkout = await readCheckoutForOrders(baseDir, orderIds);
  } catch (err) {
    fileErrors.push({ step: 'Checkout (order matching)', folder: path.join(baseDir, 'Checkout'), error: err.message, fix: 'Ensure .csv files in Checkout folder have headers: Reference, Action Type, Response Description, Customer Email, Amount.' });
    emit(`  WARNING: Checkout order matching failed (${err.message}), continuing without it.`);
  }

  emit('[5/7] Merging and applying fraud rules...');
  mergeData(records, branchInq, ax365, checkout);
  applyFraudRules(records);
  const potentialFraud = applyPotentialFraudFilter(records);

  emit('[6/7] Scanning checkout emails for Dashboard 2...');
  let allCheckoutEmails = [];
  try {
    allCheckoutEmails = await readCheckoutEmailStats(baseDir);
  } catch (err) {
    fileErrors.push({ step: 'Checkout (email stats)', folder: path.join(baseDir, 'Checkout'), error: err.message, fix: 'Ensure .csv files in Checkout folder have headers: Customer Email, Action Date UTC, Action Type, Amount.' });
    emit(`  WARNING: Checkout email stats failed (${err.message}), continuing without it.`);
  }
  const checkoutEmailTop20 = allCheckoutEmails.slice(0, 20);

  emit('[7/7] Generating Excel...');
  try {
    await generateExcel(records, potentialFraud, checkoutEmailTop20, totalDetailedRows, outputPath);
  } catch (err) {
    fileErrors.push({ step: 'Excel Generation', folder: path.dirname(outputPath), error: err.message, fix: 'Ensure the Output directory is writable and not locked by another process (close Excel if the file is open).' });
    throw new Error(`Excel generation failed: ${err.message}\n\nFile Errors:\n${JSON.stringify(fileErrors, null, 2)}`);
  }

  if (fileErrors.length > 0) {
    emit('\n--- WARNINGS: Some data sources had errors ---');
    fileErrors.forEach(e => emit(`  [${e.step}] ${e.error} | Fix: ${e.fix}`));
    emit('--- Analysis continued with available data ---\n');
  }

  const potentialFraudTotal = potentialFraud.reduce((s, r) => s + (r['Total Sale'] || 0), 0);
  const fraudPct = totalDetailedRows > 0 ? ((potentialFraud.length / totalDetailedRows) * 100).toFixed(2) : '0.00';

  const groupByField2 = (arr, field) => {
    const map = {};
    for (const r of arr) { const k = r[field] || '(blank)'; if (!map[k]) map[k] = { count: 0, total: 0 }; map[k].count++; map[k].total += r['Total Sale'] || 0; }
    return Object.entries(map).sort((a, b) => b[1].total - a[1].total).map(([k, v]) => ({ name: k, count: v.count, total: Math.round(v.total * 100) / 100 }));
  };

  const groupByComposite = (arr, fields) => {
    const map = {};
    for (const r of arr) {
      const key = fields.map(f => r[f] || '').join('||');
      if (!map[key]) { map[key] = { count: 0, total: 0, fields: {} }; fields.forEach(f => map[key].fields[f] = r[f] || ''); }
      map[key].count++; map[key].total += r['Total Sale'] || 0;
    }
    return Object.values(map).sort((a, b) => b.total - a.total).map(v => ({ ...v.fields, count: v.count, total: Math.round(v.total * 100) / 100 }));
  };

  const summary = {
    totalDetailedRows,
    potentialFraudCount: potentialFraud.length,
    potentialFraudValue: potentialFraudTotal,
    potentialFraudPct: fraudPct,
  };

  const dashboardData = {
    summary,
    potentialFraudRecords: potentialFraud.map((r, i) => ({
      serial: i + 1, restaurantId: r['Restaurant Id'], branchId: r['Branch Id'],
      restaurant: r['Restaurant'], branch: r['Branch'], orderId: r['Order Id'],
      dateTime: r['Date / Time'],
      paymentMethod: r['Payment Method'] || '', status: r['Status'] || '', serviceType: r['Service Type'] || '',
      subTotal: r['SubTotal'], totalSale: r['Total Sale'],
      netPayment: r['Net Payment Per Order'],
      customerGroup: r['Customer group'], customerAccount: r['Customer account'],
      customerEmail: r['Customer Email'],
      reasons: r['Potential Fraud Reasons'],
    })),
    byCustomerGroup: groupByField2(potentialFraud, 'Customer group'),
    byDate: (() => {
      const map = {};
      for (const r of potentialFraud) { const k = dateOnlyDDMMMYYYY(r['Date / Time']); if (!map[k]) map[k] = { count: 0, total: 0 }; map[k].count++; map[k].total += r['Total Sale'] || 0; }
      return Object.entries(map).sort((a, b) => a[0].localeCompare(b[0])).map(([k, v]) => ({ name: k, count: v.count, total: Math.round(v.total * 100) / 100 }));
    })(),
    byRestaurantAccount: groupByComposite(potentialFraud, ['Restaurant Id', 'Customer account', 'Restaurant']),
    byBranchId: groupByComposite(potentialFraud, ['Branch Id', 'Branch']),
    checkoutEmailTop20: checkoutEmailTop20,
    fileErrors: fileErrors.length > 0 ? fileErrors : null,
    generatedAt: new Date().toISOString(),
  };

  const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
  emit(`\nDone in ${elapsed}s | ${totalDetailedRows} total orders | ${potentialFraud.length} potential fraud | ${checkoutEmailTop20.length} suspicious emails`);
  return dashboardData;
}

module.exports = { runAnalysis };

if (require.main === module) {
  const gdriveDefault = 'G:/Shared drives/AR Team (Talabat UAE Finance) - Shared Drive/Claude/Fraud';
  const dataRoot = process.env.DATA_ROOT || (fs.existsSync(gdriveDefault) ? gdriveDefault : path.join(__dirname, 'data'));
  const baseDir = process.argv[2] || path.join(dataRoot, 'Input');
  const outputDir = path.join(dataRoot, 'Output');
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
  const ts = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
  const output = path.join(outputDir, `Fraud_Analysis_${ts}.xlsx`);
  runAnalysis(baseDir, output).then(data => {
    fs.writeFileSync(path.join(outputDir, 'dashboard_data.json'), JSON.stringify(data, null, 2));
    emit('Dashboard JSON saved to Output folder.');
  }).catch(err => { console.error('FATAL:', err.message, err.stack); process.exit(1); });
}
