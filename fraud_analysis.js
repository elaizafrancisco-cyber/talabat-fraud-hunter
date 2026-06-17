const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

const BASE = 'G:/Shared drives/AR Team (Talabat UAE Finance) - Shared Drive/Claude/Fraud';
const OUTPUT = path.join(__dirname, 'Fraud_Analysis_Output.xlsx');

const COLORS = {
  primary: 'FF5900',
  accent: '411517',
  secondary: 'F4EDE3',
  highlight: 'CFFF00',
  white: 'FFFFFF',
  black: '000000',
};

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
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const yyyy = d.getUTCFullYear();
  return `${yyyy}-${mm}-${dd} ${String(hrs).padStart(2,'0')}:${String(mins).padStart(2,'0')}`;
}

async function readDetailedReports() {
  const files = [
    path.join(BASE, 'Detailed Reports/Detailed__4347c522-7f73-4506-ba8b-8ddf92b749d5.xlsx'),
    path.join(BASE, 'Detailed Reports/Detailed__54525ed9-c794-4e19-94a7-83f0ec683aa2.xlsx'),
  ];

  const colMap = {
    12: 'Restaurant Id', 13: 'Restaurant', 14: 'Branch Id', 15: 'Branch',
    16: 'Order Id', 18: 'Date / Time', 19: 'Payment Method', 20: 'Status',
    21: 'Service Type', 25: 'SubTotal', 32: 'Total Sale', 64: 'Net Payment Per Order',
  };
  const colNums = Object.keys(colMap).map(Number);
  const records = [];
  const seenOrderIds = new Set();

  for (const file of files) {
    console.log(`Reading: ${path.basename(file)}`);
    const wb = new ExcelJS.stream.xlsx.WorkbookReader(file);
    let rowNum = 0;
    for await (const ws of wb) {
      for await (const row of ws) {
        rowNum++;
        if (rowNum <= 3) continue;
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
    console.log(`  Accumulated ${records.length} records (Total Sale >= 1000)`);
  }
  return records;
}

function readBranchInquiry() {
  console.log('Reading Branch Inquiry...');
  const wb = XLSX.readFile(path.join(BASE, 'Branch Inquiry/Branch Codes & Customer List.xlsx'));
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
  console.log(`  ${Object.keys(lookup).length} branch entries loaded`);
  return lookup;
}

function readAX365() {
  console.log('Reading AX365...');
  const wb = XLSX.readFile(path.join(BASE, 'AX365/AXDynamicsReport.xlsx'));
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const lookup = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;
    const branchId = String(row[4] || '');
    const branchAXCode = String(row[6] || '');
    if (branchId) {
      lookup[branchId] = { 'Branch AX Code': branchAXCode };
    }
  }
  console.log(`  ${Object.keys(lookup).length} AX entries loaded`);
  return lookup;
}

function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (inQuotes) {
      if (c === '"' && line[i + 1] === '"') { current += '"'; i++; }
      else if (c === '"') { inQuotes = false; }
      else { current += c; }
    } else {
      if (c === '"') { inQuotes = true; }
      else if (c === ',') { result.push(current); current = ''; }
      else { current += c; }
    }
  }
  result.push(current);
  return result;
}

async function readCheckout(orderIds) {
  console.log('Reading Checkout files (streaming)...');
  const checkoutDir = path.join(BASE, 'Checkout');
  const files = fs.readdirSync(checkoutDir).filter(f => f.endsWith('.csv'));
  const lookup = {};
  const orderIdSet = new Set(orderIds);

  for (const file of files) {
    const filePath = path.join(checkoutDir, file);
    console.log(`  Streaming: ${file}`);
    let headers = null;
    let refIdx, actionIdx, respIdx, emailIdx, amountIdx;
    let lineCount = 0;
    let matchCount = 0;

    const rl = readline.createInterface({
      input: fs.createReadStream(filePath, { encoding: 'utf-8' }),
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

      if (!lookup[ref]) {
        lookup[ref] = { actionType, respDesc, email, amount };
        matchCount++;
      } else if (actionType === 'Authorisation') {
        lookup[ref] = { actionType, respDesc, email, amount };
      }
    }
    console.log(`    ${lineCount} lines scanned, ${matchCount} new matches`);
  }
  console.log(`  ${Object.keys(lookup).length} total order matches from checkout`);
  return lookup;
}

function mergeData(records, branchInq, ax365, checkout) {
  console.log('Merging data...');
  for (const rec of records) {
    const bid = rec['Branch Id'];
    const oid = rec['Order Id'];

    const bi = branchInq[bid] || {};
    rec['Customer group'] = bi['Customer group'] || '';
    rec['Customer account'] = bi['Customer account'] || '';

    const ax = ax365[bid] || {};
    rec['Branch AX Code'] = ax['Branch AX Code'] || '';

    rec['MSD Checker'] = rec['Customer account'] && rec['Branch AX Code']
      ? (rec['Customer account'] === rec['Branch AX Code'] ? 'True' : 'False')
      : '';

    const co = checkout[oid] || {};
    rec['Action Type'] = co.actionType || '';
    rec['Response Description'] = co.respDesc || '';
    rec['Customer Email'] = co.email || '';
    rec['Checkout Amount'] = co.amount || null;
  }
  console.log(`  Merge complete: ${records.length} records`);
  return records;
}

function applyFraudRules(records) {
  console.log('Applying fraud detection rules...');

  const dupByRestaurantId = {};
  const dupByBranchId = {};
  const dupByCustGroup = {};
  const dupByEmail = {};
  const emailCounts = {};

  for (const rec of records) {
    const ts = rec['Total Sale'];
    const rid = rec['Restaurant Id'];
    const bid = rec['Branch Id'];
    const cg = rec['Customer group'];
    const em = rec['Customer Email'];

    const key1 = `${rid}_${ts}`;
    dupByRestaurantId[key1] = (dupByRestaurantId[key1] || 0) + 1;

    const key2 = `${bid}_${ts}`;
    dupByBranchId[key2] = (dupByBranchId[key2] || 0) + 1;

    if (cg) {
      const key3 = `${cg}_${ts}`;
      dupByCustGroup[key3] = (dupByCustGroup[key3] || 0) + 1;
    }

    if (em) {
      const key4 = `${em}_${ts}`;
      dupByEmail[key4] = (dupByEmail[key4] || 0) + 1;
      emailCounts[em] = (emailCounts[em] || 0) + 1;
    }
  }

  let flaggedCount = 0;
  for (const rec of records) {
    const reasons = [];
    const ts = rec['Total Sale'];
    const rid = rec['Restaurant Id'];
    const bid = rec['Branch Id'];
    const cg = rec['Customer group'];
    const em = rec['Customer Email'];
    const pm = rec['Payment Method'] || '';
    const coAmt = rec['Checkout Amount'];

    if (ts >= 10000) {
      reasons.push(`Total Sale unusually high: AED ${ts.toLocaleString()}`);
    }

    if (dupByRestaurantId[`${rid}_${ts}`] > 1) {
      reasons.push(`Duplicate Total Sale (AED ${ts}) under same Restaurant Id ${rid}`);
    }
    if (dupByBranchId[`${bid}_${ts}`] > 1) {
      reasons.push(`Duplicate Total Sale (AED ${ts}) under same Branch Id ${bid}`);
    }
    if (cg && dupByCustGroup[`${cg}_${ts}`] > 1) {
      reasons.push(`Duplicate Total Sale (AED ${ts}) under same Customer Group ${cg}`);
    }
    if (em && dupByEmail[`${em}_${ts}`] > 1) {
      reasons.push(`Duplicate Total Sale (AED ${ts}) under same Customer Email`);
    }

    if (em && emailCounts[em] > 2) {
      reasons.push(`Customer Email used ${emailCounts[em]} times (suspicious frequency)`);
    }

    const pmLower = pm.toLowerCase();
    if (pmLower === 'credit card' || pmLower === 'talabat credit') {
      if (coAmt !== null && coAmt !== undefined) {
        const diff = coAmt - ts;
        if (diff < -200) {
          reasons.push(`Checkout Amount (${coAmt}) minus Total Sale (${ts}) = ${diff.toFixed(2)}, below -200 threshold`);
        } else if (coAmt > ts) {
          // coAmt > ts => mark No (override to not fraud for this rule)
        }
      }
    }

    if (pmLower === 'cash') {
      reasons.push('Payment Method is Cash (high-value cash order)');
    }

    rec['Fraud Reasons'] = reasons.join('; ');
    rec['Potential Fraud'] = reasons.length > 0 ? 'Yes' : 'No';
    if (reasons.length > 0) flaggedCount++;
  }

  console.log(`  Flagged ${flaggedCount} of ${records.length} as potential fraud`);
  return records;
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
  const bgColor = isFraud
    ? 'FFFDE8E8'
    : (rowIdx % 2 === 0 ? 'FF' + COLORS.secondary : 'FF' + COLORS.white);
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    cell.font = { size: 10, color: { argb: 'FF' + COLORS.accent } };
    cell.border = {
      bottom: { style: 'hair', color: { argb: 'FFCCCCCC' } },
      left: { style: 'hair', color: { argb: 'FFCCCCCC' } },
      right: { style: 'hair', color: { argb: 'FFCCCCCC' } },
    };
    cell.alignment = { vertical: 'middle' };
  }
}

async function generateOutput(records) {
  console.log('Generating output Excel...');
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Talabat Fraud Hunter';
  wb.created = new Date();

  // --- Sheet 1: Main Processed Data ---
  const ws1 = wb.addWorksheet('Main Data', {
    properties: { tabColor: { argb: 'FF' + COLORS.primary } },
  });

  const mainCols = [
    'Restaurant Id', 'Branch Id', 'Restaurant', 'Branch', 'Order Id',
    'Date / Time', 'Payment Method', 'Status', 'Service Type',
    'SubTotal', 'Total Sale', 'Net Payment Per Order',
    'Customer group', 'Customer account', 'Branch AX Code', 'MSD Checker',
    'Action Type', 'Response Description', 'Customer Email', 'Checkout Amount',
    'Potential Fraud', 'Fraud Reasons',
  ];

  ws1.columns = mainCols.map(h => ({
    header: h, key: h,
    width: h === 'Fraud Reasons' ? 60 : h === 'Customer Email' ? 30 : 18,
  }));

  const headerRow1 = ws1.getRow(1);
  applyHeaderStyle(ws1, headerRow1, mainCols.length);
  headerRow1.height = 30;
  ws1.autoFilter = { from: 'A1', to: `${String.fromCharCode(64 + mainCols.length)}1` };

  let rowIdx1 = 0;
  for (const rec of records) {
    const row = ws1.addRow(mainCols.map(c => rec[c] ?? ''));
    rowIdx1++;
    applyDataRowStyle(ws1, row, mainCols.length, rowIdx1, rec['Potential Fraud'] === 'Yes');
  }
  ws1.views = [{ state: 'frozen', ySplit: 1 }];

  // --- Sheet 2: Potential Fraud Orders ---
  const ws2 = wb.addWorksheet('Potential Fraud', {
    properties: { tabColor: { argb: 'FFFF0000' } },
  });

  const fraudCols = [
    'Serial', 'Restaurant Id', 'Branch Id', 'Restaurant', 'Branch', 'Order Id',
    'SubTotal', 'Total Sale', 'Net Payment Per Order',
    'Customer group', 'Customer Email', 'Fraud Reasons',
  ];

  ws2.columns = fraudCols.map(h => ({
    header: h, key: h,
    width: h === 'Fraud Reasons' ? 60 : h === 'Customer Email' ? 30 : 18,
  }));

  const headerRow2 = ws2.getRow(1);
  applyHeaderStyle(ws2, headerRow2, fraudCols.length);
  headerRow2.height = 30;

  const autoFilterCol2 = String.fromCharCode(64 + fraudCols.length);
  ws2.autoFilter = { from: 'A1', to: `${autoFilterCol2}1` };

  const fraudRecords = records.filter(r => r['Potential Fraud'] === 'Yes');
  let serial = 0;
  for (const rec of fraudRecords) {
    serial++;
    const rowData = [
      serial, rec['Restaurant Id'], rec['Branch Id'], rec['Restaurant'], rec['Branch'],
      rec['Order Id'], rec['SubTotal'], rec['Total Sale'], rec['Net Payment Per Order'],
      rec['Customer group'], rec['Customer Email'], rec['Fraud Reasons'],
    ];
    const row = ws2.addRow(rowData);
    applyDataRowStyle(ws2, row, fraudCols.length, serial, true);
  }
  ws2.views = [{ state: 'frozen', ySplit: 1 }];

  // --- Sheet 3: Dashboard 1 - Fraud Summary ---
  const ws3 = wb.addWorksheet('Dashboard 1', {
    properties: { tabColor: { argb: 'FF' + COLORS.highlight } },
  });

  const yesRecords = records.filter(r => r['Potential Fraud'] === 'Yes');
  const noRecords = records.filter(r => r['Potential Fraud'] === 'No');

  // Title
  ws3.mergeCells('A1:H1');
  const titleCell = ws3.getCell('A1');
  titleCell.value = 'FRAUD ANALYSIS DASHBOARD';
  titleCell.font = { bold: true, size: 18, color: { argb: 'FF' + COLORS.white } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + COLORS.primary } };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws3.getRow(1).height = 40;

  // Summary
  ws3.getCell('A3').value = 'SUMMARY';
  ws3.getCell('A3').font = { bold: true, size: 14, color: { argb: 'FF' + COLORS.primary } };
  ws3.mergeCells('A3:H3');

  const summaryHeaders = ['Category', 'Count', 'Total Value (AED)'];
  const summaryRow = ws3.addRow(summaryHeaders);
  summaryRow.number = 4;
  ws3.getRow(4).values = summaryHeaders;
  applyHeaderStyle(ws3, ws3.getRow(4), 3);

  const yesTotal = yesRecords.reduce((s, r) => s + (r['Total Sale'] || 0), 0);
  const noTotal = noRecords.reduce((s, r) => s + (r['Total Sale'] || 0), 0);

  const r5 = ws3.getRow(5);
  r5.values = ['Potential Fraud: Yes', yesRecords.length, yesTotal.toFixed(2)];
  r5.getCell(1).font = { bold: true, color: { argb: 'FFFF0000' } };
  r5.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE8E8' } };
  r5.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE8E8' } };
  r5.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE8E8' } };
  r5.getCell(3).numFmt = '#,##0.00';

  const r6 = ws3.getRow(6);
  r6.values = ['Potential Fraud: No', noRecords.length, noTotal.toFixed(2)];
  r6.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
  r6.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
  r6.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
  r6.getCell(3).numFmt = '#,##0.00';

  ws3.getColumn(1).width = 25;
  ws3.getColumn(2).width = 15;
  ws3.getColumn(3).width = 25;

  // Breakdown by Customer Group
  let currentRow = 9;
  ws3.getCell(`A${currentRow}`).value = 'BREAKDOWN BY CUSTOMER GROUP';
  ws3.getCell(`A${currentRow}`).font = { bold: true, size: 13, color: { argb: 'FF' + COLORS.primary } };
  ws3.mergeCells(`A${currentRow}:H${currentRow}`);
  currentRow++;

  const groupByField = (arr, field) => {
    const map = {};
    for (const r of arr) {
      const key = r[field] || '(blank)';
      if (!map[key]) map[key] = { count: 0, total: 0 };
      map[key].count++;
      map[key].total += r['Total Sale'] || 0;
    }
    return Object.entries(map).sort((a, b) => b[1].total - a[1].total);
  };

  const custGroupBreakdown = groupByField(yesRecords, 'Customer group');
  ws3.getRow(currentRow).values = ['Customer Group', 'Fraud Count', 'Total Amount (AED)'];
  applyHeaderStyle(ws3, ws3.getRow(currentRow), 3);
  currentRow++;

  for (const [group, data] of custGroupBreakdown) {
    const row = ws3.getRow(currentRow);
    row.values = [group, data.count, parseFloat(data.total.toFixed(2))];
    row.getCell(3).numFmt = '#,##0.00';
    applyDataRowStyle(ws3, row, 3, currentRow, false);
    currentRow++;
  }

  // Breakdown by Restaurant Id
  currentRow += 2;
  ws3.getCell(`A${currentRow}`).value = 'BREAKDOWN BY RESTAURANT ID';
  ws3.getCell(`A${currentRow}`).font = { bold: true, size: 13, color: { argb: 'FF' + COLORS.primary } };
  ws3.mergeCells(`A${currentRow}:H${currentRow}`);
  currentRow++;

  const restBreakdown = groupByField(yesRecords, 'Restaurant Id');
  ws3.getRow(currentRow).values = ['Restaurant Id', 'Restaurant', 'Fraud Count', 'Total Amount (AED)'];
  applyHeaderStyle(ws3, ws3.getRow(currentRow), 4);
  currentRow++;

  const restNameMap = {};
  for (const r of records) restNameMap[r['Restaurant Id']] = r['Restaurant'];

  for (const [rid, data] of restBreakdown.slice(0, 50)) {
    const row = ws3.getRow(currentRow);
    row.values = [rid, restNameMap[rid] || '', data.count, parseFloat(data.total.toFixed(2))];
    row.getCell(4).numFmt = '#,##0.00';
    applyDataRowStyle(ws3, row, 4, currentRow, false);
    currentRow++;
  }

  // Breakdown by Branch Id
  currentRow += 2;
  ws3.getCell(`A${currentRow}`).value = 'BREAKDOWN BY BRANCH ID';
  ws3.getCell(`A${currentRow}`).font = { bold: true, size: 13, color: { argb: 'FF' + COLORS.primary } };
  ws3.mergeCells(`A${currentRow}:H${currentRow}`);
  currentRow++;

  const branchBreakdown = groupByField(yesRecords, 'Branch Id');
  ws3.getRow(currentRow).values = ['Branch Id', 'Branch', 'Restaurant', 'Fraud Count', 'Total Amount (AED)'];
  applyHeaderStyle(ws3, ws3.getRow(currentRow), 5);
  currentRow++;

  const branchNameMap = {};
  const branchRestMap = {};
  for (const r of records) {
    branchNameMap[r['Branch Id']] = r['Branch'];
    branchRestMap[r['Branch Id']] = r['Restaurant'];
  }

  for (const [bid, data] of branchBreakdown.slice(0, 50)) {
    const row = ws3.getRow(currentRow);
    row.values = [bid, branchNameMap[bid] || '', branchRestMap[bid] || '', data.count, parseFloat(data.total.toFixed(2))];
    row.getCell(5).numFmt = '#,##0.00';
    applyDataRowStyle(ws3, row, 5, currentRow, false);
    currentRow++;
  }

  // Breakdown by Restaurant name
  currentRow += 2;
  ws3.getCell(`A${currentRow}`).value = 'BREAKDOWN BY RESTAURANT';
  ws3.getCell(`A${currentRow}`).font = { bold: true, size: 13, color: { argb: 'FF' + COLORS.primary } };
  ws3.mergeCells(`A${currentRow}:H${currentRow}`);
  currentRow++;

  const restNameBreakdown = groupByField(yesRecords, 'Restaurant');
  ws3.getRow(currentRow).values = ['Restaurant', 'Fraud Count', 'Total Amount (AED)'];
  applyHeaderStyle(ws3, ws3.getRow(currentRow), 3);
  currentRow++;

  for (const [name, data] of restNameBreakdown.slice(0, 50)) {
    const row = ws3.getRow(currentRow);
    row.values = [name, data.count, parseFloat(data.total.toFixed(2))];
    row.getCell(3).numFmt = '#,##0.00';
    applyDataRowStyle(ws3, row, 3, currentRow, false);
    currentRow++;
  }

  // Breakdown by Branch name
  currentRow += 2;
  ws3.getCell(`A${currentRow}`).value = 'BREAKDOWN BY BRANCH';
  ws3.getCell(`A${currentRow}`).font = { bold: true, size: 13, color: { argb: 'FF' + COLORS.primary } };
  ws3.mergeCells(`A${currentRow}:H${currentRow}`);
  currentRow++;

  const branchNameBreakdown = groupByField(yesRecords, 'Branch');
  ws3.getRow(currentRow).values = ['Branch', 'Fraud Count', 'Total Amount (AED)'];
  applyHeaderStyle(ws3, ws3.getRow(currentRow), 3);
  currentRow++;

  for (const [name, data] of branchNameBreakdown.slice(0, 50)) {
    const row = ws3.getRow(currentRow);
    row.values = [name, data.count, parseFloat(data.total.toFixed(2))];
    row.getCell(3).numFmt = '#,##0.00';
    applyDataRowStyle(ws3, row, 3, currentRow, false);
    currentRow++;
  }

  ws3.getColumn(4).width = 25;
  ws3.getColumn(5).width = 22;

  // --- Sheet 4: Dashboard 2 - Customer Email Analysis ---
  const ws4 = wb.addWorksheet('Dashboard 2', {
    properties: { tabColor: { argb: 'FF' + COLORS.accent } },
  });

  ws4.mergeCells('A1:D1');
  const title4 = ws4.getCell('A1');
  title4.value = 'CUSTOMER EMAIL ORDER FREQUENCY';
  title4.font = { bold: true, size: 18, color: { argb: 'FF' + COLORS.white } };
  title4.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + COLORS.primary } };
  title4.alignment = { horizontal: 'center', vertical: 'middle' };
  ws4.getRow(1).height = 40;

  const allEmailCounts = {};
  for (const rec of records) {
    const em = rec['Customer Email'];
    if (em) allEmailCounts[em] = (allEmailCounts[em] || 0) + 1;
  }

  const emailRanking = Object.entries(allEmailCounts)
    .sort((a, b) => b[1] - a[1]);

  ws4.getRow(3).values = ['Rank', 'Customer Email', 'Order Count', 'Flag'];
  applyHeaderStyle(ws4, ws4.getRow(3), 4);

  ws4.getColumn(1).width = 8;
  ws4.getColumn(2).width = 40;
  ws4.getColumn(3).width = 15;
  ws4.getColumn(4).width = 20;

  let rank = 0;
  for (const [email, count] of emailRanking) {
    rank++;
    const row = ws4.getRow(rank + 3);
    const flag = count > 2 ? 'Suspicious' : 'Normal';
    row.values = [rank, email, count, flag];
    applyDataRowStyle(ws4, row, 4, rank, count > 2);
    if (count > 2) {
      row.getCell(4).font = { bold: true, color: { argb: 'FFFF0000' }, size: 10 };
    }
    if (rank >= 500) break;
  }
  ws4.views = [{ state: 'frozen', ySplit: 3 }];

  // --- Sheet 5: Automailer ---
  const ws5 = wb.addWorksheet('Automailer', {
    properties: { tabColor: { argb: 'FF' + COLORS.highlight } },
  });

  ws5.mergeCells('A1:F1');
  const title5 = ws5.getCell('A1');
  title5.value = 'AUTOMAILER - FRAUD ALERT NOTIFICATION';
  title5.font = { bold: true, size: 16, color: { argb: 'FF' + COLORS.white } };
  title5.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + COLORS.primary } };
  title5.alignment = { horizontal: 'center', vertical: 'middle' };
  ws5.getRow(1).height = 40;

  ws5.getCell('A3').value = 'FROM:';
  ws5.getCell('A3').font = { bold: true, size: 12, color: { argb: 'FF' + COLORS.accent } };
  ws5.getCell('B3').value = '';
  ws5.getCell('B3').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9C4' } };
  ws5.mergeCells('B3:F3');

  ws5.getCell('A4').value = 'TO:';
  ws5.getCell('A4').font = { bold: true, size: 12, color: { argb: 'FF' + COLORS.accent } };
  ws5.getCell('B4').value = '';
  ws5.getCell('B4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9C4' } };
  ws5.mergeCells('B4:F4');

  ws5.getCell('A5').value = 'CC:';
  ws5.getCell('A5').font = { bold: true, size: 12, color: { argb: 'FF' + COLORS.accent } };
  ws5.getCell('B5').value = '';
  ws5.getCell('B5').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9C4' } };
  ws5.mergeCells('B5:F5');

  ws5.getCell('A6').value = 'Subject:';
  ws5.getCell('A6').font = { bold: true, size: 12, color: { argb: 'FF' + COLORS.accent } };
  ws5.getCell('B6').value = '';
  ws5.getCell('B6').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9C4' } };
  ws5.mergeCells('B6:F6');

  ws5.getCell('A8').value = '[SEND EMAIL]';
  ws5.getCell('A8').font = { bold: true, size: 14, color: { argb: 'FF' + COLORS.white } };
  ws5.getCell('A8').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + COLORS.primary } };
  ws5.getCell('A8').alignment = { horizontal: 'center', vertical: 'middle' };
  ws5.mergeCells('A8:C8');
  ws5.getRow(8).height = 35;
  ws5.getCell('A8').border = {
    top: { style: 'medium', color: { argb: 'FF' + COLORS.accent } },
    bottom: { style: 'medium', color: { argb: 'FF' + COLORS.accent } },
    left: { style: 'medium', color: { argb: 'FF' + COLORS.accent } },
    right: { style: 'medium', color: { argb: 'FF' + COLORS.accent } },
  };

  ws5.getCell('A10').value = 'Note: This is a placeholder button. Actual email sending is manual.';
  ws5.getCell('A10').font = { italic: true, size: 10, color: { argb: 'FF999999' } };
  ws5.mergeCells('A10:F10');

  ws5.getCell('A12').value = 'FRAUD ALERT SUMMARY';
  ws5.getCell('A12').font = { bold: true, size: 13, color: { argb: 'FF' + COLORS.primary } };
  ws5.mergeCells('A12:F12');

  ws5.getRow(13).values = ['', 'Metric', 'Value'];
  applyHeaderStyle(ws5, ws5.getRow(13), 3);

  const summaryData = [
    ['Total Orders Analyzed', records.length],
    ['Potential Fraud Orders', yesRecords.length],
    ['Total Fraud Value (AED)', yesTotal.toFixed(2)],
    ['Clean Orders', noRecords.length],
    ['Fraud Rate (%)', ((yesRecords.length / records.length) * 100).toFixed(2) + '%'],
  ];

  summaryData.forEach((item, i) => {
    const row = ws5.getRow(14 + i);
    row.values = ['', item[0], item[1]];
    applyDataRowStyle(ws5, row, 3, i, false);
  });

  ws5.getColumn(1).width = 12;
  ws5.getColumn(2).width = 30;
  ws5.getColumn(3).width = 25;
  ws5.getColumn(4).width = 20;
  ws5.getColumn(5).width = 20;
  ws5.getColumn(6).width = 20;

  await wb.xlsx.writeFile(OUTPUT);
  console.log(`\nOutput saved to: ${OUTPUT}`);
  console.log(`Total records: ${records.length}`);
  console.log(`Fraud flagged: ${yesRecords.length}`);
  console.log(`Clean: ${noRecords.length}`);
}

async function main() {
  console.log('=== TALABAT FRAUD HUNTER ===\n');
  const t0 = Date.now();

  const records = await readDetailedReports();
  const branchInq = readBranchInquiry();
  const ax365 = readAX365();
  const orderIds = records.map(r => r['Order Id']);
  const checkout = await readCheckout(orderIds);

  mergeData(records, branchInq, ax365, checkout);
  applyFraudRules(records);
  await generateOutput(records);

  console.log(`\nCompleted in ${((Date.now() - t0) / 1000).toFixed(1)}s`);
}

main().catch(err => {
  console.error('FATAL:', err.message);
  console.error(err.stack);
  process.exit(1);
});
