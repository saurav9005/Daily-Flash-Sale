function sendTypeSummaryEmail() {
  // ===== CONFIG =====
  const SPREADSHEET_ID     = '1mo1BzavtXIxM8H_eBcR6oRKWOrFnoZHzcg6ejP95yYc';
  const SUMMARY_SHEET      = 'Product_Summary'; // [Type, Units, $ Sales, ASP]
  const TOP_PRODUCTS_SHEET = 'Raw_Data';        // A Title, B Type, C Option, D Image URL, E Product URL, F Units, G Sales

  // ✅ NEW (ONLY for Total Sales)
  const SALES_TOTAL_SHEET  = 'Raw_Data3';
  const SALES_TOTAL_COL    = 'Total Sales';
  const GROSS_TOTAL_COL    = 'Gross Sales';
  const DISCOUNT_TOTAL_COL = 'Discount';
  const RETURN_TOTAL_COL   = 'Returns';
  const NET_SALES_COL      = 'Net Sales';   


  const TO                 = 'sauravk@elietahari.com';
  const BASE_SUBJECT       = 'Elie Tahari | Daily Flash Report';
  const TIMEZONE           = Session.getScriptTimeZone();

  // === ATTACHMENT SETTINGS ===
  const SELECTED_SHEETS   = ['Top_Products','Product_Summary'];
  const REFERENCE_SHEETS  = ['Raw_Data','Raw_Data2'];
  const ATTACH_SELECTED_SHEETS = true;
  const ATTACH_MODE            = 'xlsx';

  // ===== HELPERS =====
  const money0 = n => '$' + Number(n || 0).toLocaleString(undefined, { minimumFractionDigits: 0 });
  const money2 = n => '$' + Number(n || 0).toLocaleString(undefined, { minimumFractionDigits: 2 });
  const num0   = n => Number(n || 0).toLocaleString();
  const toNumber = (v) => {
    if (v === null || v === '' || typeof v === 'boolean') return 0;
    if (v instanceof Date) return 0;
    if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
    const s = String(v).replace(/[^0-9.\-]/g, '');
    const n = Number(s);
    return Number.isFinite(n) ? n : 0;
  };

  // ✅ NEW: sum a column by header name
  const sumColumnByHeader_ = (sheet, headerName) => {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return 0;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
    const idx = headers.findIndex(h => h.toLowerCase() === String(headerName).trim().toLowerCase());
    if (idx === -1) throw new Error(`Column "${headerName}" not found in sheet "${sheet.getName()}".`);

    const values = sheet.getRange(2, idx + 1, lastRow - 1, 1).getValues();
    return values.reduce((sum, r) => sum + toNumber(r[0]), 0);
  };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ✅ NEW: Total Sales comes ONLY from Raw_Data3[Compare At price]
  const salesSh = ss.getSheetByName(SALES_TOTAL_SHEET);
  if (!salesSh) throw new Error('Sheet not found: ' + SALES_TOTAL_SHEET);
  const TOTAL_SALES = sumColumnByHeader_(salesSh, SALES_TOTAL_COL);
  const GROSS_SALES = sumColumnByHeader_(salesSh, GROSS_TOTAL_COL);
  const DISCOUNT_SALES = sumColumnByHeader_(salesSh, DISCOUNT_TOTAL_COL);
  const RETURNS_SALES = sumColumnByHeader_(salesSh, RETURN_TOTAL_COL);
  const NET_SALES = sumColumnByHeader_(salesSh, NET_SALES_COL);

  // ===== LOAD SUMMARY (by Product Type) =====
  const sh = ss.getSheetByName(SUMMARY_SHEET);
  if (!sh) throw new Error('Sheet not found: ' + SUMMARY_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) throw new Error('No data rows in ' + SUMMARY_SHEET);

  const raw = sh.getRange(2, 1, lastRow - 1, Math.min(4, lastCol)).getValues();
  let summary = raw.map(r => ({
    type:  String(r[0] || '').trim(),
    units: toNumber(r[1]),
    sales: toNumber(r[2]),
    asp:   toNumber(r[3])
  })).filter(r => r.type && Number.isFinite(r.units) && Number.isFinite(r.sales));

  if (!summary.length) throw new Error('No valid numeric rows in ' + SUMMARY_SHEET);

  // Totals (keep units from summary; total sales overridden by Raw_Data3)
  const totals = summary.reduce((a, r) => ({ units: a.units + r.units, sales: a.sales + r.sales }), { units: 0, sales: 0 });

  // Build dynamic subject (adds date, totals)
  const date = new Date();
  date.setDate(date.getDate() - 1);
  const dateStr = Utilities.formatDate(date, TIMEZONE, 'MMM d, yyyy');

  // ✅ CHANGED: totals.sales -> TOTAL_SALES
  const SUBJECT = `${BASE_SUBJECT} — ${dateStr} • ${num0(totals.units)} Units • ${money0(NET_SALES)}`;

  // ===== CHARTS (as images) =====
  const unitsDT = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Product Type')
    .addColumn(Charts.ColumnType.NUMBER, 'Units');
  summary.forEach(r => unitsDT.addRow([r.type, r.units]));
  const unitsChartBlob = Charts.newColumnChart()
    .setDataTable(unitsDT.build())
    .setLegendPosition(Charts.Position.NONE)
    .setTitle('Units by Product Type')
    .setDimensions(600, 400)
    .build().getAs('image/png');

  const salesDTbar = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Product Type')
    .addColumn(Charts.ColumnType.NUMBER, 'Sales');
  summary.forEach(r => salesDTbar.addRow([r.type, r.sales]));
  const salesBarBlob = Charts.newColumnChart()
    .setDataTable(salesDTbar.build())
    .setLegendPosition(Charts.Position.NONE)
    .setTitle('Revenue by Product Type')
    .setDimensions(600, 400)
    .build().getAs('image/png');

  const salesDTpie = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Product Type')
    .addColumn(Charts.ColumnType.NUMBER, 'Sales');
  summary.forEach(r => salesDTpie.addRow([r.type, r.sales]));
  const salesPieBlob = Charts.newPieChart()
    .setDataTable(salesDTpie.build())
    .setTitle('Revenue Share by Product Type')
    .setDimensions(700, 420)
    .build().getAs('image/png');

  // ===== TOP 5 PRODUCTS (from Raw_Data with direct Image URL) =====
  const tp = ss.getSheetByName(TOP_PRODUCTS_SHEET);
  if (!tp) throw new Error('Sheet not found: ' + TOP_PRODUCTS_SHEET);
  const tpLastRow = tp.getLastRow();
  if (tpLastRow < 2) throw new Error('No data rows in ' + TOP_PRODUCTS_SHEET);

  const tpVals  = tp.getRange(2, 1, tpLastRow - 1, 7).getValues();
  const tpLinks = tp.getRange(2, 5, tpLastRow - 1, 1).getRichTextValues();

  let products = tpVals.map((row, i) => {
    const title      = String(row[0] || '');
    const ptype      = String(row[1] || '');
    const option1    = String(row[2] || '');
    const imgUrl     = String(row[3] || '');
    const urlRich    = tpLinks[i] && tpLinks[i][0] ? tpLinks[i][0] : null;
    const productUrl = urlRich ? urlRich.getLinkUrl() : (row[4] || '');
    const units      = toNumber(row[5]);
    const sales      = toNumber(row[6]);
    return { title, ptype, option1, productUrl, units, sales, imgUrl };
  }).filter(p => p.title && (p.units > 0 || p.sales > 0));

  products.sort((a,b) => b.sales - a.sales);
  const top5 = products.slice(0, 5);

  const inlineTopImgs = {};
  top5.forEach((p, idx) => {
    const cid = `topImg${idx+1}`;
    let blob = p.imgUrl ? fetchImageWithAuth_(p.imgUrl) : null;
    if (blob) {
      inlineTopImgs[cid] = blob;
      p.cid = cid;
    } else {
      p.cid = null;
    }
  });

  const top5Html = top5.map(p => `
    <tr>
      <td style="padding:8px;border-bottom:1px solid #f0f2f5;width:84px">
        ${p.cid ? `<img src="cid:${p.cid}" alt="${p.title}" style="height:64px;width:auto;border-radius:6px;border:1px solid #eee"/>` : ''}
      </td>
      <td style="padding:8px;border-bottom:1px solid #f0f2f5">
        <div style="font-weight:600">${p.productUrl ? `<a href="${p.productUrl}" style="color:#0b66c3;text-decoration:none">${p.title}</a>` : p.title}</div>
        <div style="color:#666;font-size:12px">${p.ptype}${p.option1 ? ' • ' + p.option1 : ''}</div>
      </td>
      <td style="padding:8px;border-bottom:1px solid #f0f2f5;text-align:right">${num0(p.units)}</td>
      <td style="padding:8px;border-bottom:1px solid #f0f2f5;text-align:right">${money2(p.units ? p.sales / p.units : 0)}</td>
      <td style="padding:8px;border-bottom:1px solid #f0f2f5;text-align:right">${money0(p.sales)}</td>
      
    </tr>
  `).join('');

  const sortedBySales = [...summary].sort((a, b) => b.sales - a.sales);
  const typeRowsHtml = sortedBySales.map(r => `
    <tr>
      <td style="padding:8px;border-bottom:1px solid #f0f2f5">${r.type}</td>
      <td style="padding:8px;border-bottom:1px solid #f0f2f5;text-align:right">${num0(r.units)}</td>
      <td style="padding:8px;border-bottom:1px solid #f0f2f5;text-align:right">${money2(r.units ? r.sales / r.units : 0)}</td>
      <td style="padding:8px;border-bottom:1px solid #f0f2f5;text-align:right">${money0(r.sales)}</td>
      
    </tr>
  `).join('');

  const top1 = sortedBySales[0];
  const top2 = sortedBySales[1];
  const top3 = sortedBySales[2];

  // ✅ CHANGED: totals.sales -> TOTAL_SALES
  const top3Share = TOTAL_SALES > 0
    ? (((top1?.sales||0)+(top2?.sales||0)+(top3?.sales||0)) / TOTAL_SALES * 100)
    : 0;

  const tail = sortedBySales[sortedBySales.length - 1];

  const highlightsHtml = `
    <ul style="margin:0;padding-left:18px">
      ${top1 ? `<li><b>${top1.type}</b> leads with <b>${num0(top1.units)} units</b> and <b>${money0(top1.sales)}</b>.</li>` : ''}
      ${top2 ? `<li>Next: <b>${top2.type}</b> (${num0(top2.units)} units, ${money0(top2.sales)}).</li>` : ''}
      ${top3 ? `<li>Top 3 categories account for ~<b>${top3Share.toFixed(1)}%</b> of revenue.</li>` : ''}
      ${tail && tail.sales === 0 ? `<li>Low activity in <b>${tail.type}</b></li>` : ''}
    </ul>
  `;

  // ✅ CHANGED: “generating totals.sales” -> TOTAL_SALES
  const html = `
  <div style="font-family:Arial,Helvetica,sans-serif;line-height:1.5;color:#111;font-size:14px">
    <h2 style="margin:0 0 6px">Daily Flash Report</h2>
    <div style="color:#555;margin:0 0 16px">${dateStr}</div>

    <p style="margin:0 0 12px">
      We sold <b>${num0(totals.units)} units</b> across <b>${summary.length}</b> product types, generating <b>${money0(GROSS_SALES)}</b> in gross sales.
     
    </p>
    <ul>
    <li>Total Discount: ${money0(DISCOUNT_SALES)}</li>
    <li>Total Returns: ${money0(RETURNS_SALES)}</li>
    <li>Net Sale: ${money0(NET_SALES)}</li>
    <li>Total Sale: ${money0(TOTAL_SALES)}</li>
    </ul>
    <p> Below are the highlights, charts, today’s Top 5 products, and full breakdown by product type.</p>

    <h3 style="margin:16px 0 8px">Key Highlights</h3>
    ${highlightsHtml}

    <table role="presentation" cellpadding="0" cellspacing="0" style="margin:16px 0">
      <tr>
        <td style="padding-right:24px"><div style="font-size:12px;color:#666">Total Units</div><div style="font-size:20px;font-weight:700">${num0(totals.units)}</div></td>
        <td style="padding-right:24px"><div style="font-size:12px;color:#666">Total Revenue</div><div style="font-size:20px;font-weight:700">${money0(TOTAL_SALES)}</div></td>
        <td><div style="font-size:12px;color:#666"># Product Types</div><div style="font-size:20px;font-weight:700">${summary.length}</div></td>
      </tr>
    </table>

    <div style="display:flex;gap:12px;flex-wrap:wrap;margin:8px 0 16px">
      <img src="cid:unitsChart" style="max-width:100%;height:auto;border:1px solid #eee;border-radius:6px"/>
      <img src="cid:salesBar"  style="max-width:100%;height:auto;border:1px solid #eee;border-radius:6px"/>
      <img src="cid:salesPie"  style="max-width:100%;height:auto;border:1px solid #eee;border-radius:6px"/>
    </div>

    <h3 style="margin:16px 0 8px">Top 5 Products</h3>
    <table cellpadding="0" cellspacing="0" border="0" style="width:50%;border-collapse:collapse;margin-bottom:16px">
      <thead>
        <tr style="background:#f4f6f8">
          <th style="text-align:left;padding:10px;border-bottom:1px solid #e5e7eb">Image</th>
          <th style="text-align:left;padding:10px;border-bottom:1px solid #e5e7eb">Product</th>
          <th style="text-align:right;padding:10px;border-bottom:1px solid #e5e7eb">Units</th>
          <th style="text-align:right;padding:10px;border-bottom:1px solid #e5e7eb">ASP</th>
          <th style="text-align:right;padding:10px;border-bottom:1px solid #e5e7eb">$ Sales</th>          
        </tr>
      </thead>
      <tbody>
        ${top5Html}
      </tbody>
    </table>

    <h3 style="margin:16px 0 8px">By Product Type</h3>
    <table cellpadding="0" cellspacing="0" border="0" style="width:50%;border-collapse:collapse;margin-top:8px">
      <thead>
        <tr style="background:#f4f6f8">
          <th style="text-align:left;padding:10px;border-bottom:1px solid #e5e7eb">Product Type</th>
          <th style="text-align:right;padding:10px;border-bottom:1px solid #e5e7eb">Units</th>
          <th style="text-align:right;padding:10px;border-bottom:1px solid #e5e7eb">ASP</th>
          <th style="text-align:right;padding:10px;border-bottom:1px solid #e5e7eb">$ Sales</th>          
        </tr>
      </thead>
      <tbody>
        ${typeRowsHtml}
      </tbody>
    </table>
  </div>
  `;

  // ===== Build attachments =====
  let attachments = [];
  if (ATTACH_SELECTED_SHEETS) {
    if (ATTACH_MODE === 'xlsx') {
      const xlsxBlob = exportSheetsWithReferencesAsXlsx_(SPREADSHEET_ID, SELECTED_SHEETS, REFERENCE_SHEETS, true);
      if (xlsxBlob) attachments.push(xlsxBlob);
    } else if (ATTACH_MODE === 'csv') {
      SELECTED_SHEETS.forEach(name => {
        const csvBlob = exportSheetAsCsv_(SPREADSHEET_ID, name);
        if (csvBlob) attachments.push(csvBlob);
      });
    }
  }

  // ===== Send email =====
  GmailApp.sendEmail(
    TO,
    SUBJECT,
    '',
    {
      htmlBody: html,
      inlineImages: {
        unitsChart: unitsChartBlob,
        salesBar:   salesBarBlob,
        salesPie:   salesPieBlob,
        ...inlineTopImgs
      },
      attachments: attachments.length ? attachments : undefined
    }
  );
}

// (rest of your helper functions unchanged...)

/** Fetch image (public or Google-auth’d) and return a Blob; null on failure */
function fetchImageWithAuth_(url) {
  if (!url) return null;
  try {
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (r.getResponseCode() >= 200 && r.getResponseCode() < 300) return r.getBlob();
  } catch (e) {}
  try {
    const token = ScriptApp.getOAuthToken();
    const r = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (r.getResponseCode() >= 200 && r.getResponseCode() < 300) return r.getBlob();
  } catch (e) {}
  return null;
}

/** Export a single sheet as CSV (no cross-sheet refs supported) */
function exportSheetAsCsv_(spreadsheetId, sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found: ' + sheetName);
  const gid = sh.getSheetId();
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=${gid}`;
  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error(`CSV export failed for "${sheetName}": ` + resp.getResponseCode() + ' — ' + resp.getContentText());
  return resp.getBlob().setName(`${sheetName}.csv`);
}

/** Export selected sheets + reference sheets as a single XLSX.
 *  - Copies REFERENCE_SHEETS first (keeps formulas), then SELECTED_SHEETS (keeps formulas)
 *  - Hides all REFERENCE_SHEETS in the temp workbook
 *  - Exports to XLSX and deletes the temp file
 */
function exportSheetsWithReferencesAsXlsx_(spreadsheetId, selectedSheetNames, referenceSheetNames, hideReferenceSheets) {
  const src = SpreadsheetApp.openById(spreadsheetId);
  const temp = SpreadsheetApp.create('TMP_' + src.getName() + '_' + new Date().getTime());
  const defaultSheet = temp.getSheets()[0];

  const copiedNames = [];         // track names used in temp
  const safeCopy = (name) => {
    const sh = src.getSheetByName(name);
    if (!sh) throw new Error('Sheet not found: ' + name);
    const copied = sh.copyTo(temp);
    try { copied.setName(name); }
    catch (e) { copied.setName(name + '_' + Utilities.getUuid().slice(0,4)); }
    copiedNames.push(copied.getName());
    return copied;
  };

  try {
    // 1) Copy reference sheets first
    const copiedRefs = (referenceSheetNames || []).map(n => safeCopy(n));

    // 2) Copy selected/visible sheets
    const copiedSelected = (selectedSheetNames || []).map(n => safeCopy(n));

    if (copiedRefs.length + copiedSelected.length === 0) {
      throw new Error('No sheets were copied; nothing to export.');
    }

    // 3) Hide reference sheets (so recipients don't see them by default)
    if (hideReferenceSheets) {
      copiedRefs.forEach(s => {
        try { s.hideSheet(); } catch (e) {}
      });
    }

    // 4) Delete the default blank sheet (after we have at least one copied sheet)
    try {
      temp.deleteSheet(defaultSheet);
    } catch (e) {
      defaultSheet.clear();
      defaultSheet.setName('tmp');
      defaultSheet.hideSheet();
    }

    // 5) Export temp as XLSX
    const url = `https://docs.google.com/spreadsheets/d/${temp.getId()}/export?format=xlsx`;
    const token = ScriptApp.getOAuthToken();
    const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      throw new Error('Selected-sheets XLSX export failed: ' + resp.getResponseCode() + ' — ' + resp.getContentText());
    }

    const shortList = selectedSheetNames.join('+').slice(0, 80);
    return resp.getBlob().setName(src.getName() + ' — ' + shortList + '.xlsx');

  } finally {
    // 6) Clean up temp file in Drive
    try { DriveApp.getFileById(temp.getId()).setTrashed(true); } catch (e) {}
  }
}
