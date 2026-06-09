/*************** CONFIGURATION ***************/
const SHOPIFY_CONFIG = {
  SHOP_NAME: 'abc-def',                // e.g. 'elietahari'
  API_VERSION: '2024-01',                  // use the API version you’re on
  ACCESS_TOKEN: '',                    // TODO: replace with real Admin API token
  SHEET_NAME: 'ET_inventory'               // current snapshot sheet
};

const SPREADSHEET_ID = '1P9BMrK2KwxgKNNAVlAXfMNq2xXHcsM1oJDTkk5bTEKg';
const INVENTORY_LOG_SHEET = 'Inventory_Log'; // change log sheet name

// people who should receive the daily email report
const EMAIL_RECIPIENTS = [
  'sauravk@elietahari.com'
];
/*********************************************/
function getDynamicEmailSubject() {
  const now = new Date();
  const formatted = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
  return `Elie Tahari Inventory Report – ${formatted}`;
}
const EMAIL_SUBJECT = getDynamicEmailSubject();
const EMAIL_BODY =
  'Hi Team,\n\n' +
  'Please find attached the latest Elie Tahari inventory snapshot along with the change log.\n\n' +
  '• ET_inventory.csv → Current stock by product/variant\n' +
  '• Inventory_Log.csv → All inventory increases/decreases logged every 2 hours\n\n' +  
  'Thanks,\nSaurav Bot';

/**
 * Main function: fetch all Shopify products and write
 * product title, variant title, sku, available qty to the sheet.
 * Also log any quantity changes into Inventory_Log.
 */
function downloadShopifyInventory() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 1) Read previous snapshot from ET_inventory into a map keyed by SKU
  const oldSnapshotMap = getCurrentInventoryMap(ss);

  // 2) Fetch latest data from Shopify
  const products = fetchAllShopifyProducts();

  const rows = [];
  const header = ['product_title', 'variant_title', 'sku', 'available_qty'];
  rows.push(header);

  // We'll also prepare data for change-log comparison
  const newSnapshotMap = {}; // { sku: { product_title, variant_title, qty } }

  products.forEach(p => {
    if (p.variants && p.variants.length) {
      p.variants.forEach(v => {
        const productTitle = p.title || '';
        const variantTitle = v.title || p.title || '';
        const sku = v.sku || '';
        const qty = v.inventory_quantity ?? '';

        rows.push([productTitle, variantTitle, sku, qty]);

        if (sku) {
          newSnapshotMap[sku] = {
            product_title: productTitle,
            variant_title: variantTitle,
            qty: qty
          };
        }
      });
    }
  });

  // 3) Log changes between oldSnapshotMap and newSnapshotMap
  logInventoryChanges(ss, oldSnapshotMap, newSnapshotMap);

  // 4) Overwrite snapshot sheet with latest data
  let sheet = ss.getSheetByName(SHOPIFY_CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHOPIFY_CONFIG.SHEET_NAME);
  } else {
    sheet.clearContents();
  }
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  Logger.log(`Wrote ${rows.length - 1} variants to sheet "${SHOPIFY_CONFIG.SHEET_NAME}".`);
}

/**
 * Read current snapshot from ET_inventory and return as a map keyed by SKU.
 * { sku: { product_title, variant_title, qty } }
 */
function getCurrentInventoryMap(ss) {
  const map = {};
  const sheet = ss.getSheetByName(SHOPIFY_CONFIG.SHEET_NAME);
  if (!sheet) return map; // first run: nothing to compare

  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return map; // only header

  const header = values[0];
  const productIdx = header.indexOf('product_title');
  const variantIdx = header.indexOf('variant_title');
  const skuIdx = header.indexOf('sku');
  const qtyIdx = header.indexOf('available_qty');

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sku = row[skuIdx];
    if (!sku) continue;

    map[sku] = {
      product_title: row[productIdx],
      variant_title: row[variantIdx],
      qty: row[qtyIdx]
    };
  }
  return map;
}

/**
 * Log inventory changes into Inventory_Log sheet:
 * Columns: timestamp, product_title, variant_title, sku, old_qty, new_qty, diff
 */
function logInventoryChanges(ss, oldMap, newMap) {
  let logSheet = ss.getSheetByName(INVENTORY_LOG_SHEET);

  // If sheet does not exist, create with header
  if (!logSheet) {
    logSheet = ss.insertSheet(INVENTORY_LOG_SHEET);
  }

  // Ensure header exists
  const header = [
    'timestamp',
    'product_title',
    'variant_title',
    'sku',
    'old_qty',
    'new_qty',
    'diff'
  ];

  const firstRow = logSheet.getRange(1, 1, 1, header.length).getValues()[0];
  const headerMissing = firstRow.join('') === '' || firstRow[0] !== 'timestamp';

  if (headerMissing) {
    logSheet.clearContents();
    logSheet.appendRow(header);
  }

  // Prepare timestamp
  const timestamp = new Date();
  const logRows = [];

  // Compare new vs old snapshot
  Object.keys(newMap).forEach(sku => {
    const newEntry = newMap[sku];
    const oldEntry = oldMap[sku];

    const newQty = Number(newEntry.qty || 0);
    const oldQty = oldEntry ? Number(oldEntry.qty || 0) : null;

    if (!oldEntry) {
      if (newQty !== 0) {
        logRows.push([
          timestamp,
          newEntry.product_title,
          newEntry.variant_title,
          sku,
          '',
          newQty,
          newQty
        ]);
      }
    } else if (oldQty !== newQty) {
      logRows.push([
        timestamp,
        newEntry.product_title,
        newEntry.variant_title,
        sku,
        oldQty,
        newQty,
        newQty - oldQty
      ]);
    }
  });

  // Write rows to sheet
  if (logRows.length > 0) {
    const startRow = logSheet.getLastRow() + 1;
    logSheet.getRange(startRow, 1, logRows.length, header.length).setValues(logRows);
    Logger.log(`Logged ${logRows.length} inventory changes.`);
  } else {
    Logger.log('No inventory changes detected.');
  }
}


/**
 * Fetch all products from Shopify with REST Admin API (paginated with page_info).
 */
function fetchAllShopifyProducts() {
  const allProducts = [];
  let pageInfo = null;
  let hasNextPage = true;

  while (hasNextPage) {
    const url = buildProductsUrl(pageInfo);
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        'X-Shopify-Access-Token': SHOPIFY_CONFIG.ACCESS_TOKEN,
        'Content-Type': 'application/json'
      }
    });

    const code = response.getResponseCode();
    if (code !== 200) {
      throw new Error('Shopify API error: ' + code + ' - ' + response.getContentText());
    }

    const data = JSON.parse(response.getContentText());
    const products = data.products || [];
    allProducts.push(...products);

    const headers = response.getAllHeaders();
    const linkHeader = headers['Link'] || headers['link'];
    if (linkHeader && linkHeader.indexOf('rel="next"') !== -1) {
      const match = linkHeader.match(/<([^>]+)>;\s*rel="next"/);
      if (match && match[1]) {
        const nextUrl = match[1];
        const pageInfoMatch = nextUrl.match(/page_info=([^&]+)/);
        pageInfo = pageInfoMatch ? pageInfoMatch[1] : null;
        hasNextPage = !!pageInfo;
      } else {
        hasNextPage = false;
      }
    } else {
      hasNextPage = false;
    }

    Utilities.sleep(500); // small delay between pages
  }

  Logger.log(`Fetched ${allProducts.length} products from Shopify.`);
  return allProducts;
}

/**
 * Build Shopify products endpoint URL with optional page_info.
 */
function buildProductsUrl(pageInfo) {
  const base = `https://${SHOPIFY_CONFIG.SHOP_NAME}.myshopify.com/admin/api/${SHOPIFY_CONFIG.API_VERSION}/products.json`;
  const params = [
    'limit=250',
    'fields=title,variants'   // we pull product title + variants
  ];
  if (pageInfo) {
    params.push('page_info=' + pageInfo);
  }
  return base + '?' + params.join('&');
}

/**
 * One-time function: create a trigger to run downloadShopifyInventory() every 2 hours.
 */
function createTwoHourTrigger() {
  ScriptApp.newTrigger('downloadShopifyInventory')
    .timeBased()
    .everyHours(2)
    .create();
}

/**
 * Helper: convert a sheet to CSV blob.
 */
function sheetToCsvBlob(ss, sheetName, filename) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found.`);
  }

  const range = sheet.getDataRange();
  const values = range.getDisplayValues();

  const csv = values
    .map(row =>
      row
        .map(value => {
          const v = (value || '').toString().replace(/"/g, '""');
          return `"${v}"`;
        })
        .join(',')
    )
    .join('\r\n');

  return Utilities.newBlob(csv, 'text/csv', filename);
}

/**
 * Send an email with both ET_inventory and Inventory_Log as CSV attachments.
 */
function emailInventoryReport() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const snapshotBlob = sheetToCsvBlob(ss, SHOPIFY_CONFIG.SHEET_NAME, 'ET_inventory.csv');
  const logBlob = sheetToCsvBlob(ss, INVENTORY_LOG_SHEET, 'Inventory_Log.csv');

  const subject = getDynamicEmailSubject();
  
  EMAIL_RECIPIENTS.forEach(to => {
    MailApp.sendEmail({
      to,
      subject: subject,
      body: EMAIL_BODY,
      attachments: [snapshotBlob, logBlob]
    });
  });

  Logger.log(`Inventory report emailed to: ${EMAIL_RECIPIENTS.join(', ')}`);
}

/**
 * One-time function: create a trigger to send the email once per day.
 * Adjust .atHour(8) to whatever hour (0–23) you want.
 */
function createDailyEmailTrigger() {
  ScriptApp.newTrigger('emailInventoryReport')
    .timeBased()
    .everyDays(1)
    .atHour(8) // ~8 AM in your timezone
    .create();
}
