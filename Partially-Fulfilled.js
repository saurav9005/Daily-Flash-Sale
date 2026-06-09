/***** CONFIG *****/
const SHOP_NAME = 'elie-tahari';
const ACCESS_TOKEN = '';
const ALERT_EMAILS = [
  'sauravk@elietahari.com'
];

const START_DATE = '2026-03-01T00:00:00Z';
const SHOPIFY_API_VERSION = '2024-01';

/**
 * RUN THIS:
 * Checks partially fulfilled orders from March 1, 2026
 * where at least one line item SKU starts with "PROTECT-"
 * and emails the report.
 */
function generateProtectPartialFulfilledOrdersReport() {
  const startDate = new Date(START_DATE);

  const orders = fetchAllOrdersSinceDate(startDate);

  if (!orders.length) {
    Logger.log('No orders found since March 1, 2026.');
    return;
  }

  // Exclude cancelled orders
  const activeOrders = orders.filter(order => !order.cancelled_at);

  // Only partially fulfilled orders with PROTECT- SKU
  const partiallyFulfilledProtectOrders = activeOrders.filter(order =>
    isPartiallyFulfilledProtectOrder_(order)
  );

  sendProtectPartialFulfilledEmail_(partiallyFulfilledProtectOrders, startDate);
}

/**
 * Fetch ALL orders from Shopify since a given date (pagination included)
 */
function fetchAllOrdersSinceDate(createdAtMinDate) {
  const createdAtMinIso = createdAtMinDate.toISOString();

  const all = [];
  let nextPageUrl = buildOrdersUrl_({
    status: 'any',
    limit: 250,
    created_at_min: createdAtMinIso,
    order: 'created_at asc'
  });

  while (nextPageUrl) {
    const resp = shopifyGet_(nextPageUrl);
    if (!resp || !resp.body) break;

    const data = JSON.parse(resp.body);
    const batch = data.orders || [];
    all.push.apply(all, batch);

    nextPageUrl = parseNextPageUrl_(resp.headers);
  }

  return all;
}

/**
 * Returns true only if:
 * 1. Order is partially fulfilled
 * 2. At least one line item has SKU starting with "PROTECT-"
 */
function isPartiallyFulfilledProtectOrder_(order) {
  const isPartiallyFulfilled = order.fulfillment_status === 'partial';

  if (!isPartiallyFulfilled) return false;

  return (order.line_items || []).some(item =>
    item.sku && String(item.sku).startsWith('PROTECT-')
  );
}

/**
 * Get all PROTECT- line items from an order
 */
function getProtectItems_(order) {
  return (order.line_items || []).filter(item =>
    item.sku && String(item.sku).startsWith('PROTECT-')
  );
}

/**
 * Send email report
 */
function sendProtectPartialFulfilledEmail_(orders, startDate) {
  const subject =
    `Elie Tahari | Partially Fulfilled PROTECT Orders | Since ${formatDateYMD_(startDate)} | Count: ${orders.length}`;

  let html = `
    <p>
      <b>Partially Fulfilled PROTECT Orders Report</b><br>
      Start date: <b>${formatDateYMD_(startDate)}</b>
    </p>
    <p>
      <b>Total matching orders:</b> ${orders.length}
    </p>
  `;

  html += `
    <h3>Orders</h3>
    <table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse; width:100%;">
      <tr>
        <th>Order #</th>
        <th>Created At</th>
        <th>Financial</th>
        <th>Fulfillment</th>
        <th>Customer</th>
        <th>Email</th>
        <th>Total</th>
        <th>PROTECT SKU(s)</th>
        <th>PROTECT Qty</th>
        <th>Admin Link</th>
      </tr>
  `;

  if (orders.length) {
    orders.forEach(order => {
      const o = toOrderRow_(order);
      const adminLink = `https://${SHOP_NAME}.myshopify.com/admin/orders/${o.id}`;
      const protectItems = getProtectItems_(order);
      const protectSkus = protectItems.map(item => item.sku).join(', ');
      const protectQty = protectItems.map(item => `${item.sku} (${item.quantity})`).join(', ');

      html += `
        <tr>
          <td>${escapeHtml_(o.name)}</td>
          <td>${escapeHtml_(o.created_at)}</td>
          <td>${escapeHtml_(o.financial_status)}</td>
          <td>${escapeHtml_(o.fulfillment_status)}</td>
          <td>${escapeHtml_(o.customer_name)}</td>
          <td>${escapeHtml_(o.customer_email)}</td>
          <td>${escapeHtml_(o.total_price)}</td>
          <td>${escapeHtml_(protectSkus)}</td>
          <td>${escapeHtml_(protectQty)}</td>
          <td><a href="${adminLink}">Open</a></td>
        </tr>
      `;
    });
  } else {
    html += `
      <tr>
        <td colspan="10">No partially fulfilled orders with PROTECT- SKU found since ${formatDateYMD_(startDate)}.</td>
      </tr>
    `;
  }

  html += `</table>`;

  MailApp.sendEmail({
    to: ALERT_EMAILS.join(','),
    subject: subject,
    htmlBody: html,
    body: 'Please view this email in HTML to see the order list.'
  });
}

/** ---------------------------
 * Shopify helpers
 * --------------------------- */
function buildOrdersUrl_(queryObj) {
  const baseUrl = `https://${SHOP_NAME}.myshopify.com/admin/api/${SHOPIFY_API_VERSION}/orders.json`;
  const params = [];

  Object.keys(queryObj).forEach(key => {
    if (queryObj[key] === undefined || queryObj[key] === null) return;
    params.push(`${encodeURIComponent(key)}=${encodeURIComponent(String(queryObj[key]))}`);
  });

  return `${baseUrl}?${params.join('&')}`;
}

function shopifyGet_(url) {
  const options = {
    method: 'get',
    headers: {
      'X-Shopify-Access-Token': ACCESS_TOKEN,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  if (code !== 200) {
    Logger.log('Error fetching orders: ' + code + ' ' + response.getContentText());
    return null;
  }

  return {
    body: response.getContentText(),
    headers: response.getAllHeaders()
  };
}

function parseNextPageUrl_(headers) {
  const link = headers && (headers.Link || headers.link);
  if (!link) return null;

  const parts = String(link).split(',');
  for (var i = 0; i < parts.length; i++) {
    const p = parts[i];
    if (p.indexOf('rel="next"') !== -1) {
      const match = p.match(/<([^>]+)>/);
      return match && match[1] ? match[1] : null;
    }
  }
  return null;
}

/** ---------------------------
 * Utilities
 * --------------------------- */
function toOrderRow_(order) {
  return {
    id: String(order.id),
    name: order.name || ('#' + order.id),
    created_at: order.created_at || '',
    total_price: order.total_price || '',
    financial_status: order.financial_status || '',
    fulfillment_status: order.fulfillment_status === null ? 'unfulfilled' : String(order.fulfillment_status),
    customer_email: order.email || (order.customer && order.customer.email) || 'N/A',
    customer_name: order.customer
      ? [order.customer.first_name, order.customer.last_name].filter(Boolean).join(' ')
      : 'Guest'
  };
}

function formatDateYMD_(d) {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function escapeHtml_(s) {
  s = s === null || s === undefined ? '' : String(s);
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
