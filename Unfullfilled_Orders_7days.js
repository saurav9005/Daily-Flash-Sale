/***** CONFIG *****/
const SHOP_NAME = 'elie-tahari';
const ACCESS_TOKEN = '';
const ALERT_EMAILS = [
  'sauravk@elietahari.com'
];

const LOOKBACK_DAYS = 7;
const SHOPIFY_API_VERSION = '2024-01';

/**
 * RUN THIS:
 * Checks archived + unfulfilled orders from the last 7 days
 * and emails both sections in one report.
 */
function generateArchivedAndUnfulfilledLast7DaysReport() {
  const now = new Date();
  const startDate = new Date(now.getTime() - (LOOKBACK_DAYS * 24 * 60 * 60 * 1000));

  const orders = fetchAllOrdersSinceDate(startDate);

  if (!orders.length) {
    Logger.log('No orders found in the last 7 days.');
    return;
  }

  // Exclude cancelled unless you want them included
  const activeOrders = orders.filter(order => !order.cancelled_at);

  const archivedOrders = activeOrders.filter(order => isArchivedOrder_(order));
  const unfulfilledOrders = activeOrders.filter(order => isUnfulfilledOrder_(order));

  sendArchivedAndUnfulfilledEmail_(archivedOrders, unfulfilledOrders, startDate);
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
 * Archived proxy:
 * Shopify Order REST does not expose a simple "archived: true/false" field here,
 * so we use closed_at as the practical archived/closed indicator.
 */
function isArchivedOrder_(order) {
  return !!order.closed_at;
}

/**
 * Unfulfilled logic
 */
function isUnfulfilledOrder_(order) {
  return (
    order.fulfillment_status === null ||
    order.fulfillment_status === 'unfulfilled'
  );
}

/**
 * Send one email with both archived + unfulfilled sections
 */
function sendArchivedAndUnfulfilledEmail_(archivedOrders, unfulfilledOrders, startDate) {
  const subject =
    `Elie Tahari | Archived + Unfulfilled Orders | Last ${LOOKBACK_DAYS} Days | ` +
    `Archived: ${archivedOrders.length} | Unfulfilled: ${unfulfilledOrders.length}`;

  let html = `
    <p>
      <b>Orders report</b> for the last <b>${LOOKBACK_DAYS} days</b><br>
      Start date: <b>${formatDateYMD_(startDate)}</b>
    </p>
    <p>
      <b>Archived Orders:</b> ${archivedOrders.length}<br>
      <b>Unfulfilled Orders:</b> ${unfulfilledOrders.length}
    </p>
  `;

  // Archived section
  html += `
    <h3>Archived Orders</h3>
    <table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse; width:100%;">
      <tr>
        <th>Order #</th>
        <th>Created At</th>
        <th>Closed At</th>
        <th>Financial</th>
        <th>Fulfillment</th>
        <th>Customer</th>
        <th>Total</th>
        <th>Admin Link</th>
      </tr>
  `;

  if (archivedOrders.length) {
    archivedOrders.forEach(order => {
      const o = toOrderRow_(order);
      const adminLink = `https://${SHOP_NAME}.myshopify.com/admin/orders/${o.id}`;
      html += `
        <tr>
          <td>${escapeHtml_(o.name)}</td>
          <td>${escapeHtml_(o.created_at)}</td>
          <td>${escapeHtml_(order.closed_at || '')}</td>
          <td>${escapeHtml_(o.financial_status)}</td>
          <td>${escapeHtml_(o.fulfillment_status)}</td>
          <td>${escapeHtml_(o.customer_name)}</td>
          <td>${escapeHtml_(o.total_price)}</td>
          <td><a href="${adminLink}">Open</a></td>
        </tr>
      `;
    });
  } else {
    html += `
      <tr>
        <td colspan="8">No archived orders found in the last ${LOOKBACK_DAYS} days.</td>
      </tr>
    `;
  }

  html += `</table><br><br>`;

  // Unfulfilled section
  html += `
    <h3>Unfulfilled Orders</h3>
    <table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse; width:100%;">
      <tr>
        <th>Order #</th>
        <th>Created At</th>
        <th>Financial</th>
        <th>Fulfillment</th>
        <th>Customer</th>
        <th>Email</th>
        <th>Total</th>
        <th>Admin Link</th>
      </tr>
  `;

  if (unfulfilledOrders.length) {
    unfulfilledOrders.forEach(order => {
      const o = toOrderRow_(order);
      const adminLink = `https://${SHOP_NAME}.myshopify.com/admin/orders/${o.id}`;
      html += `
        <tr>
          <td>${escapeHtml_(o.name)}</td>
          <td>${escapeHtml_(o.created_at)}</td>
          <td>${escapeHtml_(o.financial_status)}</td>
          <td>${escapeHtml_(o.fulfillment_status)}</td>
          <td>${escapeHtml_(o.customer_name)}</td>
          <td>${escapeHtml_(o.customer_email)}</td>
          <td>${escapeHtml_(o.total_price)}</td>
          <td><a href="${adminLink}">Open</a></td>
        </tr>
      `;
    });
  } else {
    html += `
      <tr>
        <td colspan="8">No unfulfilled orders found in the last ${LOOKBACK_DAYS} days.</td>
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
