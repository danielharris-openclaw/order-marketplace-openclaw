const SPREADSHEET_ID = '1kTh8RZaAN_jshluKPJoBMhTYt4ERCPRltnhFc7-zUIQ';
const ORDERS_SHEET = 'Заказы';
const RESPONSES_SHEET = 'Отклики';

function doGet() {
  return jsonResponse(listOrders_());
}

function doPost(event) {
  try {
    const payload = JSON.parse((event.postData && event.postData.contents) || '{}');
    if (payload.action === 'list') return jsonResponse(listOrders_());
    if (payload.action === 'createOrder') return jsonResponse(createOrder_(payload));
    if (payload.action === 'createResponse') return jsonResponse(createResponse_(payload));
    if (payload.action === 'closeOrder') return jsonResponse(closeOrder_(payload));
    throw new Error('Неизвестное действие');
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message });
  }
}

function createOrder_(payload) {
  const title = clean_(payload.title);
  const summary = clean_(payload.summary);
  const deadline = clean_(payload.deadline);
  const price = Number(payload.price || 0);
  const phone = clean_(payload.phone);
  if (!title || !summary || !deadline || !phone) {
    throw new Error('Заполните название, суть, срок и телефон');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDERS_SHEET);
    const id = Utilities.getUuid();
    sheet.appendRow([id, new Date().toISOString(), title, summary, deadline, price, phone, 'active']);
    return { ok: true, id };
  } finally {
    lock.releaseLock();
  }
}

function createResponse_(payload) {
  const orderId = clean_(payload.orderId);
  const name = clean_(payload.name);
  const phone = clean_(payload.phone);
  if (!orderId || !name || !phone) {
    throw new Error('Заполните имя и телефон');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ordersSheet = spreadsheet.getSheetByName(ORDERS_SHEET);
    const orderLookup = findOrderRow_(ordersSheet, orderId);
    if (!orderLookup) throw new Error('Заявка не найдена');
    if (orderLookup.order.status === 'closed') throw new Error('Заявка уже закрыта');
    if (orderLookup.order.status === 'in_work') throw new Error('Заявка уже в работе');

    const sheet = spreadsheet.getSheetByName(RESPONSES_SHEET);
    const existingResponses = rowsToObjects_(sheet.getDataRange().getValues())
      .filter(response => String(response.orderId) === orderId);
    if (existingResponses.length) {
      ordersSheet.getRange(orderLookup.rowIndex, orderLookup.statusColumn).setValue('in_work');
      throw new Error('Заявка уже в работе');
    }

    const id = Utilities.getUuid();
    sheet.appendRow([id, orderId, new Date().toISOString(), name, phone]);
    ordersSheet.getRange(orderLookup.rowIndex, orderLookup.statusColumn).setValue('in_work');
    return { ok: true, id };
  } finally {
    lock.releaseLock();
  }
}

function closeOrder_(payload) {
  const orderId = clean_(payload.orderId);
  const customerPhone = normalizePhone_(payload.customerPhone);
  if (!orderId || !customerPhone) {
    throw new Error('Укажите телефон заказчика');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDERS_SHEET);
    const orderLookup = findOrderRow_(sheet, orderId);
    if (!orderLookup) throw new Error('Заявка не найдена');
    if (orderLookup.order.status === 'closed') return { ok: true };
    if (normalizePhone_(orderLookup.order.phone) !== customerPhone) {
      throw new Error('Телефон заказчика не совпадает');
    }
    sheet.getRange(orderLookup.rowIndex, orderLookup.statusColumn).setValue('closed');
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function listOrders_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const orderRows = rowsToObjects_(spreadsheet.getSheetByName(ORDERS_SHEET).getDataRange().getValues());
  const responseRows = rowsToObjects_(spreadsheet.getSheetByName(RESPONSES_SHEET).getDataRange().getValues());
  const responsesByOrder = responseRows.reduce((index, response) => {
    if (!index[response.orderId]) index[response.orderId] = [];
    index[response.orderId].push({
      id: response.id,
      name: response.name,
      phone: response.phone,
      createdAt: response.createdAt
    });
    return index;
  }, {});

  const orders = orderRows
    .map(order => {
      const responses = responsesByOrder[order.id] || [];
      const savedStatus = order.status || 'active';
      const status = savedStatus === 'active' && responses.length ? 'in_work' : savedStatus;
      return {
        id: order.id,
        createdAt: order.createdAt,
        title: order.title,
        summary: order.summary,
        deadline: order.deadline,
        price: Number(order.price || 0),
        phone: order.phone,
        status,
        responses
      };
    });

  return { ok: true, orders };
}

function rowsToObjects_(rows) {
  if (!rows.length) return [];
  const headers = rows[0].map(String);
  return rows.slice(1).filter(row => row.some(Boolean)).map(row => {
    return headers.reduce((object, header, index) => {
      object[header] = row[index];
      return object;
    }, {});
  });
}

function clean_(value) {
  return String(value || '').trim();
}

function normalizePhone_(value) {
  return clean_(value).replace(/\D/g, '');
}

function findOrderRow_(sheet, orderId) {
  const rows = sheet.getDataRange().getValues();
  if (!rows.length) return null;
  const headers = rows[0].map(String);
  const idColumn = headers.indexOf('id') + 1;
  const statusColumn = headers.indexOf('status') + 1;
  if (!idColumn || !statusColumn) throw new Error('В таблице нет колонок id/status');

  for (let index = 1; index < rows.length; index += 1) {
    if (String(rows[index][idColumn - 1]) === orderId) {
      const order = headers.reduce((object, header, columnIndex) => {
        object[header] = rows[index][columnIndex];
        return object;
      }, {});
      return { rowIndex: index + 1, statusColumn, order };
    }
  }
  return null;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
