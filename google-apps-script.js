const SPREADSHEET_ID = '1kTh8RZaAN_jshluKPJoBMhTYt4ERCPRltnhFc7-zUIQ';
const ORDERS_SHEET = 'Заказы';
const RESPONSES_SHEET = 'Отклики';
const SITE_URL = 'https://danielharris-openclaw.github.io/';

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
    if (payload.action === 'testMax') return jsonResponse(testMaxNotification());
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
    notifyMaxNewOrder_({ id, title, summary, deadline, price, phone });
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
    notifyMaxOrderStatus_('Заказ взят в работу', {
      id: orderLookup.order.id,
      title: orderLookup.order.title,
      summary: orderLookup.order.summary,
      deadline: orderLookup.order.deadline,
      price: orderLookup.order.price,
      phone: orderLookup.order.phone,
      status: 'in_work',
      responses: [{ id, name, phone, createdAt: new Date().toISOString() }]
    });
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
    const responses = rowsToObjects_(
      SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSES_SHEET).getDataRange().getValues()
    ).filter(response => String(response.orderId) === orderId);
    notifyMaxOrderStatus_('Заказ закрыт', {
      id: orderLookup.order.id,
      title: orderLookup.order.title,
      summary: orderLookup.order.summary,
      deadline: orderLookup.order.deadline,
      price: orderLookup.order.price,
      phone: orderLookup.order.phone,
      status: 'closed',
      responses
    });
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

function notifyMaxNewOrder_(order) {
  notifyMaxOrderStatus_('Новый заказ создан', {
    id: order.id,
    title: order.title,
    summary: order.summary,
    deadline: order.deadline,
    price: order.price,
    phone: order.phone,
    status: 'active',
    responses: []
  });
}

function notifyMaxOrderStatus_(title, order) {
  const text = formatOrderMessage_(title, order);
  sendMaxMessage_(text);
}

function sendActiveOrdersDigest() {
  const orders = listOrders_().orders.filter(order => order.status !== 'closed');
  const text = orders.length
    ? [
      'Сводка актуальных заказов',
      'Незакрытых заказов: ' + orders.length,
      '',
      orders.map((order, index) => formatDigestOrder_(order, index + 1)).join('\n\n'),
      '',
      'Сайт: ' + SITE_URL
    ].join('\n')
    : [
      'Сводка актуальных заказов',
      '',
      'Сейчас незакрытых заказов нет.',
      '',
      'Сайт: ' + SITE_URL
    ].join('\n');
  sendMaxMessage_(text);
}

function testMaxNotification() {
  const result = sendMaxMessage_('Тест из Google Apps Script: свойства MAX прочитаны, отправка в группу работает.');
  return {
    ok: result.ok,
    hasToken: result.hasToken,
    hasChatId: result.hasChatId,
    status: result.status,
    response: result.response
  };
}

function setupDailyMaxDigest() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'sendActiveOrdersDigest')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('sendActiveOrdersDigest')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(0)
    .inTimezone('Europe/Moscow')
    .create();
}

function formatOrderMessage_(title, order) {
  const performer = firstPerformer_(order);
  return [
    title,
    '',
    'Статус: ' + statusLabel_(order.status),
    'Название: ' + order.title,
    'Суть: ' + (order.summary || 'не указана'),
    'Срок: ' + order.deadline,
    'Цена: ' + formatPrice_(order.price),
    'Заказчик: ' + order.phone,
    'Исполнитель: ' + (performer ? performer.name + ' · ' + performer.phone : 'пока не назначен'),
    '',
    'Сайт: ' + SITE_URL
  ].join('\n');
}

function formatDigestOrder_(order, number) {
  const performer = firstPerformer_(order);
  return [
    number + '. ' + order.title,
    'Статус: ' + statusLabel_(order.status),
    'Срок: ' + order.deadline,
    'Цена: ' + formatPrice_(order.price),
    'Заказчик: ' + order.phone,
    'Исполнитель: ' + (performer ? performer.name + ' · ' + performer.phone : 'пока не назначен')
  ].join('\n');
}

function firstPerformer_(order) {
  return order.responses && order.responses.length ? order.responses[0] : null;
}

function statusLabel_(status) {
  return {
    active: 'актуальна',
    in_work: 'в работе',
    closed: 'закрыта'
  }[status] || 'актуальна';
}

function sendMaxMessage_(text) {
  const properties = PropertiesService.getScriptProperties();
  const token = properties.getProperty('MAX_BOT_TOKEN');
  const chatId = properties.getProperty('MAX_CHAT_ID');
  if (!token || !chatId) {
    console.log('MAX notification skipped: missing script properties');
    return {
      ok: false,
      hasToken: Boolean(token),
      hasChatId: Boolean(chatId),
      status: 'missing_properties',
      response: ''
    };
  }

  try {
    const response = UrlFetchApp.fetch('https://platform-api.max.ru/messages?chat_id=' + encodeURIComponent(chatId), {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: token
      },
      payload: JSON.stringify({
        text,
        notify: true
      }),
      muteHttpExceptions: true
    });
    const status = response.getResponseCode();
    const body = response.getContentText();
    console.log('MAX notification response: ' + status + ' ' + body);
    return {
      ok: status >= 200 && status < 300,
      hasToken: true,
      hasChatId: true,
      status,
      response: body.slice(0, 500)
    };
  } catch (error) {
    console.log('MAX notification failed: ' + error.message);
    return {
      ok: false,
      hasToken: true,
      hasChatId: true,
      status: 'exception',
      response: error.message
    };
  }
}

function formatPrice_(value) {
  const number = Number(value || 0);
  return number.toLocaleString('ru-RU') + ' ₽';
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
