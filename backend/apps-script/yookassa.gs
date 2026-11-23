/**
 * YooKassa helpers for the existing E-Hookah Google Apps Script.
 *
 * This file intentionally does not define doGet(e) or doPost(e).
 * See README.md before connecting it to the existing web-app handlers.
 *
 * TEST MODE ONLY: live payments are hard-blocked.
 * Fiscal receipts are not configured by this module.
 */

var YK_PAYMENT_HEADERS_ = [
  'Order ID', 'Created At', 'Client Request ID', 'Idempotence Key',
  'Payment ID', 'Amount', 'Currency', 'Payment Status', 'Confirmation URL',
  'Paid At', 'Test Mode', 'Cart JSON', 'Status Token', 'Customer Name',
  'Customer Phone', 'Customer Address', 'Customer Comment', 'Last Error'
];

function ykRoutePost_(e) {
  var body;
  try {
    body = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents)
      : null;
  } catch (error) {
    return null;
  }

  if (!body || body.action !== 'create_yookassa_payment') return null;

  try {
    return ykJsonOutput_({
      success: true,
      payment: ykCreatePayment_(body)
    });
  } catch (error) {
    console.error('YooKassa create payment failed:', error);
    return ykJsonOutput_({
      success: false,
      message: error && error.publicMessage
        ? error.publicMessage
        : 'Не удалось создать платёж. Попробуйте позже.'
    });
  }
}

function ykRouteGet_(e) {
  var params = e && e.parameter ? e.parameter : {};
  if (params.action !== 'yookassa_payment_status') return null;

  try {
    var result = ykGetPublicPaymentStatus_(params.order_id, params.status_token);
    var callback = String(params.callback || '');

    if (callback) {
      if (!/^[A-Za-z_$][A-Za-z0-9_$]{0,100}$/.test(callback)) {
        throw ykPublicError_('Некорректный callback.');
      }
      return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ykJsonOutput_(result);
  } catch (error) {
    return ykJsonOutput_({
      success: false,
      message: error && error.publicMessage
        ? error.publicMessage
        : 'Не удалось получить статус платежа.'
    });
  }
}

function ykCreatePayment_(request) {
  var config = ykGetConfig_();
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var cart = ykValidateCart_(request.items);
    var customer = ykValidateCustomer_(request);
    var requestId = ykValidateRequestId_(request.request_id);
    var catalog = ykReadCatalog_(config);
    var calculated = ykCalculateAmount_(cart, catalog);
    var sheet = ykGetPaymentsSheet_(config);
    var existing = ykFindPaymentByRequestId_(sheet, requestId);

    if (existing) {
      var sameOrder = existing.amount === calculated.amount &&
        existing.currency === config.currency &&
        existing.cartJson === JSON.stringify(cart) &&
        existing.customerName === customer.name &&
        existing.customerPhone === customer.phone &&
        existing.customerAddress === customer.address &&
        existing.customerComment === customer.comment;

      if (!sameOrder) {
        throw ykPublicError_(
          'Идентификатор запроса уже использован. Обновите страницу и повторите попытку.'
        );
      }

      if (existing.paymentId && existing.confirmationUrl) {
        return ykPaymentResponse_(existing, true);
      }

      return ykSendPayment_(
        config, sheet, existing.row, existing.orderId, existing.idempotenceKey,
        existing.statusToken, existing.amount, requestId
      );
    }

    var orderId = Utilities.getUuid();
    var idempotenceKey = Utilities.getUuid();
    var statusToken = Utilities.getUuid() + Utilities.getUuid();

    sheet.appendRow([
      orderId,
      new Date(),
      requestId,
      idempotenceKey,
      '',
      calculated.amount,
      config.currency,
      'creating',
      '',
      '',
      true,
      JSON.stringify(cart),
      statusToken,
      customer.name,
      customer.phone,
      customer.address,
      customer.comment,
      ''
    ]);

    return ykSendPayment_(
      config, sheet, sheet.getLastRow(), orderId, idempotenceKey,
      statusToken, calculated.amount, requestId
    );
  } finally {
    lock.releaseLock();
  }
}

function ykSendPayment_(config, sheet, rowNumber, orderId,
                        idempotenceKey, statusToken, amount, requestId) {
  var paymentRequest = {
    amount: { value: amount, currency: config.currency },
    capture: true,
    confirmation: {
      type: 'redirect',
      return_url: ykBuildReturnUrl_(config.returnUrl, orderId, statusToken)
    },
    description: 'Заказ ' + orderId,
    metadata: { order_id: orderId, request_id: requestId }
  };

  try {
    var response = ykApiRequest_(
      'post', '/payments', paymentRequest, idempotenceKey, config
    );

    if (!response.id || !response.confirmation ||
        !response.confirmation.confirmation_url || response.test !== true) {
      throw new Error('Unexpected or non-test YooKassa payment response.');
    }

    var row = sheet.getRange(rowNumber, 1, 1, YK_PAYMENT_HEADERS_.length).getValues()[0];
    row[4] = response.id;
    row[7] = response.status || 'pending';
    row[8] = response.confirmation.confirmation_url;
    row[9] = response.paid ? (response.captured_at || new Date()) : '';
    row[17] = '';
    sheet.getRange(rowNumber, 1, 1, YK_PAYMENT_HEADERS_.length).setValues([row]);

    return {
      order_id: orderId,
      payment_id: response.id,
      status: response.status || 'pending',
      confirmation_url: response.confirmation.confirmation_url,
      status_token: statusToken,
      reused: false
    };
  } catch (error) {
    sheet.getRange(rowNumber, 8).setValue('creation_failed');
    sheet.getRange(rowNumber, 18).setValue(String(error.message || error).slice(0, 500));
    throw ykPublicError_('Платёжный сервис временно недоступен. Попробуйте позже.');
  }
}

function ykPaymentResponse_(record, reused) {
  return {
    order_id: record.orderId,
    payment_id: record.paymentId,
    status: record.status,
    confirmation_url: record.confirmationUrl,
    status_token: record.statusToken,
    reused: Boolean(reused)
  };
}

function ykBuildReturnUrl_(baseUrl, orderId, statusToken) {
  var separator = baseUrl.indexOf('?') === -1 ? '?' : '&';
  return baseUrl + separator +
    'order_id=' + encodeURIComponent(orderId) +
    '&payment_status_token=' + encodeURIComponent(statusToken);
}

/**
 * Apps Script ContentService serves responses through a Google redirect URL.
 * To avoid relying on redirect handling for YooKassa notifications, pending
 * payment states are reconciled by querying YooKassa from a time-driven trigger.
 */
function ykReconcilePendingPayments_() {
  var config = ykGetConfig_();
  var sheet = ykGetPaymentsSheet_(config);
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return;

  var checked = 0;
  for (var i = values.length - 1; i >= 1 && checked < 20; i -= 1) {
    var paymentId = String(values[i][4] || '');
    var status = String(values[i][7] || '');
    if (!paymentId || (status !== 'pending' && status !== 'waiting_for_capture')) {
      continue;
    }

    checked += 1;
    var rowNumber = i + 1;

    try {
      var payment = ykApiRequest_(
        'get', '/payments/' + encodeURIComponent(paymentId), null, null, config
      );

      if (!payment.metadata ||
          payment.metadata.order_id !== String(values[i][0]) ||
          !payment.amount ||
          String(payment.amount.value) !== String(values[i][5]) ||
          payment.amount.currency !== String(values[i][6]) ||
          payment.test !== true) {
        throw new Error('YooKassa payment verification failed.');
      }

      if (payment.status === 'succeeded' && payment.paid === true) {
        sheet.getRange(rowNumber, 8).setValue('succeeded');
        sheet.getRange(rowNumber, 10).setValue(payment.captured_at || new Date());
        sheet.getRange(rowNumber, 18).setValue('');
      } else if (payment.status === 'canceled') {
        sheet.getRange(rowNumber, 8).setValue('canceled');
        sheet.getRange(rowNumber, 18).setValue('');
      } else if (payment.status === 'waiting_for_capture') {
        sheet.getRange(rowNumber, 8).setValue('waiting_for_capture');
      }
    } catch (error) {
      console.error('YooKassa status check failed for ' + paymentId + ':', error);
      sheet.getRange(rowNumber, 18).setValue(String(error.message || error).slice(0, 500));
    }
  }
}

function ykInstallPaymentReconciliationTrigger_() {
  var handler = 'ykReconcilePendingPayments_';
  var exists = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === handler;
  });

  if (!exists) {
    ScriptApp.newTrigger(handler).timeBased().everyMinutes(5).create();
  }
}

function ykRemovePaymentReconciliationTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'ykReconcilePendingPayments_') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function ykGetPublicPaymentStatus_(orderId, statusToken) {
  var validOrderId = String(orderId || '').trim();
  var validToken = String(statusToken || '').trim();
  if (!validOrderId || !validToken) throw ykPublicError_('Платёж не найден.');

  var config = ykGetConfig_();
  var sheet = ykGetPaymentsSheet_(config);
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i += 1) {
    if (String(values[i][0]) === validOrderId &&
        String(values[i][12]) === validToken) {
      return {
        success: true,
        payment: {
          order_id: validOrderId,
          status: String(values[i][7] || 'pending'),
          amount: String(values[i][5]),
          currency: String(values[i][6])
        }
      };
    }
  }

  throw ykPublicError_('Платёж не найден.');
}

function ykGetConfig_() {
  var properties = PropertiesService.getScriptProperties();
  var config = {
    enabled: properties.getProperty('YK_PAYMENTS_ENABLED') === 'true',
    mode: properties.getProperty('YK_MODE') || 'disabled',
    shopId: properties.getProperty('YK_SHOP_ID') || '',
    secretKey: properties.getProperty('YK_SECRET_KEY') || '',
    currency: properties.getProperty('YK_CURRENCY') || '',
    returnUrl: properties.getProperty('YK_RETURN_URL') || '',
    spreadsheetId: properties.getProperty('YK_CATALOG_SPREADSHEET_ID') || '',
    catalogSheetName: properties.getProperty('YK_CATALOG_SHEET_NAME') || '',
    paymentsSheetName: properties.getProperty('YK_PAYMENTS_SHEET_NAME') || 'YK_Payments'
  };

  if (!config.enabled) throw ykPublicError_('Тестовая оплата пока отключена.');
  if (config.mode !== 'test') {
    throw ykPublicError_('В этом модуле разрешён только тестовый режим ЮKassa.');
  }
  if (!config.shopId || !config.secretKey) {
    throw ykPublicError_('Не настроены тестовые учётные данные ЮKassa.');
  }
  if (!/^[A-Z]{3}$/.test(config.currency)) {
    throw ykPublicError_('Укажите согласованную трёхбуквенную валюту магазина.');
  }
  if (!/^https:\/\//i.test(config.returnUrl)) {
    throw ykPublicError_('Не настроен HTTPS-адрес возврата после оплаты.');
  }
  if (!config.spreadsheetId || !config.catalogSheetName) {
    throw ykPublicError_('Не настроена таблица и вкладка каталога Google Sheets.');
  }

  return config;
}

function ykReadCatalog_(config) {
  var spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
  var sheet = spreadsheet.getSheetByName(config.catalogSheetName);
  if (!sheet) throw ykPublicError_('Не найден лист каталога товаров.');

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) throw ykPublicError_('Каталог товаров пуст.');

  var headers = values[0].map(function(value) { return String(value).trim(); });
  var nameIndex = headers.indexOf('Название');
  var priceHeader = 'Цена, ' + config.currency;
  var priceIndex = headers.indexOf(priceHeader);
  var statusIndex = headers.indexOf('Статус');

  if (nameIndex < 0 || priceIndex < 0 || statusIndex < 0) {
    throw ykPublicError_(
      'В каталоге должны быть столбцы «Название», «' + priceHeader +
      '» и «Статус». Валюта должна совпадать с валютой магазина ЮKassa.'
    );
  }

  var products = {};
  var usedIds = {};

  values.slice(1).forEach(function(row) {
    var name = String(row[nameIndex] || '').trim();
    if (!name) return;

    var baseId = ykGetProductId_(name);
    var id = baseId;
    var suffix = 2;
    while (usedIds[id]) {
      id = baseId + '-' + suffix;
      suffix += 1;
    }
    usedIds[id] = true;

    var rawPrice = row[priceIndex];
    var price = typeof rawPrice === 'number'
      ? rawPrice
      : Number(String(rawPrice).trim().replace(',', '.'));
    var status = String(row[statusIndex] || '').trim();

    if (!Number.isFinite(price) || price <= 0) return;
    products[id] = { name: name, price: price, status: status };
  });

  return products;
}

function ykGetProductId_(name) {
  var normalized = String(name).trim().toLowerCase();
  var known = {
    'e-hookah air one': 'air-one',
    'e-hookah air one pro': 'air-one-pro'
  };
  return known[normalized] || 'product';
}

function ykValidateCustomer_(request) {
  var name = String(request.name || '').trim();
  var phone = String(request.phone || '').trim();
  var address = String(request.address || '').trim();
  var comment = String(request.comment || '').trim();

  if (name.length < 2 || name.length > 80 ||
      !/^[\p{L}\p{M}][\p{L}\p{M}'’ -]*$/u.test(name)) {
    throw ykPublicError_('Укажите корректное ФИО.');
  }
  if (!/^[+]?\d[\d\s().-]{6,28}\d$/.test(phone)) {
    throw ykPublicError_('Укажите корректный номер телефона.');
  }
  if (address.length < 5 || address.length > 200) {
    throw ykPublicError_('Укажите корректный адрес.');
  }
  if (comment.length > 500) {
    throw ykPublicError_('Комментарий слишком длинный.');
  }

  return { name: name, phone: phone, address: address, comment: comment };
}

function ykValidateCart_(items) {
  if (!Array.isArray(items) || items.length === 0 || items.length > 80) {
    throw ykPublicError_('Корзина пуста или содержит слишком много позиций.');
  }

  var seen = {};
  return items.map(function(item) {
    var id = item && String(item.id || '').trim();
    var quantity = Number(item && item.quantity);

    if (!id || seen[id] || !Number.isInteger(quantity) ||
        quantity < 1 || quantity > 99) {
      throw ykPublicError_('Проверьте состав заказа и количество товаров.');
    }

    seen[id] = true;
    return { id: id, quantity: quantity };
  });
}

function ykCalculateAmount_(cart, catalog) {
  var total = 0;

  cart.forEach(function(item) {
    var product = catalog[item.id];
    if (!product || product.status !== 'В наличии') {
      throw ykPublicError_('Один из товаров недоступен. Обновите каталог.');
    }
    total += product.price * item.quantity;
  });

  if (!Number.isFinite(total) || total <= 0) {
    throw ykPublicError_('Не удалось рассчитать сумму заказа.');
  }

  return { amount: total.toFixed(2) };
}

function ykGetPaymentsSheet_(config) {
  var spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
  var sheet = spreadsheet.getSheetByName(config.paymentsSheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(config.paymentsSheetName);
    sheet.getRange(1, 1, 1, YK_PAYMENT_HEADERS_.length)
      .setValues([YK_PAYMENT_HEADERS_]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function ykFindPaymentByRequestId_(sheet, requestId) {
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i += 1) {
    if (String(values[i][2]) !== requestId) continue;

    return {
      row: i + 1,
      orderId: String(values[i][0]),
      idempotenceKey: String(values[i][3]),
      paymentId: String(values[i][4]),
      amount: String(values[i][5]),
      currency: String(values[i][6]),
      status: String(values[i][7]),
      confirmationUrl: String(values[i][8]),
      cartJson: String(values[i][11]),
      statusToken: String(values[i][12]),
      customerName: String(values[i][13]),
      customerPhone: String(values[i][14]),
      customerAddress: String(values[i][15]),
      customerComment: String(values[i][16])
    };
  }

  return null;
}

function ykApiRequest_(method, path, payload, idempotenceKey, config) {
  var headers = {
    Authorization: 'Basic ' +
      Utilities.base64Encode(config.shopId + ':' + config.secretKey),
    Accept: 'application/json'
  };

  if (idempotenceKey) headers['Idempotence-Key'] = idempotenceKey;
  if (payload) headers['Content-Type'] = 'application/json';

  var options = {
    method: method,
    headers: headers,
    muteHttpExceptions: true
  };
  if (payload) options.payload = JSON.stringify(payload);

  var response = UrlFetchApp.fetch(
    'https://api.yookassa.ru/v3' + path,
    options
  );

  var statusCode = response.getResponseCode();
  var text = response.getContentText();
  var result;

  try {
    result = text ? JSON.parse(text) : {};
  } catch (error) {
    throw new Error('ЮKassa вернула некорректный JSON.');
  }

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error(
      'YooKassa API error: ' +
      (result.description || result.message || ('HTTP ' + statusCode))
    );
  }

  return result;
}

function ykValidateRequestId_(value) {
  var id = String(value || '').trim();
  if (!/^[a-zA-Z0-9_-]{8,80}$/.test(id)) {
    throw ykPublicError_(
      'Не удалось идентифицировать запрос. Обновите страницу и повторите попытку.'
    );
  }
  return id;
}

function ykJsonOutput_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function ykPublicError_(message) {
  var error = new Error(message);
  error.publicMessage = message;
  return error;
}
