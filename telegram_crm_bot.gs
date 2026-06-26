// === НАСТРОЙКИ ===
var BOT_TOKEN = "8634819434:AAFIhUF8nScpHK6KPH6wFFu7kemRuoN_gFM";
var SHEET_ID = "136qqtgUzHFDYakHvNkdA_AnOI1VlLYa86ZbL50nbJmk";
var SHEET_NAME = "Продажи";
var CONTENT_SHEET = "Контент";
var TELEGRAM_API = "https://api.telegram.org/bot" + BOT_TOKEN;
var PRICE_PER_GRAM = 0.14;

// === CONTENT API: GET ===
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "";

  if (action === "getContent") {
    var content = loadContentFromSheet();
    return ContentService.createTextOutput(JSON.stringify(content))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "ok", version: "2.0" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// === WEBHOOK (TELEGRAM + CONTENT API) ===
function doPost(e) {
  try {
    if (!e || !e.postData) return;
    var raw = e.postData.contents;
    var data = JSON.parse(raw);

    // Content API: save
    if (data.action === "saveContent" && data.content) {
      saveContentToSheet(data.content);
      return ContentService.createTextOutput(JSON.stringify({ status: "saved" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Production bot: save batch
    if (data.action === "saveBatch" && data.batch) {
      saveBatchToSheet(data.batch, data.stock);
      return ContentService.createTextOutput(JSON.stringify({ status: "saved" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Telegram webhook
    if (data.callback_query) {
      handleCallback(data.callback_query);
      return;
    }

    var msg = data.message;
    if (!msg || !msg.text) return;

    var chatId = msg.chat.id;
    var text = msg.text.trim();

    if (text === "/start" || text === "/sale") {
      startSale(chatId);
      return;
    }
    if (text === "/today") { sendMessage(chatId, getTodaySales()); return; }
    if (text === "/total") { sendMessage(chatId, getMonthTotal()); return; }
    if (text === "/help") {
      sendMessage(chatId,
        "🥩 <b>PicadaRica CRM</b>\n\n" +
        "/sale — новая продажа (с кнопками)\n" +
        "/today — продажи за сегодня\n" +
        "/total — итоги за месяц\n" +
        "/help — эта справка");
      return;
    }

    var state = getState(chatId);
    if (state && state.step === "client") {
      state.client = text;
      state.step = "note";
      setState(chatId, state);
      sendKeyboard(chatId, "🚚 Способ доставки:", [
        [
          { text: "Самовывоз", callback_data: "note_pickup" },
          { text: "Доставка", callback_data: "note_delivery" }
        ],
        [
          { text: "Подарок", callback_data: "note_gift" },
          { text: "Без примечания", callback_data: "note_none" }
        ],
        [
          { text: "✏️ Своё", callback_data: "note_custom" }
        ]
      ]);
      return;
    }
    if (state && state.step === "note") {
      state.note = text;
      saveSale(chatId, state);
      return;
    }
    if (state && state.step === "custom_weight") {
      var w = parseInt(text);
      if (isNaN(w) || w <= 0) {
        sendMessage(chatId, "❌ Введи вес в граммах (число). Пример: 175");
        return;
      }
      state.weight = w;
      state.step = "price";
      setState(chatId, state);
      askPrice(chatId, state);
      return;
    }
    if (state && state.step === "custom_price") {
      var p = parseFloat(text);
      if (isNaN(p) || p < 0) {
        sendMessage(chatId, "❌ Введи сумму в USD (число). Пример: 25");
        return;
      }
      state.amount = p;
      if (state.channel === "Ярмарка") {
        state.client = "—";
        state.note = "";
        saveSale(chatId, state);
      } else {
        state.step = "client";
        setState(chatId, state);
        sendMessage(chatId, "👤 Напиши имя клиента:");
      }
      return;
    }

    startSale(chatId);

  } catch (err) {
    Logger.log("Error: " + err.toString());
  }
}

// === CONTENT STORAGE ===
function getOrCreateContentSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(CONTENT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONTENT_SHEET);
    sheet.getRange(1, 1).setValue("content_json");
    sheet.getRange(1, 2).setValue("updated_at");
    sheet.getRange(2, 1).setValue("{}");
    sheet.getRange(2, 2).setValue(new Date().toISOString());
  }
  return sheet;
}

function loadContentFromSheet() {
  var sheet = getOrCreateContentSheet();
  var raw = sheet.getRange(2, 1).getValue();
  try {
    return JSON.parse(raw);
  } catch (e) {
    return {};
  }
}

function saveContentToSheet(contentObj) {
  var sheet = getOrCreateContentSheet();
  sheet.getRange(2, 1).setValue(JSON.stringify(contentObj));
  sheet.getRange(2, 2).setValue(new Date().toISOString());
}

// === НАЧАЛО ПРОДАЖИ ===
function startSale(chatId) {
  clearState(chatId);
  var state = { step: "channel" };
  setState(chatId, state);

  sendKeyboard(chatId, "📣 Источник продажи:", [
    [
      { text: "👋 Прямая", callback_data: "channel_direct" },
      { text: "🎪 Ярмарка", callback_data: "channel_fair" }
    ]
  ]);
}

// === ОБРАБОТКА КНОПОК ===
function handleCallback(cb) {
  var chatId = cb.message.chat.id;
  var msgId = cb.message.message_id;
  var data = cb.data;

  answerCallback(cb.id);

  var state = getState(chatId) || {};

  if (data.indexOf("product_") === 0) {
    var products = { "product_beef": "Говядина", "product_chicken": "Курица", "product_mix": "Говядина+Курица" };
    state.product = products[data];
    state.step = "weight";
    setState(chatId, state);

    sendKeyboard(chatId, "⚖️ " + state.product + "\nВыбери вес:", [
      [
        { text: "25г", callback_data: "weight_25" },
        { text: "50г", callback_data: "weight_50" },
        { text: "100г", callback_data: "weight_100" }
      ],
      [
        { text: "250г", callback_data: "weight_250" },
        { text: "300г", callback_data: "weight_300" },
        { text: "500г", callback_data: "weight_500" }
      ],
      [
        { text: "✏️ Свой вес", callback_data: "weight_custom" }
      ]
    ]);
    return;
  }

  if (data.indexOf("weight_") === 0) {
    if (data === "weight_custom") {
      state.step = "custom_weight";
      setState(chatId, state);
      sendMessage(chatId, "✏️ Введи вес в граммах:");
      return;
    }
    state.weight = parseInt(data.replace("weight_", ""));
    state.step = "price";
    setState(chatId, state);
    askPrice(chatId, state);
    return;
  }

  if (data.indexOf("price_") === 0) {
    if (data === "price_custom") {
      state.step = "custom_price";
      setState(chatId, state);
      sendMessage(chatId, "✏️ Введи сумму в USD:");
      return;
    }
    if (data === "price_free") {
      state.amount = 0;
    } else {
      state.amount = parseFloat(data.replace("price_", ""));
    }
    if (state.channel === "Ярмарка") {
      state.client = "—";
      state.note = "";
      saveSale(chatId, state);
    } else {
      state.step = "client";
      setState(chatId, state);
      sendMessage(chatId, "👤 Напиши имя клиента:");
    }
    return;
  }

  if (data.indexOf("channel_") === 0) {
    var channels = {
      "channel_direct": "Прямая",
      "channel_fair": "Ярмарка"
    };
    state.channel = channels[data];
    state.step = "product";
    setState(chatId, state);

    sendKeyboard(chatId, "🥩 Выбери продукт:", [
      [{ text: "🐄 Говядина", callback_data: "product_beef" }],
      [{ text: "🐔 Курица", callback_data: "product_chicken" }],
      [{ text: "🔀 Микс", callback_data: "product_mix" }]
    ]);
    return;
  }

  if (data.indexOf("note_") === 0) {
    if (data === "note_custom") {
      state.step = "note";
      setState(chatId, state);
      sendMessage(chatId, "✏️ Напиши примечание:");
      return;
    }
    var notes = {
      "note_pickup": "Самовывоз",
      "note_delivery": "Доставка",
      "note_gift": "Подарок",
      "note_none": ""
    };
    state.note = notes[data];
    saveSale(chatId, state);
    return;
  }
}

// === СПРОСИТЬ ЦЕНУ ===
function askPrice(chatId, state) {
  var suggested = Math.round(state.weight * PRICE_PER_GRAM);
  sendKeyboard(chatId, "💰 " + state.product + " " + state.weight + "г\nВыбери сумму (USD):", [
    [
      { text: "$" + suggested + " (авто)", callback_data: "price_" + suggested }
    ],
    [
      { text: "$0 бесплатно", callback_data: "price_free" },
      { text: "✏️ Своя сумма", callback_data: "price_custom" }
    ]
  ]);
}

// === СПРОСИТЬ КАНАЛ ===
function askChannel(chatId) {
  sendKeyboard(chatId, "📣 Канал продажи:", [
    [
      { text: "👋 Прямая", callback_data: "channel_direct" },
      { text: "✈️ Telegram", callback_data: "channel_telegram" }
    ],
    [
      { text: "📱 WhatsApp", callback_data: "channel_whatsapp" },
      { text: "📸 Instagram", callback_data: "channel_instagram" }
    ],
    [
      { text: "🎪 Ярмарка", callback_data: "channel_fair" },
      { text: "♠️ Покер", callback_data: "channel_poker" }
    ],
    [
      { text: "🛒 MercadoLibre", callback_data: "channel_ml" }
    ]
  ]);
}

// === СОХРАНЕНИЕ ПРОДАЖИ ===
function saveSale(chatId, state) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  var lastRow = sheet.getLastRow();
  var newRow = lastRow + 1;
  var rowNum = lastRow - 3 + 1;
  var orderId = "PR-" + String(rowNum).padStart(3, "0");

  var today = new Date();
  var dateStr = Utilities.formatDate(today, "America/Argentina/Buenos_Aires", "dd.MM.yyyy");

  sheet.getRange(newRow, 1).setValue(orderId);
  sheet.getRange(newRow, 2).setValue(dateStr);
  sheet.getRange(newRow, 3).setValue(state.product);
  sheet.getRange(newRow, 4).setValue(state.weight);
  sheet.getRange(newRow, 5).setValue(state.amount);
  sheet.getRange(newRow, 6).setValue(state.client);
  sheet.getRange(newRow, 7).setValue(state.note || "");
  sheet.getRange(newRow, 8).setValue(state.channel);

  clearState(chatId);

  var msg = "✅ <b>Записано " + orderId + "</b>\n\n" +
    "📦 " + state.product + " " + state.weight + "г\n" +
    "💰 $" + state.amount + "\n" +
    "👤 " + state.client + "\n" +
    "📣 " + state.channel + "\n" +
    (state.note ? "📝 " + state.note : "");

  if (state.channel !== "Ярмарка") {
    var productNames = { "Говядина": "Bife", "Курица": "Pollo", "Говядина+Курица": "Mix" };
    var productES = productNames[state.product] || state.product;
    var noteES = "";
    if (state.note === "Самовывоз") noteES = "Retiro en mano";
    else if (state.note === "Доставка") noteES = "Entrega a domicilio";
    else if (state.note === "Подарок") noteES = "Regalo";
    else if (state.note) noteES = state.note;

    var clientMsg = "✅ PicadaRica — Pedido " + orderId + "\n" +
      "📦 " + productES + " " + state.weight + "g\n" +
      "💰 $" + state.amount + "\n" +
      (noteES ? "🚚 " + noteES + "\n" : "") +
      "¡Gracias! Te avisamos cuando esté listo.";

    msg += "\n\n<b>Клиенту (нажми чтобы скопировать):</b>\n" +
      "<pre>" + clientMsg + "</pre>";
  }

  msg += "\n/sale — ещё одна продажа";
  sendMessage(chatId, msg);
}

// === СОСТОЯНИЕ ДИАЛОГА ===
function getState(chatId) {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty("state_" + chatId);
  if (!raw) return null;
  return JSON.parse(raw);
}

function setState(chatId, state) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty("state_" + chatId, JSON.stringify(state));
}

function clearState(chatId) {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty("state_" + chatId);
}

// === ПРОДАЖИ ЗА СЕГОДНЯ ===
function getTodaySales() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var today = new Date();
  var todayStr = Utilities.formatDate(today, "America/Argentina/Buenos_Aires", "dd.MM.yyyy");
  var sales = [], totalAmount = 0, totalWeight = 0;

  for (var i = 3; i < data.length; i++) {
    if (String(data[i][1]) === todayStr) {
      sales.push(data[i]);
      totalAmount += Number(data[i][4]) || 0;
      totalWeight += Number(data[i][3]) || 0;
    }
  }
  if (sales.length === 0) return "📊 Сегодня (" + todayStr + ")\n\nПродаж пока нет.";

  var text = "📊 <b>Сегодня (" + todayStr + ")</b>\n\n";
  for (var j = 0; j < sales.length; j++) {
    text += "• " + sales[j][2] + " " + sales[j][3] + "г — $" + sales[j][4] + " — " + sales[j][5] + "\n";
  }
  text += "\n💰 Итого: $" + totalAmount + " / " + totalWeight + "г / " + sales.length + " продаж";
  return text;
}

// === ИТОГИ ЗА МЕСЯЦ ===
function getMonthTotal() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var today = new Date();
  var currentMonth = Utilities.formatDate(today, "America/Argentina/Buenos_Aires", "MM.yyyy");
  var totalAmount = 0, totalWeight = 0, count = 0;

  for (var i = 3; i < data.length; i++) {
    var dateStr = String(data[i][1]);
    if (dateStr.length >= 10 && dateStr.substring(3) === currentMonth) {
      totalAmount += Number(data[i][4]) || 0;
      totalWeight += Number(data[i][3]) || 0;
      count++;
    }
  }
  return "📈 <b>Итоги за месяц</b>\n\n" +
    "🛒 Продаж: " + count + "\n" +
    "💰 Выручка: $" + totalAmount + "\n" +
    "⚖️ Вес: " + totalWeight + "г\n" +
    "📊 Средний чек: $" + (count > 0 ? Math.round(totalAmount / count) : 0);
}

// === TELEGRAM API ===
function sendMessage(chatId, text) {
  UrlFetchApp.fetch(TELEGRAM_API + "/sendMessage", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: "HTML" })
  });
}

function sendKeyboard(chatId, text, buttons) {
  UrlFetchApp.fetch(TELEGRAM_API + "/sendMessage", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      chat_id: chatId,
      text: text,
      parse_mode: "HTML",
      reply_markup: { inline_keyboard: buttons }
    })
  });
}

function answerCallback(callbackId) {
  UrlFetchApp.fetch(TELEGRAM_API + "/answerCallbackQuery", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ callback_query_id: callbackId })
  });
}

// === УСТАНОВКА WEBHOOK ===
function setWebhook() {
  var webAppUrl = ScriptApp.getService().getUrl();
  var response = UrlFetchApp.fetch(TELEGRAM_API + "/setWebhook?url=" + webAppUrl);
  Logger.log(response.getContentText());
}

// === ПРОИЗВОДСТВО: ЗАПИСЬ ПАРТИИ ===
function saveBatchToSheet(batch, stock) {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // --- Лист "Производство" ---
  var prodSheet = ss.getSheetByName("Производство");
  if (!prodSheet) {
    prodSheet = ss.insertSheet("Производство");
    prodSheet.getRange(1, 1, 1, 11).setValues([[
      "#", "Дата", "Мясо", "Сырой вес (кг)", "Финал вес (кг)",
      "Выход %", "Сушка (ч)", "25г шт", "50г шт", "100г шт", "250г шт"
    ]]);
    prodSheet.setFrozenRows(1);
    prodSheet.getRange(1, 1, 1, 11).setFontWeight("bold")
      .setBackground("#1F2933").setFontColor("#FFFFFF");
  }

  var packs = batch.packs || {};
  var yieldPct = batch.raw_weight > 0
    ? Math.round(batch.final_weight / batch.raw_weight * 100 * 10) / 10
    : 0;
  var meatName = batch.meat === "🐄" ? "Говядина" : "Курица";
  var dateStr = Utilities.formatDate(new Date(), "America/Argentina/Buenos_Aires", "dd.MM.yyyy");

  var lastRow = prodSheet.getLastRow();
  prodSheet.getRange(lastRow + 1, 1, 1, 11).setValues([[
    batch.id, dateStr, meatName,
    batch.raw_weight, batch.final_weight,
    yieldPct + "%", batch.drying_hours || 0,
    packs["25"] || 0, packs["50"] || 0, packs["100"] || 0, packs["250"] || 0
  ]]);

  // --- Лист "Остатки" ---
  var stockSheet = ss.getSheetByName("Остатки");
  if (!stockSheet) {
    stockSheet = ss.insertSheet("Остатки");
  }
  stockSheet.clearContents();

  var rows = [
    ["Мясо", "25г (шт)", "50г (шт)", "100г (шт)", "250г (шт)", "Итого (г)", "Обновлено"]
  ];
  var meatMap = { "🐄": "Говядина", "🐔": "Курица" };
  var updatedAt = Utilities.formatDate(new Date(), "America/Argentina/Buenos_Aires", "dd.MM.yyyy HH:mm");

  for (var emoji in stock) {
    var s = stock[emoji];
    var total = (s["25"] || 0) * 25 + (s["50"] || 0) * 50 + (s["100"] || 0) * 100 + (s["250"] || 0) * 250;
    rows.push([
      meatMap[emoji] || emoji,
      s["25"] || 0, s["50"] || 0, s["100"] || 0, s["250"] || 0,
      total, updatedAt
    ]);
  }

  stockSheet.getRange(1, 1, rows.length, 7).setValues(rows);
  stockSheet.getRange(1, 1, 1, 7).setFontWeight("bold")
    .setBackground("#1F2933").setFontColor("#FFFFFF");
  stockSheet.setFrozenRows(1);
}

// === ИНИЦИАЛИЗАЦИЯ КОНТЕНТА ===
function initContent() {
  var defaultContent = {
    "brand": "PicadaRica",
    "whatsapp": "541126466244",
    "telegram": "@zakazPRica",
    "instagram": "@picadarica",
    "phone": "+54 011 26466244",
    "year": "2026",
    "hero": {
      "eyebrow": { "es": "Hecho en Buenos Aires, Argentina", "ru": "Сделано в Буэнос-Айресе, Аргентина", "en": "Made in Buenos Aires, Argentina" },
      "title_line1": { "es": "El auténtico", "ru": "Настоящий", "en": "The authentic" },
      "title_line2": { "es": "bife deshidratado", "ru": "вяленый стейк", "en": "dried beef steak" },
      "title_line3": { "es": "de Argentina", "ru": "из Аргентины", "en": "from Argentina" },
      "subtitle": { "es": "36 g de proteína · 0% azúcar · 100% carne vacuna.", "ru": "36 г белка · 0% сахара · 100% говядина.", "en": "36g protein · 0% sugar · 100% beef." },
      "cta_primary": { "es": "Probá ahora", "ru": "Попробуй сейчас", "en": "Try now" },
      "cta_secondary": { "es": "Conocé más", "ru": "Узнать больше", "en": "Learn more" }
    },
    "product": {
      "image": "https://i.postimg.cc/YSvDNYws/Gemini-Generated-Image-2zu6kt2zu6kt2zu6.png",
      "ingredients": { "es": "Carne vacuna (cuadril), sal, especias naturales.", "ru": "Говядина (кадриль), соль, натуральные специи.", "en": "Beef (cuadril), salt, natural spices." },
      "badges": [
        { "es": "Sin soja", "ru": "Без сои", "en": "No soy" },
        { "es": "Sin glutamato", "ru": "Без глутамата", "en": "No glutamate" },
        { "es": "Sin conservantes", "ru": "Без консервантов", "en": "No preservatives" },
        { "es": "Sin aromatizantes", "ru": "Без ароматизаторов", "en": "No flavorings" }
      ],
      "nutrition": { "protein": 36, "fat": 8, "carbs": 6, "sugar": 0, "calories": 300 }
    },
    "pricing": [
      { "id": "small", "weight": "50", "unit": "g", "price": 8, "featured": false, "label": { "es": "Para probar", "ru": "Для пробы", "en": "To try" } },
      { "id": "medium", "weight": "250", "unit": "g", "price": 35, "featured": true, "label": { "es": "El equilibrio perfecto", "ru": "Идеальный баланс", "en": "Perfect balance" } },
      { "id": "large", "weight": "1", "unit": "kg", "price": 120, "featured": false, "label": { "es": "Para verdaderos fanáticos", "ru": "Для настоящих фанатов", "en": "For true fans" } }
    ],
    "guarantee": {
      "title": { "es": "100% garantizado", "ru": "100% гарантия", "en": "100% guaranteed" },
      "subtitle": { "es": "Si no te gusta, lo cambiamos. Sin preguntas.", "ru": "Если не понравится — заменим. Без вопросов.", "en": "If you don't like it, we'll replace it. No questions." },
      "description": { "es": "Probá PicadaRica sin riesgo.", "ru": "Попробуйте PicadaRica без риска.", "en": "Try PicadaRica risk-free." }
    },
    "sections_visible": { "hero": true, "producto": true, "para_quien": true, "proceso": true, "precios": true, "garantia": true, "cta_final": true }
  };
  saveContentToSheet(defaultContent);
  Logger.log("Content initialized");
}
