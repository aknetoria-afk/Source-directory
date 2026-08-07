/**
 * Реляционная перестройка "Справочник источников" -> Sources / Phones / Urls
 *
 * НАСТРОЙКА ПЕРЕД ПУБЛИКАЦИЕЙ:
 * 1. Замените SECRET_TOKEN ниже на любую свою случайную строку (пароль).
 * 2. Deploy -> New deployment -> Type: Web app
 *      Execute as: Me
 *      Who has access: Anyone with the link
 * 3. Скопируйте URL веб-приложения (заканчивается на /exec) и пришлите его в чат.
 * 4. Дальше просто присылайте в чат: "запусти пересборку" — и Claude вызовет
 *    этот URL с вашим токеном через web_fetch.
 */

var SECRET_TOKEN = 'ЗАМЕНИ_МЕНЯ_НА_СВОЙ_ПАРОЛЬ';

function doGet(e) {
  // Старый режим: вызов ?action=rebuild&token=... пересобирает вкладки (используется вручную из браузера)
  if (e && e.parameter && e.parameter.action === 'rebuild') {
    try {
      if (e.parameter.token !== SECRET_TOKEN) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Неверный токен' })
        ).setMimeType(ContentService.MimeType.JSON);
      }
      var result = rebuildRelationalTabs();
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'ok', result: result })
      ).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'error', message: String(err) })
      ).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // По умолчанию — показываем форму-генератор UTM-ссылок
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Генератор UTM-ссылок')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Данные для выпадающих списков в форме генератора.
 * Читает вкладку "Обозначения" (столбцы F-I: Аналитика/Аналитика utm/utm_source/utm_medium;
 * J-K: ЖК/Проект utm; L-M: РА/РА_utm).
 */
function getGeneratorData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Обозначения');
  if (!sh) throw new Error('Вкладка "Обозначения" не найдена — сначала запустите rebuildRelationalTabs');
  var data = sh.getDataRange().getValues();

  var tip = [], razdel = [], istochnik = [], analitika = [], proekt = [], ra = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (row[0] !== '' && row[0] !== null) tip.push(row[0]);                 // A Тип
    if (row[2] !== '' && row[2] !== null) razdel.push(row[2]);              // C Раздел
    if (row[3] !== '' && row[3] !== null) istochnik.push(row[3]);           // D Источник
    if (row[5] !== '' && row[5] !== null) {
      // F=Аналитика(5), G=Аналитика utm(6), H=utm_source(7), I=utm_medium(8)
      analitika.push({ label: row[5], utmSource: row[7], utmMedium: row[8] });
    }
    if (row[9] !== '' && row[9] !== null) {
      // J=ЖК.Источник(9), K=Проект utm(10)
      proekt.push({ label: row[9], code: row[10] });
    }
    if (row[11] !== '' && row[11] !== null) {
      // L=РА(11), M=РА_utm(12)
      ra.push({ label: row[11], code: row[12] });
    }
  }
  return { tip: tip, razdel: razdel, istochnik: istochnik, analitika: analitika, proekt: proekt, ra: ra };
}

/**
 * Список доменов, для которых имеет смысл обогащать ссылку utm-метками.
 */
var ALLOWED_DOMAINS = [
  'sever-group.ru',
  'double-double.ru',
  'literaturny.ru',
  'xn--80ajrlru9c.xn--p1ai',
  'dius-mfk.ru',
];

function extractHost(url) {
  var m = String(url).match(/^https?:\/\/([^\/]+)/i);
  if (!m) return null;
  return m[1].toLowerCase().replace(/^www\./, '');
}

function isAllowedDomain(url) {
  var host = extractHost(url);
  if (!host) return false;
  return ALLOWED_DOMAINS.some(function (d) {
    return host === d || host.indexOf('.' + d) === host.length - d.length - 1;
  });
}

/**
 * Строит финальный url_с_utm по той же логике, что формула в Sources.
 */
function buildUrlWithUtm(urlBase, utmSource, utmMedium, utmCampaign, utmTerm) {
  if (!urlBase) return '';
  var sep = urlBase.indexOf('?') === -1 ? '?' : '&';
  return urlBase + sep +
    'utm_source=' + utmSource +
    '&utm_medium=' + utmMedium +
    '&utm_campaign=' + utmCampaign +
    '&utm_term=' + utmTerm;
}

/**
 * Главный обработчик отправки формы: пишет новую строку в "Справочник источников",
 * забирает term из резерва, пересобирает Sources/Phones (через rebuildRelationalTabs),
 * при необходимости шлёт уведомление о нехватке резерва.
 */
function processSubmission(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Проверка домена — если не из разрешённого списка, дальше не идём
  if (!isAllowedDomain(payload.urlBase)) {
    return { status: 'domain_not_allowed', message: 'нет необходимости обогащать utm-метками' };
  }

  var shOboz = ss.getSheetByName('Обозначения');
  var ozData = shOboz.getDataRange().getValues();

  function findInOznCol(labelCol, codeCol, label) {
    for (var r = 1; r < ozData.length; r++) {
      if (ozData[r][labelCol] === label) return ozData[r][codeCol];
    }
    return null;
  }

  var utmSource = findInOznCol(5, 7, payload.analitika);   // F -> H
  var utmMedium = findInOznCol(5, 8, payload.analitika);   // F -> I
  var proektUtm = findInOznCol(9, 10, payload.proekt);     // J -> K
  var raUtm = findInOznCol(11, 12, payload.ra);             // L -> M

  if (utmSource === null || proektUtm === null || raUtm === null) {
    return { status: 'error', message: 'Не удалось найти код для одного из выбранных значений на вкладке Обозначения' };
  }

  // 2. Берём первый свободный резервный utm_term
  var shReserve = ss.getSheetByName('Резервные_UTM_term');
  if (!shReserve) return { status: 'error', message: 'Вкладка "Резервные_UTM_term" не найдена' };
  var reserveData = shReserve.getDataRange().getValues();
  var reservedTerm = null;
  for (var rr = 1; rr < reserveData.length; rr++) {
    if (reserveData[rr][0] && reserveData[rr][1] === 'Свободен') {
      reservedTerm = reserveData[rr][0];
      break;
    }
  }
  if (!reservedTerm) {
    return { status: 'error', message: 'Нет свободных резервных utm_term. Пополните вкладку "Резервные_UTM_term".' };
  }

  // 3. Название кампании — одно слово, без пробела в конце
  var campaignWord = String(payload.campaignWord || '').replace(/\s+$/, '').trim();
  if (!campaignWord) {
    return { status: 'error', message: 'Укажите название кампании' };
  }
  var now = new Date();
  var year = now.getFullYear();
  var month = ('0' + (now.getMonth() + 1)).slice(-2);
  var utmCampaign = raUtm + '|' + proektUtm + '|' + campaignWord + '_' + year + '_' + month;

  var urlWithUtm = buildUrlWithUtm(payload.urlBase, utmSource, utmMedium, utmCampaign, reservedTerm);
  var analitikaTail = String(payload.analitika).replace(/^\d+\.\d+\s*/, '');
  var calltouchPool = payload.istochnik + ' / ' + analitikaTail + ' / ' + payload.proekt;

  // 4. Пишем новую строку в первую вкладку ("Справочник источников") по названиям заголовков
  var src = ss.getSheets()[0];
  var header = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0];
  function colOf(name) {
    for (var i = 0; i < header.length; i++) {
      if (String(header[i]).trim().toLowerCase().indexOf(name.toLowerCase()) === 0) return i;
    }
    return -1;
  }
  var newRow = new Array(header.length).fill('');
  newRow[colOf('Тип')] = payload.tip;
  newRow[colOf('Категория')] = 'без телефона';
  newRow[colOf('Раздел')] = payload.razdel;
  newRow[colOf('Источник')] = payload.istochnik;
  newRow[colOf('utm_source')] = utmSource;
  newRow[colOf('utm_medium')] = utmMedium;
  newRow[colOf('Аналитика')] = payload.analitika;
  newRow[colOf('ЖК')] = payload.proekt;
  newRow[colOf('UTM_term')] = reservedTerm;
  newRow[colOf('URL')] = payload.urlBase;
  var campaignColIdx = colOf('UTM_Campaign');
  if (campaignColIdx !== -1) newRow[campaignColIdx] = utmCampaign;
  var urlUtmColIdx = colOf('URL с UTM');
  if (urlUtmColIdx !== -1) newRow[urlUtmColIdx] = urlWithUtm;
  var calltouchColIdx = colOf('Calltouch');
  if (calltouchColIdx !== -1) newRow[calltouchColIdx] = calltouchPool;

  src.getRange(src.getLastRow() + 1, 1, 1, newRow.length).setValues([newRow]);

  // 5. Пересобираем Sources / Phones (и заодно Обозначения / статус резерва)
  rebuildRelationalTabs();

  // 6. Проверяем остаток резерва и, если нужно, уведомляем
  var reserveNotice = checkReserveAndNotify();

  return {
    status: 'ok',
    urlWithUtm: urlWithUtm,
    utmTerm: reservedTerm,
    utmCampaign: utmCampaign,
    reserveNotice: reserveNotice,
  };
}



/**
 * ЗАГОТОВКА на будущее — пока НЕ подключена ни к какому триггеру.
 * Когда решим, как уведомлять (email / чат), допишем тело этой функции
 * и поставим Time-driven trigger, чтобы она проверяла запас регулярно.
 */
function checkReserveAndNotify() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Резервные_UTM_term');
  if (!sh) return { notified: false };
  var free = sh.getRange('E1').getValue();
  var threshold = sh.getRange('E2').getValue() || 5;
  if (free < threshold) {
    try {
      MailApp.sendEmail(
        Session.getEffectiveUser().getEmail(),
        'Мало резервных utm_term',
        'На вкладке "Резервные_UTM_term" осталось свободных значений: ' + free +
        ' (порог уведомления: ' + threshold + '). Пополните список.'
      );
    } catch (e) {
      // если отправка почты недоступна — не роняем основной процесс
    }
    return { notified: true, free: free, threshold: threshold };
  }
  return { notified: false, free: free, threshold: threshold };
}

function rebuildRelationalTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheets()[0]; // первая вкладка — исходные данные

  function uniqueValues(rows, colIdx) {
    var seen = {};
    var out = [];
    rows.forEach(function (r) {
      var v = r[colIdx];
      if (v !== '' && v !== null && v !== undefined && !seen[v]) {
        seen[v] = true;
        out.push(v);
      }
    });
    return out;
  }

  // В русской (и большинстве европейских) локалей разделитель аргументов в формулах — ";", а не ","
  var locale = ss.getSpreadsheetLocale() || 'en_US';
  var SEP = locale.toLowerCase().indexOf('en') === 0 ? ',' : ';';
  var data = src.getDataRange().getValues();
  var header = data[0];

  // Находим индексы нужных колонок по названию заголовка (устойчиво к перестановке столбцов)
  function col(name) {
    for (var i = 0; i < header.length; i++) {
      if (String(header[i]).trim().toLowerCase().indexOf(name.toLowerCase()) === 0) return i;
    }
    return -1;
  }
  var idx = {
    tip: col('Тип'),
    kat: col('Категория'),
    razdel: col('Раздел'),
    istochnik: col('Источник'),
    utmSource: col('utm_source'),
    utmMedium: col('utm_medium'),
    analitika: col('Аналитика'),
    zhk: col('ЖК'),
    telefon: col('Телефон'),
    komment: col('Комментарий'),
    bitrix: col('ID Битрикс'),
    utmTerm: col('UTM_term'),
    proekt: col('Проект в Колтач'),
    url: col('URL'),
    utmCampaign: col('UTM_Campaign'),
    calltouch: col('Calltouch'),
  };

  var rows = data.slice(1).filter(function (r) {
    return r[idx.utmTerm] !== '' && r[idx.utmTerm] !== null;
  });

  // ---------- SOURCES (дедуп по utm_term, первое вхождение) ----------
  var sourcesMap = {};
  var sourceOrder = [];
  rows.forEach(function (r) {
    var term = r[idx.utmTerm];
    if (!sourcesMap[term]) {
      sourcesMap[term] = r;
      sourceOrder.push(term);
    }
  });

  // ---------- PHONES (одна строка на непустой телефон) ----------
  var phones = [];
  rows.forEach(function (r) {
    if (r[idx.telefon]) phones.push([r[idx.utmTerm], r[idx.telefon]]);
  });

  function getOrCreateSheet(name) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      sh.clear();
    } else {
      sh = ss.insertSheet(name);
    }
    return sh;
  }

  // ---------- Записываем Sources ----------
  var shSources = getOrCreateSheet('Sources');
  var srcHeaders = [
    'utm_term (PK)', 'тип', 'категория', 'раздел', 'источник', 'utm_source',
    'utm_medium', 'аналитика', 'жк_источник', 'комментарий', 'id_bitrix24',
    'проект_в_колтач', 'utm_campaign', 'calltouch_pool', 'url_base', 'url_с_utm (формула)',
  ];
  var srcRows = sourceOrder.map(function (term) {
    var r = sourcesMap[term];
    return [
      term, r[idx.tip], r[idx.kat], r[idx.razdel], r[idx.istochnik], r[idx.utmSource],
      r[idx.utmMedium], r[idx.analitika], r[idx.zhk], r[idx.komment], r[idx.bitrix],
      r[idx.proekt], r[idx.utmCampaign], r[idx.calltouch], r[idx.url], '',
    ];
  });
  shSources.getRange(1, 1, 1, srcHeaders.length).setValues([srcHeaders]).setFontWeight('bold');
  if (srcRows.length) shSources.getRange(2, 1, srcRows.length, srcHeaders.length).setValues(srcRows);
  // Формула url_с_utm — считается по столбцам той же строки (utm_source=F, utm_medium=G,
  // utm_campaign=M, url_base=O, utm_term=A), без отдельной таблицы Urls
  sourceOrder.forEach(function (term, i) {
    var row = i + 2;
    shSources.getRange(row, 16).setFormula(
      '=IF(OR(O' + row + '=""' + SEP + 'O' + row + '="нет ссылки")' + SEP + '""' + SEP +
      'O' + row + '&IF(ISNUMBER(SEARCH("~?"' + SEP + 'O' + row + '))' + SEP + '"&"' + SEP + '"?")' +
      '&"utm_source="&F' + row +
      '&"&utm_medium="&G' + row +
      '&"&utm_campaign="&M' + row +
      '&"&utm_term="&A' + row + ')'
    );
  });
  shSources.setFrozenRows(1);

  // ---------- Записываем Phones ----------
  var shPhones = getOrCreateSheet('Phones');
  shPhones.getRange(1, 1, 1, 4)
    .setValues([['phone_id (PK)', 'utm_term (FK)', 'телефон', 'источник (проверка, из Sources)']])
    .setFontWeight('bold');
  phones.forEach(function (p, i) {
    var row = i + 2;
    shPhones.getRange(row, 1, 1, 3).setValues([[i + 1, p[0], p[1]]]);
    shPhones.getRange(row, 4).setFormula(
      '=IFERROR(INDEX(Sources!H:H' + SEP + 'MATCH(B' + row + SEP + 'Sources!A:A' + SEP + '0))' + SEP + '"")'
    );
  });
  shPhones.setFrozenRows(1);

  // ---------- Удаляем вкладку Urls, если осталась от прошлой версии ----------
  var oldUrls = ss.getSheetByName('Urls');
  if (oldUrls) ss.deleteSheet(oldUrls);

  // ---------- Записываем Резервные_UTM_term (создаётся один раз, дальше не затирается) ----------
  var reserveSheetName = 'Резервные_UTM_term';
  var shReserve = ss.getSheetByName(reserveSheetName);
  var reserveIsNew = false;
  if (!shReserve) {
    shReserve = ss.insertSheet(reserveSheetName);
    reserveIsNew = true;
  }
  shReserve.getRange(1, 1, 1, 2).setValues([['utm_term (резерв)', 'статус']]).setFontWeight('bold');
  if (reserveIsNew) {
    var initialReserve = ['726|', '727|', '740|', '741|', '743|'];
    shReserve.getRange(2, 1, initialReserve.length, 1).setValues(initialReserve.map(function (v) { return [v]; }));
  }
  var reserveLastRow = shReserve.getLastRow();
  if (reserveLastRow >= 2) {
    for (var rr = 2; rr <= reserveLastRow; rr++) {
      // Резерв считается "Использован", если этот utm_term уже появился в Sources
      // (т.е. кто-то завёл под ним реальный источник в "Справочник источников")
      shReserve.getRange(rr, 2).setFormula(
        '=IF(COUNTIF(Sources!A:A' + SEP + 'A' + rr + ')>0' + SEP + '"Использован"' + SEP + '"Свободен")'
      );
    }
  }
  var reserveLastRowForCount = Math.max(reserveLastRow, 2);
  shReserve.getRange(1, 4).setValue('Свободных резервов:').setFontWeight('bold');
  shReserve.getRange(1, 5).setFormula(
    '=COUNTIF(B2:B' + reserveLastRowForCount + SEP + '"Свободен")'
  );
  shReserve.setFrozenRows(1);
  // Порог для будущего уведомления ("настроим его позже") — держим отдельной именованной ячейкой,
  // чтобы не хардкодить число "5" внутри кода уведомления.
  shReserve.getRange(2, 4).setValue('Порог уведомления:');
  if (reserveIsNew) shReserve.getRange(2, 5).setValue(5);

  // ---------- Записываем Обозначения (уникальные значения + доп. столбец "Источник utm") ----------
  var shOboz = getOrCreateSheet('Обозначения');
  var istochnikVals = uniqueValues(rows, idx.istochnik);
  var analitikaVals = uniqueValues(rows, idx.analitika);
  var zhkVals = uniqueValues(rows, idx.zhk);
  // Соответствие "Источник" -> короткий код utm_source-префикса.
  // Добавляйте сюда новые пары, если появятся другие значения "Источник".
  var sourceUtmMap = {
    '01.02. Базы недвижимости': 'base',
    '02.05 Печатная продукция': 'press',
  };
  // Соответствие "Аналитика (новое название)" -> короткий utm-код.
  var analitikaUtmMap = {
    '01.02 Авито': 'avito',
    '01.02 Домклик': 'domclick',
    '02.05 Вся печать': 'all',
    '02.05 Каталог': 'catalog',
    '02.05 Лифлет': 'liflet',
  };
  // Соответствие "ЖК. Источник" -> короткий код проекта для utm_campaign.
  // Добавляйте сюда новые пары, если появятся другие ЖК/офисы.
  var zhkUtmMap = {
    'Айрис': 'irs',
    'Дабл-Дабл': 'dbl',
    'Диус': 'diu',
    'Диус. Коммерция': 'dcm',
    'Екатеринбург': 'ekb',
    'Литературный': 'ltr',
    'Москва': 'msk',
    'Север': 'svr',
    'Смородина': 'smd',
    'Тюмень': 'tmn',
    'Эхо леса': 'eho',
  };
  // Сопоставляем "Источник" и "Аналитика (новое название)" по числовому префиксу
  // (например, "01.02" в "01.02. Базы недвижимости" и в "01.02 Авито")
  function numPrefix(str) {
    var m = String(str).match(/^\d+\.\d+/);
    return m ? m[0] : '';
  }
  var istochnikByPrefix = {};
  istochnikVals.forEach(function (v) {
    istochnikByPrefix[numPrefix(v)] = v;
  });
  var utmSourceVals = analitikaVals.map(function (a) {
    var matchedIstochnik = istochnikByPrefix[numPrefix(a)];
    if (!matchedIstochnik) return '';
    var srcCode = sourceUtmMap[matchedIstochnik];
    var anCode = analitikaUtmMap[a];
    if (!srcCode || !anCode) return '';
    return srcCode + '_' + anCode;
  });
  // utm_medium — определяется по значению utm_source
  var utmSourceMediumMap = {
    'base_avito': 'cpa',
    'base_domclick': 'cpa',
    'press_all': 'print',
    'press_catalog': 'qr',
    'press_liflet': 'qr',
  };
  var utmMediumVals = utmSourceVals.map(function (s) { return utmSourceMediumMap[s] || ''; });
  // "РА" — фиксированный список (не из исходных данных, задаётся вручную)
  var raVals = ['Север', 'SA', 'E-promo'];
  var raUtmMap = {
    'Север': 'svr',
    'SA': 'sam',
    'E-promo': 'epr',
  };
  var columnsDef = [
    { label: header[idx.tip], values: uniqueValues(rows, idx.tip) },
    { label: header[idx.kat], values: uniqueValues(rows, idx.kat) },
    { label: header[idx.razdel], values: uniqueValues(rows, idx.razdel) },
    { label: header[idx.istochnik], values: istochnikVals },
    { label: 'Источник utm', values: istochnikVals.map(function (v) { return sourceUtmMap[v] || ''; }) },
    { label: header[idx.analitika], values: analitikaVals },
    { label: 'Аналитика utm', values: analitikaVals.map(function (v) { return analitikaUtmMap[v] || ''; }) },
    { label: 'utm_source', values: utmSourceVals },
    { label: 'utm_medium', values: utmMediumVals },
    { label: header[idx.zhk], values: zhkVals },
    { label: 'Проект utm', values: zhkVals.map(function (v) { return zhkUtmMap[v] || ''; }) },
    { label: 'РА', values: raVals },
    { label: 'РА_utm', values: raVals.map(function (v) { return raUtmMap[v] || ''; }) },
  ];
  var oznCounts = {};
  columnsDef.forEach(function (colDef, ci) {
    oznCounts[colDef.label] = colDef.values.length;
    shOboz.getRange(1, ci + 1).setValue(colDef.label).setFontWeight('bold');
    if (colDef.values.length) {
      var colArr = colDef.values.map(function (v) { return [v]; });
      shOboz.getRange(2, ci + 1, colDef.values.length, 1).setValues(colArr);
    }
  });
  shOboz.setFrozenRows(1);

  return {
    строк_в_источнике: rows.length,
    уникальных_utm_term: sourceOrder.length,
    записей_phones: phones.length,
    обозначения: oznCounts,
  };
}
