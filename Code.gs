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

/**
 * Возвращает вкладку "Справочник источников" по ИМЕНИ, а не по позиции.
 * ВАЖНО: раньше код брал ss.getSheets()[0] ("первая вкладка слева"), что незаметно
 * ломалось, если появлялась любая другая вкладка левее (например, дубль "... old"
 * для сравнения версий) — код тогда читал/писал не туда, без единой ошибки.
 * Эта функция ищет вкладку по названию и намеренно игнорирует любые вкладки,
 * в названии которых встречается "old" (регистр не важен), даже если структура
 * поменяется в будущем.
 */
function getSourceSheet(ss) {
  var byName = ss.getSheetByName('Справочник источников');
  if (byName) return byName;
  // Резервный сценарий (вкладку переименовали) — берём первую вкладку, которая не входит
  // в список наших служебных вкладок и не похожа на архивную копию ("... old").
  var KNOWN_SERVICE_SHEETS = ['Sources', 'Phones', 'Обозначения', 'Резервные_UTM_term'];
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (KNOWN_SERVICE_SHEETS.indexOf(name) !== -1) continue;
    if (name.toLowerCase().indexOf('old') !== -1) continue;
    return sheets[i];
  }
  throw new Error('Не удалось найти вкладку "Справочник источников" — проверьте название вкладки');
}

var SECRET_TOKEN = 'ЗАМЕНИ_МЕНЯ_НА_СВОЙ_ПАРОЛЬ';

/**
 * ДИАГНОСТИКА: запустите эту функцию из редактора Apps Script (выбрать в списке функций -> ▶️ Выполнить),
 * затем откройте "Просмотр -> Журналы выполнения" (View -> Logs / Executions), чтобы увидеть,
 * что именно возвращают normalizeUrl/isAllowedDomain в вашей текущей живой копии кода.
 */
function testDomainCheck() {
  var tests = ['sever-group.ru/', 'https://sever-group.ru/', 'sever-group.ru'];
  tests.forEach(function (t) {
    var norm = normalizeUrl(t);
    var host = extractHost(norm);
    var allowed = isAllowedDomain(norm);
    Logger.log('ВХОД: "%s" | normalizeUrl -> "%s" | host -> "%s" | ALLOWED_DOMAINS -> %s | isAllowedDomain -> %s',
      t, norm, host, JSON.stringify(ALLOWED_DOMAINS), allowed);
  });
}

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
      var mediumOptions = utmSourceMediumMap[row[7]] || (row[8] ? [row[8]] : []);
      analitika.push({ label: row[5], utmSource: row[7], utmMedium: row[8], mediumOptions: mediumOptions });
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
  var srcSheet = getSourceSheet(ss);
  var sourceSheetUrl = ss.getUrl() + '#gid=' + srcSheet.getSheetId();
  // Ответственный за utm-метку — фиксированный список (не из исходных данных, задаётся вручную)
  var otvetstvennyy = ['Копылова', 'Семенчук', 'Тишкова', 'Лучинкина', 'Дроздецкий'];
  return {
    tip: tip, razdel: razdel, istochnik: istochnik, analitika: analitika, proekt: proekt, ra: ra,
    otvetstvennyy: otvetstvennyy,
    sourceSheetUrl: sourceSheetUrl,
  };
}

/**
 * Ищет строки в "Справочник источников", подходящие ПОД ВСЕ заполненные фильтры сразу
 * (логика "И" — пустые поля в filters игнорируются). Работает напрямую через SpreadsheetApp
 * (без Advanced Google Sheets API, без Filter Views) — простой и надёжный способ показать
 * существующие записи прямо в форме, не открывая саму таблицу.
 *
 * filters: { tip, razdel, istochnik, analitika, proekt, ra } — любое подмножество полей,
 * пустая строка/отсутствие ключа = не фильтровать по этому полю.
 * Возвращает { count: число совпадений, rows: [ {tip, razdel, istochnik, analitika, zhk, utmTerm, url}, ... ] }
 * (rows ограничено первыми 30 совпадениями, чтобы не перегружать форму).
 */
function searchSourceRowsMulti(filters) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = getSourceSheet(ss);
  var data = src.getDataRange().getValues();
  var header = data[0];

  function colOf(name) {
    for (var i = 0; i < header.length; i++) {
      if (String(header[i]).trim().toLowerCase().indexOf(name.toLowerCase()) === 0) return i;
    }
    return -1;
  }

  var FIELD_COLUMN = {
    tip: 'Тип',
    razdel: 'Раздел',
    istochnik: 'Источник',
    analitika: 'Аналитика',
    proekt: 'ЖК',
  };

  var idxTip = colOf('Тип');
  var idxRazdel = colOf('Раздел');
  var idxIstochnik = colOf('Источник');
  var idxAnalitika = colOf('Аналитика');
  var idxZhk = colOf('ЖК');
  var idxUtmTerm = colOf('UTM_term');
  var idxUrlUtm = colOf('URL с UTM');
  var idxPhone = colOf('Телефон');
  if (idxPhone === -1) idxPhone = colOf('Дополнительно об источнике'); // столбец был переименован, см. DECISIONS.md
  var idxCampaign = colOf('UTM_Campaign');

  // Собираем список условий (каждое — функция row -> boolean), которые нужно выполнить ВСЕ.
  var checks = [];
  ['tip', 'razdel', 'istochnik', 'analitika', 'proekt'].forEach(function (key) {
    var value = filters[key];
    if (!value) return;
    var colIdx = colOf(FIELD_COLUMN[key]);
    if (colIdx === -1) return;
    checks.push(function (row) { return row[colIdx] === value; });
  });
  if (filters.ra) {
    // В "Справочник источников" нет отдельного столбца РА — код РА закодирован
    // первым сегментом в UTM_Campaign ("ра|проект|название"), поэтому проверяем "начинается с".
    var shOboz = ss.getSheetByName('Обозначения');
    if (!shOboz) throw new Error('Вкладка "Обозначения" не найдена');
    var ozData = shOboz.getDataRange().getValues();
    var raCode = null;
    for (var r = 1; r < ozData.length; r++) {
      if (ozData[r][11] === filters.ra) { raCode = ozData[r][12]; break; } // L=РА, M=РА_utm
    }
    if (!raCode) throw new Error('Код РА не найден для значения: ' + filters.ra);
    var prefix = raCode + '|';
    checks.push(function (row) { return String(row[idxCampaign] || '').indexOf(prefix) === 0; });
  }

  var matches = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var ok = checks.every(function (check) { return check(row); });
    if (ok) {
      matches.push({
        tip: row[idxTip],
        razdel: row[idxRazdel],
        istochnik: row[idxIstochnik],
        analitika: row[idxAnalitika],
        zhk: row[idxZhk],
        utmTerm: row[idxUtmTerm],
        urlUtm: row[idxUrlUtm],
        phone: row[idxPhone],
      });
      if (matches.length >= 30) break;
    }
  }

  return { count: matches.length, rows: matches };
}

// utm_medium — определяется по значению utm_source.
// Значение — МАССИВ вариантов (для некоторых utm_source в реальных данных встречается
// больше одного medium). Первый элемент — вариант по умолчанию (если у Аналитики
// только один вариант). Если вариантов несколько — форма предложит выбор пользователю.
var utmSourceMediumMap = {
  'base_avito': ['cpa'],
  'base_cian': ['cpa'],
  'base_domclick': ['cpa'],
  'base_domrf': ['cpa'],
  'base_yan': ['cpa'],
  'base_yansearch': ['cpa'],
  'building_fence': ['qr'],
  'building_navigation': ['call'],
  'building_passport': ['qr'],
  'building_site': ['call'],
  'corporate_presentation': ['qr'],
  'crm_sms': ['send'],
  'direct_yandex': ['cpc'],
  'event_festival': ['qr'],
  'event_presentation': ['qr', 'call'],
  'gift_clients': ['call'],
  'indoor_bc': ['video', 'banner'],
  'indoor_tc': ['video'],
  'indoor_transp': ['video'],
  'leeds_bigdata': ['cpa'],
  'leeds_flatoutlet': ['cpa'],
  'leeds_getflat': ['cpa'],
  'leeds_gidmarket': ['cpa'],
  'leeds_hitmedia': ['cpa'],
  'leeds_marketcall': ['cpa'],
  'leeds_novostroygid': ['cpa'],
  'leeds_otclick': ['cpa'],
  'leeds_realty': ['cpa'],
  'leeds_rocketad': ['cpa'],
  'leeds_rosttech': ['cpa'],
  'leeds_solutionspro': ['cpa'],
  'leeds_stroynov': ['cpa'],
  'leeds_toogeo': ['cpa'],
  'leeds_сallexchange': ['cpa'],
  'maps_2gis': ['cpa'],
  'maps_yandex': ['cpa'],
  'media_avito': ['cpm'],
  'media_mosdom': ['cpm'],
  'media_vk': ['cpv'],
  'media_yan': ['cpm'],
  'media_yandex': ['cpm'],
  'outdoor_dooh': ['banner'],
  'outdoor_ooh': ['banner'],
  'ownmedia_dzen': ['post'],
  'ownmedia_orm': ['post'],
  'ownmedia_tg': ['post'],
  'ownmedia_vk': ['post'],
  'pr_posev': ['post'],
  'pr_release': ['post'],
  'pr_smm': ['post'],
  'press_album': ['qr'],
  'press_all': ['print'],
  'press_brochure': ['qr'],
  'press_catalog': ['qr'],
  'press_envelope': ['qr'],
  'press_folder': ['qr'],
  'press_liflet': ['print', 'qr'],
  'press_uralairlains': ['qr'],
  'radio_broadcast': ['audio'],
  'salesoffice_lightbox': ['qr'],
  'salesoffice_panno': ['qr'],
  'salesoffice_presentation': ['call', 'qr'],
  'search_google': ['organic'],
  'search_yandex': ['organic'],
  'self_nocall': ['visit'],
  'site_callback': ['call'],
  'site_commercial': ['call'],
  'site_contacts': ['call'],
  'site_header': ['call'],
  'site_phone': ['call'],
  'site_printedsheet': ['print'],
  'target_vk': ['cpc'],
  'theme_danteresearch': ['cpm'],
  'theme_mirkvartir': ['cpm'],
  'theme_move': ['cpm'],
  'theme_poisknovostroyki': ['cpm'],
  'theme_realty': ['cpm'],
  'theme_restate': ['cpm'],
  'theme_russianrealty': ['cpm'],
  'tv_broadcast': ['video'],
  'tv_pr': ['video'],
};

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

/**
 * Кириллические/человекочитаемые алиасы доменов -> их канонический punycode-адрес.
 * Добавляйте сюда новые пары, если появятся другие кириллические варианты написания.
 */
var CYRILLIC_DOMAIN_ALIASES = {
  'эхолеса.рф': 'xn--80ajrlru9c.xn--p1ai',
};

/**
 * Защита от formula injection: Google Sheets трактует значение, записанное через setValues,
 * как живую формулу, если оно начинается на =, +, -, @ (или на управляющий символ табуляции/
 * возврата каретки перед одним из них). Если пользовательский текст начинается с одного из этих
 * символов — добавляем ведущий апостроф, чтобы ячейка осталась текстом, а не выполнялась как формула.
 */
function sanitizeCell(value) {
  var s = String(value === null || value === undefined ? '' : value);
  if (/^[=+\-@\t\r]/.test(s)) return "'" + s;
  return s;
}

function extractHost(url) {
  var m = String(url).match(/^https?:\/\/([^\/?#]+)/i);
  if (!m) return null;
  var authority = m[1];
  // Отсекаем userinfo (user:pass@host) — берём то, что после последнего "@", если оно есть,
  // чтобы "evil.com@sever-group.ru" не давал ложного совпадения на реальный домен ни в какую сторону.
  var atIdx = authority.lastIndexOf('@');
  if (atIdx !== -1) authority = authority.slice(atIdx + 1);
  // Отсекаем порт (":8080")
  authority = authority.split(':')[0];
  return authority.toLowerCase().replace(/^www\./, '');
}

/**
 * Приводит введённый пользователем адрес к полноценному URL:
 * - добавляет "https://", если протокол не указан
 * - заменяет кириллический алиас домена на его punycode-эквивалент
 * - добавляет "/" в конец, если указан только домен без пути
 *
 * Все следующие варианты ввода признаются равнозначными и проходят как ALLOWED
 * (проверено тестом при разработке — см. DECISIONS.md, пункт 15):
 *   https://sever-group.ru/            sever-group.ru/            sever-group.ru
 *   https://double-double.ru/          double-double.ru/          double-double.ru
 *   https://literaturny.ru/            literaturny.ru/            literaturny.ru
 *   https://xn--80ajrlru9c.xn--p1ai/   xn--80ajrlru9c.xn--p1ai/   xn--80ajrlru9c.xn--p1ai
 *   https://dius-mfk.ru/               dius-mfk.ru/               dius-mfk.ru
 *   эхолеса.рф  ==  https://xn--80ajrlru9c.xn--p1ai  ==  https://xn--80ajrlru9c.xn--p1ai/
 */
function normalizeUrl(input) {
  var url = String(input || '').trim();
  if (!url) return '';
  if (!/^https?:\/\//i.test(url)) {
    url = 'https://' + url;
  }
  var host = extractHost(url);
  if (host && CYRILLIC_DOMAIN_ALIASES[host]) {
    url = url.replace(host, CYRILLIC_DOMAIN_ALIASES[host]);
  }
  // Если указан только домен без пути — добавляем "/" (единый вид домена-корня)
  if (/^https?:\/\/[^\/]+$/i.test(url)) {
    url = url + '/';
  }
  return url;
}

function isAllowedDomain(url) {
  var host = extractHost(url);
  if (!host) return false;
  return ALLOWED_DOMAINS.some(function (d) {
    if (host === d) return true;
    var suffix = '.' + d;
    return host.length > suffix.length && host.slice(host.length - suffix.length) === suffix;
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

  var kategoria = payload.kategoria === 'Звонки' ? 'Звонки' : 'без телефона';
  var isCallMode = kategoria === 'Звонки';

  // Телефон и "Проект в Колтач" обязательны в режиме "Звонки", иначе не используются вовсе.
  var phone = String(payload.phone || '').trim();
  var proektKoltach = String(payload.proektKoltach || '').trim();
  if (isCallMode) {
    if (!phone) {
      return { status: 'error', message: 'Укажите номер телефона' };
    }
    if (/\s/.test(phone)) {
      return { status: 'error', message: 'Номер телефона не должен содержать пробелов' };
    }
    if (!proektKoltach) {
      return { status: 'error', message: 'Укажите "Проект в Колтач"' };
    }
  }

  // 0a. Проверка на пробелы внутри URL — частая ошибка при копировании ссылки.
  // В режиме "Звонки" URL необязателен — если не указан, просто пропускаем эти шаги.
  var rawUrl = String(payload.urlBase || '').trim();
  var hasUrl = rawUrl !== '';
  if (!isCallMode && !hasUrl) {
    return { status: 'error', message: 'Укажите URL' };
  }
  if (hasUrl && /\s/.test(rawUrl)) {
    return {
      status: 'error',
      message: 'URL не должен содержать пробелов. Проверьте ссылку — возможно, при копировании попал лишний пробел.',
    };
  }

  var urlBase = '';
  if (hasUrl) {
    // 0b. Приводим URL к полному виду: добавляем https://, если не указан протокол,
    // и заменяем кириллический алиас домена на канонический punycode-адрес
    urlBase = normalizeUrl(rawUrl);

    // 1. Проверка домена — если не из разрешённого списка, дальше не идём
    if (!isAllowedDomain(urlBase)) {
      return { status: 'domain_not_allowed', message: 'нет необходимости обогащать utm-метками' };
    }
  }
  payload.urlBase = urlBase;

  var shOboz = ss.getSheetByName('Обозначения');
  var ozData = shOboz.getDataRange().getValues();

  function findInOznCol(labelCol, codeCol, label) {
    for (var r = 1; r < ozData.length; r++) {
      if (ozData[r][labelCol] === label) return ozData[r][codeCol];
    }
    return null;
  }

  var utmSource = findInOznCol(5, 7, payload.analitika);   // F -> H
  var proektUtm = findInOznCol(9, 10, payload.proekt);     // J -> K
  var raUtm = findInOznCol(11, 12, payload.ra);             // L -> M

  if (utmSource === null || proektUtm === null || raUtm === null) {
    return { status: 'error', message: 'Не удалось найти код для одного из выбранных значений на вкладке Обозначения' };
  }

  // utm_medium — если у этого utm_source несколько вариантов, берём выбор пользователя
  // (payload.utmMedium), иначе используем единственный/первый вариант по умолчанию.
  var mediumOptions = utmSourceMediumMap[utmSource] || [];
  var utmMedium;
  if (mediumOptions.length > 1) {
    if (!payload.utmMedium || mediumOptions.indexOf(payload.utmMedium) === -1) {
      return {
        status: 'error',
        message: 'Для этой площадки нужно выбрать тип размещения (medium): ' + mediumOptions.join(', '),
      };
    }
    utmMedium = payload.utmMedium;
  } else {
    utmMedium = mediumOptions[0] || '';
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

  // 3. Название кампании — одно слово латиницей, без пробелов и разделителей внутри.
  // В режиме "Звонки" необязательно: если не указано, UTM_Campaign просто не считаем.
  var campaignWord = String(payload.campaignWord || '').replace(/\s+$/, '').trim();
  if (!campaignWord && !isCallMode) {
    return { status: 'error', message: 'Укажите название кампании' };
  }
  if (campaignWord && !/^[A-Za-z0-9]+$/.test(campaignWord)) {
    return {
      status: 'error',
      message: 'Название кампании должно быть одним словом на латинице: только буквы A-Z и цифры, без пробелов, кириллицы, дефисов и подчёркиваний',
    };
  }
  var utmCampaign = '';
  if (campaignWord) {
    var now = new Date();
    var year = now.getFullYear();
    var month = ('0' + (now.getMonth() + 1)).slice(-2);
    var day = ('0' + now.getDate()).slice(-2);
    utmCampaign = raUtm + '|' + proektUtm + '|' + campaignWord + '_' + year + '_' + month + '_' + day;
  }

  // URL с UTM считаем только если реально есть URL — иначе поле остаётся пустым
  // (в исходных данных для звонков без ссылки принято писать "нет ссылки" в URL).
  var urlWithUtm = hasUrl ? buildUrlWithUtm(urlBase, utmSource, utmMedium, utmCampaign, reservedTerm) : '';
  var analitikaTail = String(payload.analitika).replace(/^\d+\.\d+\s*/, '');
  var calltouchPool = payload.istochnik + ' / ' + analitikaTail + ' / ' + payload.proekt;

  // 4. Пишем новую строку в первую вкладку ("Справочник источников") по названиям заголовков
  var src = getSourceSheet(ss);
  var header = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0];
  function colOf(name) {
    for (var i = 0; i < header.length; i++) {
      if (String(header[i]).trim().toLowerCase().indexOf(name.toLowerCase()) === 0) return i;
    }
    return -1;
  }

  // Гарантируем наличие столбца "Ответственный за utm-метку" сразу после "URL с UTM" — создаём при
  // первом запуске, если его ещё нет (физическая вставка колонки в нужном месте листа,
  // а не в конец). Это ДРУГОЕ поле, не путать с "Ответственный за площадку (РА)" в форме.
  var otvColIdx = colOf('Ответственный за utm-метку');
  if (otvColIdx === -1) {
    var urlUtmColIdx = colOf('URL с UTM');
    if (urlUtmColIdx !== -1) {
      src.insertColumnAfter(urlUtmColIdx + 1); // insertColumnAfter — позиция 1-индексная
      otvColIdx = urlUtmColIdx + 1;
      src.getRange(1, otvColIdx + 1).setValue('Ответственный за utm-метку');
      header.splice(otvColIdx, 0, 'Ответственный за utm-метку');
    } else {
      // "URL с UTM" не нашли (переименовали?) — добавляем в конец, чтобы не потерять поле
      otvColIdx = header.length;
      src.getRange(1, otvColIdx + 1).setValue('Ответственный за utm-метку');
      header.push('Ответственный за utm-метку');
    }
  }

  // Гарантируем наличие столбца "Статус" — создаём при первом запуске, если его ещё нет.
  // Новые записи из формы всегда попадают со статусом "На модерации": видны в самой
  // таблице сразу, но не участвуют в Sources/Phones, пока модератор не поменяет статус
  // на "Одобрено" (и заодно не впишет присвоенный utm_term в Колтач — см. DECISIONS.md).
  var statusColIdx = colOf('Статус');
  if (statusColIdx === -1) {
    statusColIdx = header.length;
    src.getRange(1, statusColIdx + 1).setValue('Статус');
    header.push('Статус');
  }

  var newRow = new Array(header.length).fill('');
  newRow[otvColIdx] = sanitizeCell(payload.otvetstvenny);
  newRow[colOf('Тип')] = sanitizeCell(payload.tip);
  newRow[colOf('Категория')] = kategoria;
  newRow[colOf('Раздел')] = sanitizeCell(payload.razdel);
  newRow[colOf('Источник')] = sanitizeCell(payload.istochnik);
  newRow[colOf('utm_source')] = utmSource;
  newRow[colOf('utm_medium')] = utmMedium;
  newRow[colOf('Аналитика')] = sanitizeCell(payload.analitika);
  newRow[colOf('ЖК')] = sanitizeCell(payload.proekt);
  newRow[colOf('UTM_term')] = reservedTerm;
  newRow[colOf('URL')] = hasUrl ? urlBase : 'нет ссылки';
  var telefonColIdx = colOf('Телефон');
  if (telefonColIdx === -1) telefonColIdx = colOf('Дополнительно об источнике');
  if (telefonColIdx !== -1 && phone) newRow[telefonColIdx] = sanitizeCell(phone);
  // Для звонков "Проект в Колтач" заполняется вручную (поле формы), а не выводится
  // автоматически из ЖК — соответствие не всегда 1:1 (см. "Не заведен" для некоторых ЖК).
  if (isCallMode && proektKoltach) {
    var proektKoltachIdx = colOf('Проект в Колтач');
    if (proektKoltachIdx !== -1) newRow[proektKoltachIdx] = sanitizeCell(proektKoltach);
  }
  newRow[statusColIdx] = 'На модерации';
  var campaignColIdx = colOf('UTM_Campaign');
  if (campaignColIdx !== -1) newRow[campaignColIdx] = utmCampaign;
  var urlUtmColIdx = colOf('URL с UTM');
  if (urlUtmColIdx !== -1) newRow[urlUtmColIdx] = urlWithUtm || (hasUrl ? '' : 'нет ссылки');
  var calltouchColIdx = colOf('Calltouch');
  if (calltouchColIdx !== -1) newRow[calltouchColIdx] = sanitizeCell(calltouchPool);

  // Не используем src.getLastRow() напрямую — он считает последней строку с ЛЮБЫМ
  // содержимым/форматированием (в т.ч. "пустой хвост", оставшийся после очистки старых
  // данных), а не последнюю реально заполненную. Ищем последнюю строку с непустым
  // UTM_term и пишем сразу после неё, чтобы не оставлять пустые строки-разрывы.
  var utmTermColIdx = colOf('UTM_term');
  var allValues = src.getRange(1, 1, src.getLastRow(), header.length).getValues();
  var lastDataRow = 1; // 1 = только заголовок, если данных вообще нет
  for (var lr = allValues.length - 1; lr >= 1; lr--) {
    if (allValues[lr][utmTermColIdx] !== '' && allValues[lr][utmTermColIdx] !== null) {
      lastDataRow = lr + 1; // +1: переводим 0-индекс массива в номер строки листа
      break;
    }
  }

  src.getRange(lastDataRow + 1, 1, 1, newRow.length).setValues([newRow]);

  // 5. Пересобираем Sources / Phones (и заодно Обозначения / статус резерва)
  rebuildRelationalTabs();

  // 6. Проверяем остаток резерва и, если нужно, уведомляем
  var reserveNotice = checkReserveAndNotify();

  return {
    status: 'ok',
    kategoria: kategoria,
    phone: phone,
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
  var src = getSourceSheet(ss); // по имени "Справочник источников", не по позиции — см. getSourceSheet()

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
  // То же самое, но пробует несколько вариантов названия по очереди (на случай, если
  // столбец переименовали в исходной таблице) — возвращает первый найденный.
  function colAny(names) {
    for (var i = 0; i < names.length; i++) {
      var found = col(names[i]);
      if (found !== -1) return found;
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
    // Столбец с телефонами в какой-то момент был переименован из "Телефон"
    // в "Дополнительно об источнике" — ищем по любому из двух названий.
    telefon: colAny(['Телефон', 'Дополнительно об источнике']),
    komment: col('Комментарий'),
    bitrix: col('ID Битрикс'),
    utmTerm: col('UTM_term'),
    proekt: col('Проект в Колтач'),
    url: col('URL'),
    utmCampaign: col('UTM_Campaign'),
    calltouch: col('Calltouch'),
    status: col('Статус'), // может отсутствовать (-1), если ни одной формы ещё не отправляли — тогда фильтрация ниже просто не применяется
  };

  // rows — ВСЕ строки с непустым utm_term, включая "На модерации"/"Отклонено".
  // Нужны в первую очередь для проверки резерва: term должен считаться занятым
  // сразу при создании черновика, а не только после одобрения (иначе его сможет
  // забрать кто-то другой, пока черновик ждёт модерации).
  var rows = data.slice(1).filter(function (r) {
    return r[idx.utmTerm] !== '' && r[idx.utmTerm] !== null;
  });

  var PENDING_STATUSES = ['На модерации', 'Отклонено'];
  // visibleRows — только "боевые" записи (одобренные или вообще без статуса — то есть
  // все исторические строки, заведённые до появления модерации). Именно из них строятся
  // Sources/Phones/Обозначения, чтобы черновики на модерации не попадали в отчётность.
  var visibleRows = idx.status === -1 ? rows : rows.filter(function (r) {
    return PENDING_STATUSES.indexOf(r[idx.status]) === -1;
  });

  // Выпадающий список для модератора на столбце "Статус" (если столбец уже существует).
  if (idx.status !== -1) {
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['На модерации', 'Одобрено. Добавлено в Колтач', 'Отклонено'], true)
      .setAllowInvalid(true)
      .build();
    src.getRange(2, idx.status + 1, Math.max(src.getMaxRows() - 1, 1), 1).setDataValidation(statusRule);
  }

  // ---------- SOURCES (дедуп по utm_term, первое вхождение; только "боевые" записи) ----------
  var sourcesMap = {};
  var sourceOrder = [];
  visibleRows.forEach(function (r) {
    var term = r[idx.utmTerm];
    if (!sourcesMap[term]) {
      sourcesMap[term] = r;
      sourceOrder.push(term);
    }
  });

  // ---------- PHONES (одна строка на непустой телефон; только "боевые" записи) ----------
  var phones = [];
  visibleRows.forEach(function (r) {
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
    // Множество utm_term, которые СЧИТАЮТСЯ занятыми: все строки, кроме "Отклонено" —
    // отклонённая заявка возвращает свой term в резерв как свободный, его можно выдать
    // снова. "На модерации" по-прежнему занимает номер (заявка ещё не решена).
    // Вычисляем в скрипте (не формулой), чтобы не зависеть ни от локали таблицы,
    // ни от фильтрации Sources.
    var allTermsSet = {};
    rows.forEach(function (r) {
      if (idx.status !== -1 && r[idx.status] === 'Отклонено') return;
      allTermsSet[r[idx.utmTerm]] = true;
    });
    var reserveTerms = shReserve.getRange(2, 1, reserveLastRow - 1, 1).getValues();
    var statusValues = reserveTerms.map(function (row) {
      return [allTermsSet[row[0]] ? 'Использован' : 'Свободен'];
    });
    shReserve.getRange(2, 2, statusValues.length, 1).setValues(statusValues);
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
    '01.01. Контекстная реклама': 'direct',
    '01.02. Базы недвижимости': 'base',
    '01.03 Медийная реклама': 'media',
    '01.04 Лидогенерация': 'leeds',
    '01.05. Геосервисы': 'maps',
    '01.06 Реклама в соцсетях': 'target',
    '01.07 Органика': 'search',
    '01.08 CRM маркетинг': 'crm',
    '01.09 Собственные медиа': 'ownmedia',
    '01.10 PR': 'pr',
    '01.11 Звонок с сайта': 'site',
    '01.12 Тематические площадки': 'theme',
    '01.13 Корпоративные программы': 'corporate',
    '02.01 Наружная реклама': 'outdoor',
    '02.02 Indoor': 'indoor',
    '02.03 ТВ': 'tv',
    '02.04 Радио': 'radio',
    '02.05 Печатная продукция': 'press',
    '02.06 Мероприятия': 'event',
    '02.07 Самоход': 'self',
    '02.08 Офис продаж': 'salesoffice',
    '02.09 Строительная площадка': 'building',
    '02.11 Подарки': 'gift',
  };
  // Соответствие "Аналитика (новое название)" -> короткий utm-код.
  // ВСЕ 98 значений подтверждены реальными данными из полного CSV-экспорта таблицы
  // (правило: код = часть utm_source после первого "_").
  var analitikaUtmMap = {
    '01.01 Яндекс': 'yandex',
    '01.02 Авито': 'avito',
    '01.02 Домклик': 'domclick',
    '01.02 Наш.дом.рф': 'domrf',
    '01.02 Циан': 'cian',
    '01.02 Яндекс Недвижимость': 'yan',
    '01.02 Яндекс Поиск квартиры': 'yansearch',
    '01.03 Авито': 'avito',
    '01.03 Мосдом': 'mosdom',
    '01.03 Яндекс': 'yandex',
    '01.03 Яндекс Недвижимость': 'yan',
    '01.03 VK Видеореклама': 'vk',
    '01.04 Сallexchange': 'сallexchange',
    '01.04 BigData': 'bigdata',
    '01.04 FlatOutlet': 'flatoutlet',
    '01.04 Getflat': 'getflat',
    '01.04 GidMarket': 'gidmarket',
    '01.04 Hitmedia': 'hitmedia',
    '01.04 MarketCall': 'marketcall',
    '01.04 Novostroy gid': 'novostroygid',
    '01.04 Otclick': 'otclick',
    '01.04 Realty': 'realty',
    '01.04 RocketAd': 'rocketad',
    '01.04 Rosttech': 'rosttech',
    '01.04 Solutions Pro': 'solutionspro',
    '01.04 Stroynov': 'stroynov',
    '01.04 Toogeo': 'toogeo',
    '01.05 2Gis': '2gis',
    '01.05 Яндекс карты': 'yandex',
    '01.06 VK Лидформы': 'vk',
    '01.07 Яндекс. Сайт эхолеса.рф': 'yandex',
    '01.07 Яндекс. Сайт dius-mfk.ru': 'yandex',
    '01.07 Яндекс. Сайт kvartal1702.ru': 'yandex',
    '01.07 Яндекс. Сайт literaturny.ru': 'yandex',
    '01.07 Яндекс. Сайт sever-group.ru': 'yandex',
    '01.07 Google. Сайт эхолеса.рф': 'google',
    '01.07 Google. Сайт dius-mfk.ru': 'google',
    '01.07 Google. Сайт kvartal1702.ru': 'google',
    '01.07 Google. Сайт literaturny.ru': 'google',
    '01.07 Google. Сайт sever-group.ru': 'google',
    '01.08 SMS рассылки': 'sms',
    '01.09 Яндекс Дзен': 'dzen',
    '01.09 ORM': 'orm',
    '01.09 Telegramm': 'tg',
    '01.09 VK': 'vk',
    '01.10 Посевы': 'posev',
    '01.10 Пресс-релизы': 'release',
    '01.10 SMM': 'smm',
    '01.11 Печать планировок с сайта': 'printedsheet',
    '01.11 Сайт эхолеса.рф': 'phone',
    '01.11 Сайт эхолеса.рф. Обратный звонок': 'callback',
    '01.11 Сайт dius-mfk.ru': 'phone',
    '01.11 Сайт dius-mfk.ru. Коммерция': 'commercial',
    '01.11 Сайт dius-mfk.ru. Обратный звонок': 'callback',
    '01.11 Сайт double-double.ru. Контакты': 'contacts',
    '01.11 Сайт double-double.ru': 'phone',
    '01.11 Сайт enso': 'phone',
    '01.11 Сайт kvartal1702.ru': 'phone',
    '01.11 Сайт literaturny.ru': 'phone',
    '01.11 Сайт literaturny.ru. Обратный звонок': 'callback',
    '01.11 Сайт sever-group.ru': 'phone',
    '01.11 Сайт sever-group.ru. Контакты': 'contacts',
    '01.11 Сайт sever-group.ru. Обратный звонок': 'callback',
    '01.11 Сайт sever-group.ru. Шапка': 'header',
    '01.12 Realty': 'realty',
    '01.12 Restate': 'restate',
    '01.12 Dante research': 'danteresearch',
    '01.12 Mirvartir': 'mirkvartir',
    '01.12 Move': 'move',
    '01.12 Poisknovostroyki': 'poisknovostroyki',
    '01.12 Russianrealty': 'russianrealty',
    '01.13 Mailing': 'presentation',
    '02.01 DOOH': 'dooh',
    '02.01 OOH': 'ooh',
    '02.02 Бизнес-центры': 'bc',
    '02.02 Торговые центры': 'tc',
    '02.02 Транспорт': 'transp',
    '02.03 Рекламные ролики': 'broadcast',
    '02.03 PR ролики': 'pr',
    '02.04 Рекламные ролики': 'broadcast',
    '02.05 Альбом планировочных решений': 'album',
    '02.05 Брошюра проекта': 'brochure',
    '02.05 Вся печать': 'all',
    '02.05 Каталог': 'catalog',
    '02.05 Конверт': 'envelope',
    '02.05 Лифлет': 'liflet',
    '02.05 Папка для клиентов': 'folder',
    '02.05 Уральские авиалинии': 'uralairlains',
    '02.06 Партнерство и спонсорство': 'festival',
    '02.06 Презентация': 'presentation',
    '02.07 Без звонка': 'nocall',
    '02.08 Лайтбокс': 'lightbox',
    '02.08 Панно': 'panno',
    '02.08 Презентация': 'presentation',
    '02.09 Забор': 'fence',
    '02.09 Навигация': 'navigation',
    '02.09 Паспорт объекта': 'passport',
    '02.09 Строительная площадка': 'site',
    '02.11 Подарки клиентам': 'clients',
  };
  // Соответствие "ЖК. Источник" -> короткий код проекта для utm_campaign.
  // Добавляйте сюда новые пары, если появятся другие ЖК/офисы.
  var zhkUtmMap = {
    'Айрис': 'irs',
    'Дабл-Дабл': 'dbl',
    'Мед': 'med',
    'Хьюгард': 'hgd',
    'Диус': 'diu',
    'Диус. Коммерция': 'dcm',
    'Екатеринбург': 'ekb',
    'Литературный': 'ltr',
    'Москва': 'msk',
    'Север': 'svr',
    'Смородина': 'smd',
    'Тюмень': 'tmn',
    'Энвилль': 'env',
    'Эхо леса': 'eho',
    'ЭНСО': 'ens',
    'Квартал 1702': 'skr',
    'Эхо леса. Вершины': 'lt2',
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
  // utmSourceMediumMap теперь глобальная константа (см. верх файла) —
  // это нужно, чтобы getGeneratorData() тоже мог её использовать для формы.
  // Для "Обозначения" показываем ВСЕ варианты medium через " / " (если их несколько) —
  // чтобы было видно прямо в таблице, что у площадки есть выбор. Форма при создании новой
  // записи использует тот же массив вариантов (см. processSubmission/getGeneratorData).
  var utmMediumVals = utmSourceVals.map(function (s) {
    var options = utmSourceMediumMap[s];
    return options && options.length ? options.join(' / ') : '';
  });
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

/**
 * Сравнивает вкладку "Справочник источников" (текущая, первая по счёту) с вкладкой
 * "Справочник источников old" построчно по UTM_term (на один term может быть несколько
 * строк — разные телефоны/URL/площадки) и подсвечивает жёлтым те строки на "Справочник
 * источников", у которых нет ТОЧНО такой же строки (по всем столбцам) в старой версии
 * с тем же term. Запускать вручную из редактора Apps Script, когда нужно сверить версии.
 *
 * ВНИМАНИЕ: строки, которые были в old, но полностью пропали в новой версии, подсветить
 * невозможно — их физически нет на вкладке "Справочник источников". Такие случаи функция
 * перечисляет в возвращаемом объекте (removedRows), но не подсвечивает (негде).
 */
function highlightDifferencesFromOld() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var newSheet = getSourceSheet(ss);
  var oldSheet = ss.getSheetByName('Справочник источников old');
  if (!oldSheet) throw new Error('Вкладка "Справочник источников old" не найдена');

  var newData = newSheet.getDataRange().getValues();
  var oldData = oldSheet.getDataRange().getValues();
  var header = newData[0];
  var ncols = header.length;

  function colOf(name) {
    for (var i = 0; i < header.length; i++) {
      if (String(header[i]).trim().toLowerCase().indexOf(name.toLowerCase()) === 0) return i;
    }
    return -1;
  }
  var utmIdx = colOf('UTM_term');

  function groupByTerm(data) {
    var map = {};
    for (var i = 1; i < data.length; i++) {
      var term = String(data[i][utmIdx] || '').trim();
      if (!term) continue;
      if (!map[term]) map[term] = [];
      map[term].push({ rowIndex: i, values: data[i] });
    }
    return map;
  }

  function rowSignature(values) {
    return values.map(function (v) { return String(v === null || v === undefined ? '' : v); }).join('\u0001');
  }

  var oldByTerm = groupByTerm(oldData);
  var newByTerm = groupByTerm(newData);

  var oldSigsByTerm = {};
  Object.keys(oldByTerm).forEach(function (term) {
    oldSigsByTerm[term] = oldByTerm[term].map(function (e) { return rowSignature(e.values); });
  });

  // Новые/изменённые строки — есть на новой вкладке, но такой же строки с тем же term нет в old
  var rowsToHighlight = [];
  Object.keys(newByTerm).forEach(function (term) {
    var oldSigs = oldSigsByTerm[term] || [];
    newByTerm[term].forEach(function (entry) {
      var sig = rowSignature(entry.values);
      if (oldSigs.indexOf(sig) === -1) {
        rowsToHighlight.push(entry.rowIndex + 1); // 1-индексный номер строки листа
      }
    });
  });

  // Строки, которые были в old, но их term вообще пропал из new — физически негде подсветить
  var removedTerms = Object.keys(oldByTerm).filter(function (term) { return !newByTerm[term]; });

  // Сбрасываем прошлую подсветку (только фон, оставленный этой же функцией) и подсвечиваем заново
  var lastRow = newSheet.getLastRow();
  if (lastRow > 1) newSheet.getRange(2, 1, lastRow - 1, ncols).setBackground(null);
  rowsToHighlight.forEach(function (rowNum) {
    newSheet.getRange(rowNum, 1, 1, ncols).setBackground('#fff2ac');
  });

  return { highlightedRows: rowsToHighlight, removedTerms: removedTerms };
}
