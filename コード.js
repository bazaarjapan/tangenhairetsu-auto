const CONFIG = {
  MODEL: 'gemini-flash-latest',
  TABLE_SHEET_NAME: '単元計画一覧',
  RAW_SHEET_NAME: '抽出データ',
  CACHE_SHEET_NAME: '処理キャッシュ',
  OCR_LANGUAGE: 'ja',
  MONTHS: ['4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月']
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('単元配列表')
    .addItem('Drive PDFから生成', 'buildUnitTableFromDrivePdfs')
    .addItem('処理キャッシュをクリア', 'clearProcessingCache')
    .addToUi();
}

function buildUnitTableFromDrivePdfs() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('Script Properties に GEMINI_API_KEY を設定してください。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('アクティブなスプレッドシートが見つかりません。');

  const folderId = getFolderIdFromA1_(ss);
  const files = listPdfFilesRecursive_(folderId);
  if (files.length === 0) {
    throw new Error('指定フォルダ配下にPDFファイルが見つかりません。');
  }

  const cache = loadCacheMap_(ss);
  const extracted = [];

  files.forEach(file => {
    try {
      const cachedPayload = getCachedPayloadIfFresh_(cache, file);
      if (cachedPayload) {
        extracted.push(cachedPayload);
        return;
      }

      const parsed = extractUnitsFromPdfWithFallback_(file, apiKey);
      if (!parsed || !Array.isArray(parsed.items)) return;
      extracted.push(parsed);

      cache[file.id] = {
        fileId: file.id,
        fileName: file.name,
        updatedMs: String(file.updatedMs),
        subject: parsed.subject || '',
        payloadJson: JSON.stringify(parsed),
        processedAt: new Date().toISOString()
      };
      Utilities.sleep(600);
    } catch (err) {
      Logger.log(`PDF処理失敗: ${file.name} (${file.id}) ${err}`);
    }
  });

  writeCacheMap_(ss, cache);

  const rowsByPdf = buildPdfRows_(extracted);
  renderMainTable_(ss, rowsByPdf);
  renderSubjectSheets_(ss, rowsByPdf);
  renderRawSheet_(ss, extracted);
}

function getFolderIdFromA1_(ss) {
  const baseSheet = ss.getActiveSheet() || ss.getSheets()[0];
  if (!baseSheet) throw new Error('シートが存在しません。');

  const a1 = String(baseSheet.getRange('A1').getDisplayValue()).trim();
  if (!a1) {
    throw new Error('A1セルにDriveフォルダURLまたはフォルダIDを入力してください。');
  }

  const id = extractDriveFolderId_(a1);
  if (!id) {
    throw new Error('A1セルからフォルダIDを取得できませんでした。URL形式を確認してください。');
  }
  return id;
}

function extractDriveFolderId_(text) {
  const directId = text.match(/^[a-zA-Z0-9_-]{20,}$/);
  if (directId) return directId[0];

  const folderPath = text.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (folderPath) return folderPath[1];

  const queryId = text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (queryId) return queryId[1];

  return '';
}

function listPdfFilesRecursive_(folderId) {
  const result = [];
  const root = DriveApp.getFolderById(folderId);
  walkFolder_(root, result);
  return result;
}

function walkFolder_(folder, collector) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const mimeType = file.getMimeType();
    if (mimeType === MimeType.PDF || mimeType === 'application/pdf') {
      collector.push({
        id: file.getId(),
        name: file.getName(),
        updatedMs: file.getLastUpdated().getTime()
      });
    }
  }

  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    walkFolder_(subFolders.next(), collector);
  }
}

function loadCacheMap_(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.CACHE_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  if (values.length === 0) {
    sheet.getRange(1, 1, 1, 6).setValues([['fileId', 'fileName', 'updatedMs', 'subject', 'payloadJson', 'processedAt']]);
    return {};
  }

  const map = {};
  for (let i = 1; i < values.length; i++) {
    const [fileId, fileName, updatedMs, subject, payloadJson, processedAt] = values[i];
    if (!fileId) continue;
    map[String(fileId)] = {
      fileId: String(fileId),
      fileName: String(fileName || ''),
      updatedMs: String(updatedMs || ''),
      subject: String(subject || ''),
      payloadJson: String(payloadJson || ''),
      processedAt: String(processedAt || '')
    };
  }
  return map;
}

function writeCacheMap_(ss, cache) {
  const sheet = getOrCreateSheet_(ss, CONFIG.CACHE_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, 6).setValues([['fileId', 'fileName', 'updatedMs', 'subject', 'payloadJson', 'processedAt']]);

  const rows = Object.keys(cache).sort().map(key => {
    const c = cache[key];
    return [c.fileId, c.fileName, c.updatedMs, c.subject, c.payloadJson, c.processedAt];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
  sheet.hideSheet();
}

function getCachedPayloadIfFresh_(cache, file) {
  const cached = cache[file.id];
  if (!cached) return null;
  if (String(cached.updatedMs) !== String(file.updatedMs)) return null;
  if (!cached.payloadJson) return null;

  const payload = parseJsonSafe_(cached.payloadJson);
  if (!Array.isArray(payload.items) || payload.items.length === 0) return null;
  payload.sourceFile = payload.sourceFile || file.name;
  payload.fileId = payload.fileId || file.id;
  payload.fileUpdatedMs = payload.fileUpdatedMs || file.updatedMs;
  return payload;
}

function extractUnitsFromPdfWithFallback_(file, apiKey) {
  try {
    return extractUnitsFromPdfDirectGemini_(file, apiKey);
  } catch (directErr) {
    Logger.log(`PDF直接解析失敗: ${file.name} (${file.id}) ${directErr}`);
  }

  const text = extractTextFromPdfViaOcr_(file);
  if (!text.trim()) {
    throw new Error('OCRテキストが空でした。');
  }
  return extractUnitsWithGeminiFromText_(file, text, apiKey);
}

function extractUnitsFromPdfDirectGemini_(file, apiKey) {
  const driveFile = DriveApp.getFileById(file.id);
  const blob = driveFile.getBlob();
  const payload = {
    contents: [{
      parts: [
        { text: buildExtractionPrompt_(file.name) },
        {
          inlineData: {
            mimeType: 'application/pdf',
            data: Utilities.base64Encode(blob.getBytes())
          }
        }
      ]
    }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json'
    }
  };

  const jsonText = callGeminiRaw_(payload, apiKey);
  const parsed = parseJsonSafe_(jsonText);
  parsed.sourceFile = file.name;
  parsed.fileId = file.id;
  parsed.fileUpdatedMs = file.updatedMs;
  return parsed;
}

function extractTextFromPdfViaOcr_(file) {
  const driveFile = DriveApp.getFileById(file.id);
  const mimeType = driveFile.getMimeType();
  if (mimeType !== MimeType.PDF && mimeType !== 'application/pdf') {
    throw new Error(`PDF以外のファイルです。mimeType=${mimeType}`);
  }

  const pdfBlob = driveFile.getBlob();
  const resource = {
    title: `[tmp_ocr] ${file.name}`,
    mimeType: MimeType.GOOGLE_DOCS
  };
  const inserted = Drive.Files.insert(resource, pdfBlob, {
    ocr: true,
    ocrLanguage: CONFIG.OCR_LANGUAGE
  });

  try {
    return DocumentApp.openById(inserted.id).getBody().getText();
  } finally {
    DriveApp.getFileById(inserted.id).setTrashed(true);
  }
}

function extractUnitsWithGeminiFromText_(file, text, apiKey) {
  const prompt = [
    buildExtractionPrompt_(file.name),
    '--- PDF本文ここから ---',
    text.slice(0, 150000),
    '--- PDF本文ここまで ---'
  ].join('\n');
  const jsonText = callGemini_(prompt, apiKey);
  const payload = parseJsonSafe_(jsonText);
  payload.sourceFile = file.name;
  payload.fileId = file.id;
  payload.fileUpdatedMs = file.updatedMs;
  return payload;
}

function callGemini_(prompt, apiKey) {
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json'
    }
  };
  return callGeminiRaw_(payload, apiKey);
}

function callGeminiRaw_(payload, apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(CONFIG.MODEL)}:generateContent?key=${encodeURIComponent(apiKey)}`;

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    throw new Error(`Gemini API error: ${res.getResponseCode()} ${res.getContentText()}`);
  }

  const json = JSON.parse(res.getContentText());
  const text = json?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) {
    throw new Error(`Geminiの応答テキストが空です: ${res.getContentText()}`);
  }
  return text;
}

function buildExtractionPrompt_(fileName) {
  return [
    'あなたは小学校の年間指導計画データを構造化するアシスタントです。',
    '添付資料から「活動時期」「単元名」「配当時数」を抽出してください。',
    '厳密にJSONのみを返してください。コードブロック記法は禁止です。',
    'hoursは必ず数値（整数または小数）にしてください。',
    'monthsは4月〜3月の配列で返し、推定が必要ならperiodに根拠文字列を残してください。',
    '',
    '返却JSONスキーマ:',
    '{',
    '  "subject": "算数 または 社会 など",',
    '  "items": [',
    '    {',
    '      "period": "原文の活動時期（例: 4月, 4〜5月, 通年）",',
    '      "months": ["4月","5月"],',
    '      "unit": "単元名",',
    '      "hours": 3',
    '    }',
    '  ]',
    '}',
    '',
    `ファイル名: ${fileName}`
  ].join('\n');
}

function parseJsonSafe_(text) {
  try {
    return JSON.parse(text);
  } catch (_) {
    const cleaned = text
      .replace(/^```json\s*/i, '')
      .replace(/^```\s*/i, '')
      .replace(/\s*```$/, '')
      .trim();
    return JSON.parse(cleaned);
  }
}

function buildPdfRows_(payloads) {
  return payloads.map(payload => {
    const subject = String(payload.subject || inferSubjectFromFileName_(payload.sourceFile)).trim();
    const items = Array.isArray(payload.items) ? payload.items : [];

    const byMonth = {};
    CONFIG.MONTHS.forEach(m => {
      byMonth[m] = [];
    });

    items.forEach(item => {
      const unit = String(item.unit || '').trim();
      if (!unit) return;

      const hours = Number(item.hours);
      const safeHours = Number.isFinite(hours) ? hours : null;
      const label = `${unit}${safeHours !== null ? `（${safeHours}）` : ''}`;
      const months = normalizeMonths_(item.months, item.period);
      months.forEach(m => {
        if (byMonth[m]) byMonth[m].push(label);
      });
    });

    return {
      fileId: payload.fileId || '',
      fileName: payload.sourceFile || '',
      subject,
      months: byMonth,
      raw: payload
    };
  });
}

function normalizeMonths_(months, periodText) {
  if (Array.isArray(months) && months.length > 0) {
    const normalized = months
      .map(m => normalizeMonthLabel_(String(m)))
      .filter(Boolean);
    if (normalized.length > 0) return [...new Set(normalized)];
  }

  const text = String(periodText || '');
  const matched = [];
  const monthNums = text.match(/(1[0-2]|[1-9])\s*月/g) || [];
  monthNums.forEach(m => {
    const label = normalizeMonthLabel_(m);
    if (label) matched.push(label);
  });
  if (matched.length > 0) return [...new Set(matched)];

  const range = text.match(/(1[0-2]|[1-9])\s*[~〜\-－]\s*(1[0-2]|[1-9])\s*月?/);
  if (range) {
    const start = Number(range[1]);
    const end = Number(range[2]);
    return monthRangeLabels_(start, end);
  }

  if (/通年/.test(text)) return CONFIG.MONTHS.slice();
  return [];
}

function normalizeMonthLabel_(value) {
  const m = value.match(/(1[0-2]|[1-9])\s*月?/);
  if (!m) return '';
  return `${Number(m[1])}月`;
}

function monthRangeLabels_(start, end) {
  const order = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3];
  const startIdx = order.indexOf(start);
  const endIdx = order.indexOf(end);
  if (startIdx < 0 || endIdx < 0) return [];
  if (startIdx <= endIdx) return order.slice(startIdx, endIdx + 1).map(n => `${n}月`);
  return order.slice(startIdx).concat(order.slice(0, endIdx + 1)).map(n => `${n}月`);
}

function inferSubjectFromFileName_(name) {
  const lower = String(name || '').toLowerCase();
  if (lower.includes('算数')) return '算数';
  if (lower.includes('社会')) return '社会';
  return '不明';
}

function renderMainTable_(ss, rowsByPdf) {
  const sheet = getOrCreateSheet_(ss, CONFIG.TABLE_SHEET_NAME);
  sheet.clear();

  const header = ['教科', 'PDFファイル'].concat(CONFIG.MONTHS);
  const rows = rowsByPdf.map(row => {
    return [row.subject, row.fileName].concat(CONFIG.MONTHS.map(m => row.months[m].join('\n')));
  });

  sheet.getRange(1, 1).setValue('活動時期');
  sheet.getRange(1, 2, 1, header.length - 1).merge();
  sheet.getRange(1, 2).setValue('4月〜3月');

  sheet.getRange(2, 1, 1, header.length).setValues([header]);
  if (rows.length > 0) {
    sheet.getRange(3, 1, rows.length, header.length).setValues(rows);
  }

  styleTable_(sheet, rows.length + 2, header.length);
  sheet.setColumnWidth(2, 260);
}

function renderSubjectSheets_(ss, rowsByPdf) {
  const subjects = [...new Set(rowsByPdf.map(r => r.subject).filter(Boolean))];
  subjects.forEach(subject => {
    const name = `計画_${subject}`.slice(0, 95);
    const sheet = getOrCreateSheet_(ss, name);
    sheet.clear();

    const header = ['PDFファイル'].concat(CONFIG.MONTHS);
    const filtered = rowsByPdf.filter(r => r.subject === subject);
    const rows = filtered.map(row => {
      return [row.fileName].concat(CONFIG.MONTHS.map(m => row.months[m].join('\n')));
    });

    sheet.getRange(1, 1).setValue(`教科: ${subject}`);
    sheet.getRange(1, 2, 1, header.length - 1).merge();
    sheet.getRange(1, 2).setValue('4月〜3月');

    sheet.getRange(2, 1, 1, header.length).setValues([header]);
    if (rows.length > 0) {
      sheet.getRange(3, 1, rows.length, header.length).setValues(rows);
    }

    styleTable_(sheet, rows.length + 2, header.length);
    sheet.setColumnWidth(1, 260);
  });
}

function styleTable_(sheet, totalRows, totalCols) {
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);

  for (let i = 1; i <= totalCols; i++) {
    if (i !== 2) sheet.setColumnWidth(i, 140);
  }

  sheet.getRange(1, 1, 1, totalCols)
    .setBackground('#d9ead3')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange(2, 1, 1, totalCols)
    .setBackground('#cfe2f3')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  if (totalRows >= 3 && totalCols >= 2) {
    sheet.getRange(3, 1, totalRows - 2, 1)
      .setBackground('#f3f3f3')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.getRange(3, 2, totalRows - 2, totalCols - 1)
      .setWrap(true)
      .setVerticalAlignment('top');
  }

  sheet.getRange(1, 1, totalRows, totalCols).setBorder(true, true, true, true, true, true);
}

function renderRawSheet_(ss, payloads) {
  const raw = getOrCreateSheet_(ss, CONFIG.RAW_SHEET_NAME);
  raw.clear();
  raw.getRange(1, 1, 1, 6).setValues([['sourceFile', 'subject', 'period', 'months', 'unit', 'hours']]);

  const rows = [];
  payloads.forEach(payload => {
    const subject = payload.subject || '';
    const items = Array.isArray(payload.items) ? payload.items : [];
    items.forEach(item => {
      rows.push([
        payload.sourceFile || '',
        subject,
        item.period || '',
        Array.isArray(item.months) ? item.months.join(',') : '',
        item.unit || '',
        item.hours != null ? item.hours : ''
      ]);
    });
  });

  if (rows.length > 0) {
    raw.getRange(2, 1, rows.length, 6).setValues(rows);
  }
  raw.setFrozenRows(1);
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function clearProcessingCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  const sheet = ss.getSheetByName(CONFIG.CACHE_SHEET_NAME);
  if (!sheet) return;
  sheet.clear();
  sheet.getRange(1, 1, 1, 6).setValues([['fileId', 'fileName', 'updatedMs', 'subject', 'payloadJson', 'processedAt']]);
}
