function doGet(e) {
  if (e && e.parameter && e.parameter.action === "blankRows") {
    return getBlankRowsForAnalysis_(e);
  }

  const sheetId = "1BNgNVakEKw1YO47LzmulV9aweytIo18Hxtq5ihG1wV0";
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("news");
  const lastRow = sheet.getLastRow();

  const today = (e && e.parameter && e.parameter.date)
    ? e.parameter.date
    : todayStr_();

  const ASSETS = [
    { key: "GLOBAL", name: "全球指數" },
    { key: "TW0050", name: "0050" },
    { key: "BOND", name: "美債" },
    { key: "TSMC", name: "台積電" },
    { key: "GOOGL", name: "Google" },
    { key: "TSLA", name: "特斯拉" },
    { key: "BTC", name: "比特幣" },
    { key: "GOLD", name: "黃金" }
  ];

  if (lastRow < 2) {
    return jsonResponse_({ date: today, assets: emptyAssets_(ASSETS) });
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  const sevenDaysAgo = new Date(today + "T00:00:00+08:00");
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 6);
  const sevenDaysAgoStr = Utilities.formatDate(sevenDaysAgo, "Asia/Taipei", "yyyy-MM-dd");

  const result = ASSETS.map(function(asset) {
    const assetRows = data.filter(function(row) {
      return String(row[1] || "").trim() === asset.key;
    });

    const todayRows = assetRows.filter(function(row) {
      return rowDate_(row) === today;
    });

    let todaySummary = "";
    for (let i = todayRows.length - 1; i >= 0; i--) {
      if (todayRows[i][8]) {
        todaySummary = todayRows[i][8];
        break;
      }
    }

    const weekStrongCount = assetRows.filter(function(row) {
      const d = rowDate_(row);
      return d >= sevenDaysAgoStr
        && d <= today
        && String(row[6] || "").trim() === "強";
    }).length;

    return {
      key: asset.key,
      name: asset.name,
      today: todaySummary,
      weekStrong: weekStrongCount,
      hasNewsToday: todayRows.length > 0
    };
  });

  return jsonResponse_({ date: today, assets: result });
}

function getBlankRowsForAnalysis_(e) {
  const sheetId = "1BNgNVakEKw1YO47LzmulV9aweytIo18Hxtq5ihG1wV0";
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("news");
  const lastRow = sheet.getLastRow();
  const tz = "Asia/Taipei";

  const rawLimit = Number(e.parameter.limit || 200);
  const limit = Math.max(1, Math.min(rawLimit, 500));

  const rawDays = e.parameter.days;
  const hasDays = rawDays !== undefined && rawDays !== null && String(rawDays).trim() !== "";
  const days = hasDays ? Math.max(1, Number(rawDays)) : null;

  if (lastRow < 2) {
    return textResponse_("mode=blankRows\ncount=0\ntotalBlank=0\n");
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  let maxDate = "";
  data.forEach(function(row) {
    const d = normalizeDateForAnalysis_(row[0], tz);
    if (d && d > maxDate) maxDate = d;
  });

  let startDate = "";
  if (days && maxDate) {
    const start = new Date(maxDate + "T00:00:00+08:00");
    start.setDate(start.getDate() - days + 1);
    startDate = Utilities.formatDate(start, tz, "yyyy-MM-dd");
  }

  const lines = [];
  let totalBlank = 0;
  let firstRow = null;
  let lastMatchedRow = null;

  for (let i = 0; i < data.length; i++) {
    const rowNumber = i + 2;
    const row = data[i];

    const date = normalizeDateForAnalysis_(row[0], tz);
    const tag = String(row[1] || "").trim();
    const title = String(row[2] || "").trim();

    const f = String(row[5] || "").trim();
    const g = String(row[6] || "").trim();
    const h = String(row[7] || "").trim();
    const iValue = String(row[8] || "").trim();

    if (!date) continue;
    if (days && (date < startDate || date > maxDate)) continue;
    if (f && g && h) continue;

    totalBlank++;

    if (lines.length >= limit) continue;

    if (firstRow === null) firstRow = rowNumber;
    lastMatchedRow = rowNumber;

    lines.push(rowNumber + "|" + date + "|" + tag + "|i=" + (iValue ? "Y" : "N") + "|" + title);
  }

  const header = [
    "mode=blankRows",
    "days=" + (days || "ALL"),
    "limit=" + limit,
    "dateRange=" + (days ? startDate + "~" + maxDate : "ALL"),
    "rowRange=" + (firstRow === null ? "" : firstRow + "-" + lastMatchedRow),
    "count=" + lines.length,
    "totalBlank=" + totalBlank,
    ""
  ];

  return textResponse_(header.concat(lines).join("\n"));
}

function emptyAssets_(ASSETS) {
  return ASSETS.map(function(a) {
    return { key: a.key, name: a.name, today: "", weekStrong: 0, hasNewsToday: false };
  });
}

function rowDate_(row) {
  return normalizeDateForAnalysis_(row[0], "Asia/Taipei");
}

function normalizeDateForAnalysis_(value, tz) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, tz || "Asia/Taipei", "yyyy-MM-dd");
  }

  const s = String(value || "").trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  if (/^\d{4}\/\d{1,2}\/\d{1,2}/.test(s)) {
    const parts = s.split(/[\/\s]/);
    return parts[0] + "-" + ("0" + parts[1]).slice(-2) + "-" + ("0" + parts[2]).slice(-2);
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, tz || "Asia/Taipei", "yyyy-MM-dd");
  }

  return "";
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function textResponse_(text) {
  return ContentService
    .createTextOutput(text)
    .setMimeType(ContentService.MimeType.TEXT);
}

function todayStr_() {
  return Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd");
}
