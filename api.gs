function doGet(e) {
  const sheetId = "1BNgNVakEKw1YO47LzmulV9aweytIo18Hxtq5ihG1wV0";
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("news");
  const lastRow = sheet.getLastRow();

  const today = (e && e.parameter && e.parameter.date)
    ? e.parameter.date
    : todayStr_();

  // 7 個資產，按真實分類順序
  const ASSETS = [
    { key: "GLOBAL", name: "\u5168\u7403\u6307\u6578" },
    { key: "TW0050", name: "0050" },
    { key: "BOND", name: "\u7f8e\u50b5" },
    { key: "TSMC", name: "\u53f0\u7a4d\u96fb" },
    { key: "GOOGL", name: "Google" },
    { key: "TSLA", name: "\u7279\u65af\u62c9" },
    { key: "BTC", name: "\u6bd4\u7279\u5e63" },
    { key: "GOLD", name: "\u9ec3\u91d1" }
  ];

  if (lastRow < 2) {
    return jsonResponse_({ date: today, assets: emptyAssets_(ASSETS) });
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  // 7 天範圍
  const sevenDaysAgo = new Date(today);
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 6);

  const result = ASSETS.map(function(asset) {
    const assetRows = data.filter(function(row) {
      return row[1] === asset.key;
    });

    // 今日資料
    const todayRows = assetRows.filter(function(row) {
      return rowDate_(row) === today;
    });

    // 今日判讀（取 I 欄第一個非空的）
    let todaySummary = "";
    for (let i = 0; i < todayRows.length; i++) {
      if (todayRows[i][8]) { todaySummary = todayRows[i][8]; break; }
    }

    // 7 天強訊號數
    const weekStrongCount = assetRows.filter(function(row) {
      const d = rowDate_(row);
      return d >= Utilities.formatDate(sevenDaysAgo, "Asia/Taipei", "yyyy-MM-dd")
        && d <= today
        && row[6] === "\u5f37";
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

function emptyAssets_(ASSETS) {
  return ASSETS.map(function(a) {
    return { key: a.key, name: a.name, today: "", weekStrong: 0, hasNewsToday: false };
  });
}

function rowDate_(row) {
  return row[0] instanceof Date
    ? Utilities.formatDate(row[0], "Asia/Taipei", "yyyy-MM-dd")
    : String(row[0]);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function todayStr_() {
  return Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd");
}