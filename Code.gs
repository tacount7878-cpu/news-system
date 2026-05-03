// 標籤標準化
function normalizeTag(rawTag) {
  const map = {
    "0050": "TW0050",
    "0050tw": "TW0050",
    "tw0050": "TW0050",
    "台積電": "TSMC",
    "2330": "TSMC",
    "tsm": "TSMC",
    "比特幣": "BTC",
    "bitcoin": "BTC",
    "黃金": "GOLD",
    "gold": "GOLD",
    "美債": "BOND",
    "bond": "BOND",
    "other": "OTHER"
  };

  const key = String(rawTag).trim().toLowerCase();
  return map[key] || String(rawTag).trim().toUpperCase();
}

// 清理歷史標籤（手動跑一次）
function cleanHistoricalTags() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("news");
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B 欄

  const normalized = data.map(function(row) {
    return [normalizeTag(row[0])];
  });

  sheet.getRange(2, 2, normalized.length, 1).setValues(normalized);
  Logger.log("已清理 " + normalized.length + " 筆標籤");
}
/**
 * News Fetcher - ?身??
 * ??敶梢?湧?鞈???啗?
 * 甈?嚗?交? B璅惜 C璅? D??? E?澆??? F?? G撘瑕漲 H憿? I隞?方?
 */

function autoFetchNews() {
  const sheetName = "news";
  const headers = ["\u65e5\u671f","\u6a19\u7c64","\u6a19\u984c","\u9023\u7d50","\u767c\u5e03\u6642\u9593","\u60c5\u7dd2","\u5f37\u5ea6","\u985e\u578b","\u4eca\u65e5\u5224\u8b80"];

  const RSS_SOURCES = [
    // Google News嚗敹??萄?嚗項?蝛??050??～ed????萸??∪之?扎?舀??oogle???孵馳????
    "https://news.google.com/rss/search?q=\u53f0\u7a4d\u96fb+OR+TSMC+OR+0050+OR+\u53f0\u80a1&hl=zh-TW&gl=TW&ceid=TW:zh-Hant",
    "https://news.google.com/rss/search?q=Fed+OR+\u806f\u6e96\u6703+OR+\u5229\u7387+OR+\u7f8e\u50b5&hl=zh-TW&gl=TW&ceid=TW:zh-Hant",
    "https://news.google.com/rss/search?q=Tesla+OR+TSLA+OR+Google+OR+Alphabet&hl=zh-TW&gl=TW&ceid=TW:zh-Hant",
    "https://news.google.com/rss/search?q=Bitcoin+OR+BTC+OR+\u6bd4\u7279\u5e63+OR+Gold+OR+\u9ec3\u91d1&hl=zh-TW&gl=TW&ceid=TW:zh-Hant",

    // ??鞎∠?銝餅?嚗歇撽??舐嚗?
    "https://feeds.reuters.com/reuters/businessNews",
    "https://www.cna.com.tw/rss/aall.xml",
    "https://www.cnbc.com/id/100003114/device/rss/rss.html",
    "https://feeds.marketwatch.com/marketwatch/topstories/",

    // ??鞎典馳
    "https://www.coindesk.com/arc/outboundfeeds/rss/"
  ];

  const BLACKLIST = ["facebook","cmoney","line","ptt","youtube"];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheet.getRange(1,1,1,headers.length).setValues([headers]);

  const sinceMs = Date.now() - 24*60*60*1000;
  const tz = Session.getScriptTimeZone();

  const existingLinks = getExistingLinks_(sheet);
  const processedTitles = new Set();

  let allItems = [];
  RSS_SOURCES.forEach(url => {
    const items = fetchRssItems_(url);
    if (items.length) allItems = allItems.concat(items);
  });

  let rows = [];

  for (let i = 0; i < allItems.length; i++) {
    const it = allItems[i];
    if (!it.pubDate || isNaN(it.pubDate.getTime())) continue;
    if (it.pubDate.getTime() < sinceMs) continue;

    const link = normalizeUrl_(it.link);
    if (!link || BLACKLIST.some(b => link.toLowerCase().includes(b))) continue;
    if (existingLinks[link]) continue;

    const cleanTitle = normalizeTitle_(it.title);
    if (processedTitles.has(cleanTitle)) continue;

    const label = classifyAsset_(it.title);

    rows.push([
      Utilities.formatDate(it.pubDate, tz, "yyyy-MM-dd"),
      label,
      it.title,
      link,
      Utilities.formatDate(it.pubDate, tz, "yyyy-MM-dd HH:mm:ss"),
      "", "", "", ""
    ]);

    processedTitles.add(cleanTitle);
    existingLinks[link] = true;

    if (rows.length >= 50) break;
  }

  if (rows.length) {
    rows.sort((a,b) => new Date(b[4]) - new Date(a[4]));
    sheet.getRange(sheet.getLastRow()+1,1,rows.length,headers.length).setValues(rows);
  }
}

function normalizeUrl_(url) {
  if (!url) return "";
  try {
    const decoded = decodeURIComponent(url);
    const m = decoded.match(/url=(https?:\/\/[^&]+)/);
    if (m) return m[1].split("?")[0];
  } catch(e){}
  return url.split("?")[0];
}

function normalizeTitle_(title) {
  return (title || "")
    .split(" - ")[0]
    .split("\uff5c")[0]
    .split(":")[0]
    .replace(/\s+/g,"")
    .toLowerCase()
    .trim();
}

function fetchRssItems_(url) {
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return [];
    const doc = XmlService.parse(resp.getContentText());
    const items = doc.getRootElement().getChild("channel").getChildren("item");
    return items.map(it => ({
      title: it.getChildText("title") || "",
      link: it.getChildText("link") || "",
      pubDate: new Date(it.getChildText("pubDate"))
    }));
  } catch(e) {
    return [];
  }
}

/**
 * ??閬?嚗???UI ?∠???7 ????
 * GLOBAL嚗???賂?/ TW0050嚗?∪之?歹?/ BOND嚗??蛛?
 * TSMC嚗蝛嚗? GOOGL / TSLA / BTC / GOLD
 */
function classifyAsset_(title) {
  const t = title.toLowerCase();

  // ??芸?嚗?◤憭抒??嚗?
  if (t.includes("\u53f0\u7a4d") || t.includes("tsmc") || t.includes("2330")) return "TSMC";
  if (t.includes("tesla") || t.includes("tsla") || t.includes("\u7279\u65af\u62c9")) return "TSLA";
  if (t.includes("google") || t.includes("alphabet") || t.includes("googl")) return "GOOGL";

  // ?? / 暺?
  if (t.includes("bitcoin") || t.includes("btc") || t.includes("\u6bd4\u7279\u5e63")) return "BTC";
  if (t.includes("gold") || t.includes("\u9ec3\u91d1") || t.includes("gldm")) return "GOLD";

  // 蝢 / Fed / ?拍?嚗蔣??BOND嚗?
  if (t.includes("fed") || t.includes("\u806f\u6e96\u6703") || t.includes("\u5229\u7387") ||
      t.includes("\u6b96\u5229\u7387") || t.includes("treasury") || t.includes("\u7f8e\u50b5") ||
      t.includes("powell") || t.includes("\u9bee\u91d1\u6708")) return "BOND";

  // ?函??嚗&P / ???? / 蝢憭抒嚗?
  if (t.includes("s&p") || t.includes("\u90a3\u65af\u9054\u514b") || t.includes("nasdaq") ||
      t.includes("\u9053\u743c") || t.includes("\u7f8e\u80a1") || t.includes("vt") ||
      t.includes("vwra")) return "GLOBAL";

  // ?啗憭抒
  if (t.includes("0050") || t.includes("\u53f0\u80a1") || t.includes("\u52a0\u6b0a") ||
      t.includes("\u53f0\u7063\u52a0\u6b0a")) return "TW0050";

  return "OTHER";
}

function getExistingLinks_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const values = sheet.getRange(2,4,lastRow-1,1).getValues();
  const map = {};
  values.forEach(v => { if (v[0]) map[String(v[0])] = true; });
  return map;
}

function runTodayPayload() {
  applyAnalysisPayloadByBlankRows({
    fgh: [
      ["中性","中","資金"],
      ["中性","中","技術"],
      ["中性","中","技術"],
      ["中性","中","資金"],
      ["中性","中","資金"],
      ["中性","中","技術"],
      ["中性","弱","雜訊"],
      ["中性","中","政策"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["負面","中","政策"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["負面","強","政策"],
      ["中性","弱","雜訊"],
      ["正面","強","技術"],
      ["正面","強","技術"],
      ["正面","中","技術"],
      ["負面","中","資金"],
      ["負面","中","政策"],
      ["負面","中","政策"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["正面","中","技術"],
      ["正面","中","資金"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["正面","中","技術"],
      ["正面","中","技術"],
      ["負面","強","政策"],
      ["中性","弱","雜訊"],
      ["正面","強","技術"],
      ["正面","中","技術"],
      ["負面","中","政策"],
      ["中性","中","資金"],
      ["中性","弱","雜訊"],
      ["正面","中","技術"],
      ["負面","中","政策"],
      ["負面","中","情緒"],
      ["中性","弱","雜訊"],
      ["中性","中","政策"],
      ["正面","中","技術"],
      ["中性","中","情緒"],
      ["中性","中","情緒"],
      ["正面","中","技術"],
      ["負面","強","政策"],
      ["中性","中","政策"],
      ["負面","中","政策"],
      ["正面","中","技術"],
      ["中性","弱","雜訊"],
      ["中性","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","中","資金"],
      ["中性","中","資金"],
      ["中性","弱","雜訊"],
      ["負面","中","政策"],
      ["負面","中","情緒"],
      ["負面","強","政策"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","中","資金"],
      ["中性","中","技術"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","中","資金"],
      ["中性","弱","資金"],
      ["中性","弱","雜訊"],
      ["中性","中","情緒"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["正面","中","技術"],
      ["負面","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["負面","中","資金"],
      ["負面","中","情緒"],
      ["中性","弱","雜訊"],
      ["負面","中","情緒"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","中","資金"],
      ["正面","中","資金"],
      ["中性","中","技術"],
      ["中性","弱","雜訊"],
      ["中性","中","技術"],
      ["負面","中","情緒"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","中","資金"],
      ["中性","中","政策"],
      ["中性","弱","雜訊"],
      ["正面","中","技術"],
      ["中性","中","資金"],
      ["中性","中","技術"],
      ["中性","中","政策"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","中","技術"],
      ["正面","中","資金"],
      ["中性","中","技術"],
      ["正面","強","技術"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","中","政策"],
      ["中性","中","政策"],
      ["中性","中","政策"],
      ["中性","中","技術"],
      ["中性","中","技術"],
      ["中性","弱","雜訊"],
      ["負面","中","政策"],
      ["中性","中","政策"],
      ["負面","中","政策"],
      ["中性","中","政策"],
      ["中性","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["負面","中","情緒"],
      ["中性","中","情緒"],
      ["中性","弱","雜訊"],
      ["負面","強","政策"],
      ["負面","強","政策"],
      ["正面","中","資金"],
      ["中性","中","情緒"],
      ["中性","弱","雜訊"],
      ["中性","中","資金"],
      ["中性","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","中","資金"],
      ["正面","強","技術"],
      ["正面","中","資金"],
      ["正面","中","資金"],
      ["負面","中","政策"],
      ["負面","強","政策"],
      ["中性","弱","雜訊"],
      ["正面","中","技術"],
      ["中性","弱","雜訊"],
      ["正面","中","技術"],
      ["負面","強","政策"],
      ["正面","中","資金"],
      ["負面","中","情緒"],
      ["負面","中","政策"],
      ["中性","弱","雜訊"],
      ["負面","中","資金"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["負面","中","政策"],
      ["正面","中","資金"],
      ["中性","中","情緒"],
      ["中性","弱","雜訊"],
      ["負面","中","政策"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["負面","強","政策"],
      ["中性","中","政策"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","中","情緒"],
      ["負面","中","政策"],
      ["中性","弱","雜訊"],
      ["正面","中","技術"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","中","技術"],
      ["中性","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","中","資金"],
      ["正面","強","技術"],
      ["負面","強","政策"],
      ["負面","強","政策"],
      ["正面","中","資金"],
      ["中性","中","資金"],
      ["正面","中","資金"],
      ["中性","中","資金"],
      ["中性","弱","雜訊"],
      ["負面","強","政策"],
      ["中性","弱","雜訊"],
      ["中性","中","技術"],
      ["正面","中","技術"],
      ["負面","中","情緒"],
      ["負面","強","政策"],
      ["正面","強","技術"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","中","情緒"],
      ["負面","強","政策"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["正面","強","技術"],
      ["中性","弱","雜訊"],
      ["中性","中","技術"],
      ["正面","中","技術"],
      ["正面","強","技術"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","弱","情緒"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["正面","中","技術"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["負面","強","政策"],
      ["中性","中","技術"],
      ["中性","中","資金"],
      ["中性","中","資金"],
      ["中性","中","技術"],
      ["負面","中","政策"],
      ["中性","中","技術"],
      ["中性","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["負面","中","政策"],
      ["中性","中","資金"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","中","技術"],
      ["負面","強","政策"],
      ["中性","中","情緒"],
      ["正面","中","技術"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["正面","中","技術"],
      ["負面","強","政策"],
      ["正面","中","技術"],
      ["中性","弱","雜訊"],
      ["負面","中","政策"],
      ["中性","中","資金"],
      ["中性","中","技術"],
      ["中性","中","情緒"],
      ["中性","中","資金"],
      ["正面","中","技術"],
      ["中性","中","技術"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["負面","中","政策"],
      ["中性","中","技術"],
      ["中性","中","資金"],
      ["中性","中","資金"],
      ["中性","中","技術"],
      ["中性","弱","雜訊"],
      ["中性","中","技術"],
      ["正面","中","技術"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","中","技術"],
      ["正面","中","技術"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["負面","強","政策"],
      ["正面","強","技術"],
      ["中性","中","技術"],
      ["中性","弱","雜訊"],
      ["中性","中","資金"],
      ["中性","弱","雜訊"],
      ["正面","中","資金"],
      ["中性","中","技術"],
      ["負面","中","政策"],
      ["中性","中","情緒"],
      ["負面","中","政策"],
      ["中性","弱","雜訊"],
      ["中性","中","資金"],
      ["中性","中","技術"],
      ["正面","中","資金"],
      ["中性","弱","雜訊"],
      ["中性","中","技術"],
      ["中性","中","技術"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["負面","強","政策"],
      ["中性","中","技術"],
      ["中性","中","技術"],
      ["中性","中","情緒"],
      ["負面","中","政策"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["中性","弱","雜訊"],
      ["負面","中","政策"],
      ["正面","中","資金"],
      ["負面","中","政策"],
      ["中性","弱","雜訊"],
      ["負面","中","情緒"]
    ],
    iMap: {
      "2026-05-02|BOND": "Fed轉鷹壓抑降息預期 美債先觀察",
      "2026-05-02|TSMC": "封裝人才與CoPoS題材延續 長期邏輯未變",
      "2026-05-02|BTC": "BTC站回78K但清算風險升溫",
      "2026-05-02|GOOGL": "TPU與AI財報支撐 Google動能延續",
      "2026-05-02|GOLD": "金價拉回與雜訊並存 避險邏輯未變",
      "2026-05-02|TW0050": "台股ETF資金熱度仍高 指數化配置不變",
      "2026-05-02|GLOBAL": "美股創高但伊朗與油價風險升溫",
      "2026-05-02|TSLA": "薪酬與治理雜訊升溫 先觀察",
      "2026-05-03|TSLA": "Semi量產下線 屬產品利多非結構改變",
      "2026-05-03|TW0050": "四大利多撐台股 但高點風險需留意",
      "2026-05-03|BTC": "BTC動能延續但槓桿修正風險升溫",
      "2026-05-03|BOND": "Fed官員轉鷹 美債仍受利率壓制",
      "2026-05-03|GOLD": "避險需求仍在 但黃金週雜訊偏多",
      "2026-05-03|GLOBAL": "美股抬轎但短線震盪風險升溫",
      "2026-05-03|TSMC": "新廠與先進製程利多 但處置風險升溫",
      "2026-05-03|GOOGL": "TPU供應鏈升級 Google AI動能延續"
    }
  });
}

function applyAnalysisPayloadByBlankRows(payload) {
  if (!payload || !payload.fgh || !payload.iMap) {
    throw new Error("payload 缺少必要欄位：fgh / iMap");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("news");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("news 工作表沒有資料");
  }

  const tz = "Asia/Taipei";
  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const blankRows = [];

  for (let i = 0; i < data.length; i++) {
    const fValue = String(data[i][5] || "").trim();
    if (fValue === "") {
      blankRows.push({
        sheetRow: i + 2,
        values: data[i]
      });
    }
  }

  if (blankRows.length === 0) {
    throw new Error("沒有 F 欄空白列，不需要補資料");
  }

  if (blankRows.length !== payload.fgh.length) {
    throw new Error(
      "FGH 筆數不一致：Sheet 目前 F 欄空白 " +
      blankRows.length +
      " 筆，但 payload.fgh 是 " +
      payload.fgh.length +
      " 筆。停止寫入，避免錯位。"
    );
  }

  const fghValues = payload.fgh.map(function(triplet, index) {
    if (!Array.isArray(triplet) || triplet.length !== 3) {
      throw new Error("第 " + (index + 1) + " 筆 FGH 格式錯誤");
    }
    return [
      String(triplet[0] || "").trim(),
      String(triplet[1] || "").trim(),
      String(triplet[2] || "").trim()
    ];
  });

  for (let i = 0; i < blankRows.length; i++) {
    sheet.getRange(blankRows[i].sheetRow, 6, 1, 3).setValues([fghValues[i]]);
  }

  const written = {};
  let iWritten = 0;

  for (let i = 0; i < blankRows.length; i++) {
    const row = blankRows[i].values;
    const dateKey = normalizePayloadDate_(row[0], tz);
    const tag = String(row[1] || "").trim();
    if (tag === "OTHER") continue;

    const mapKey = dateKey + "|" + tag;
    if (!payload.iMap[mapKey]) continue;
    if (written[mapKey]) continue;

    sheet.getRange(blankRows[i].sheetRow, 9).setValue(payload.iMap[mapKey]);
    written[mapKey] = true;
    iWritten++;
  }

  Logger.log("✅ 空白列補寫完成");
  Logger.log("FGH：" + fghValues.length + " 筆");
  Logger.log("I 欄：" + iWritten + " 個 日期+資產：" + Object.keys(written).join(", "));
}

function normalizePayloadDate_(value, tz) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, tz || "Asia/Taipei", "yyyy-MM-dd");
  }

  const s = String(value || "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{4}\/\d{1,2}\/\d{1,2}/.test(s)) {
    const parts = s.split(/[\/\s]/);
    const y = parts[0];
    const m = ("0" + parts[1]).slice(-2);
    const d = ("0" + parts[2]).slice(-2);
    return y + "-" + m + "-" + d;
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, tz || "Asia/Taipei", "yyyy-MM-dd");
  }

  return s.slice(0, 10).replace(/\//g, "-");
}

/**
 * 把 AI 產生的 FGHI 寫回 Sheet
 * FGH：依照當天新聞順序逐列寫入 F~H
 * I：依照 iMap，寫在每個資產當天第一筆 I 欄
 */
function applyAnalysisPayload(payload) {
  if (!payload || !payload.date || !payload.fgh || !payload.iMap) {
    throw new Error("payload 缺少必要欄位：date / fgh / iMap");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("news");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("news 工作表沒有資料");
  }

  const tz = "Asia/Taipei";
  const targetDate = payload.date;
  const fghList = payload.fgh;
  const iMap = payload.iMap;

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  const targetRows = [];
  for (let i = 0; i < data.length; i++) {
    const rowDate = data[i][0] instanceof Date
      ? Utilities.formatDate(data[i][0], tz, "yyyy-MM-dd")
      : String(data[i][0]).slice(0, 10);

    if (rowDate === targetDate) {
      targetRows.push({
        sheetRow: i + 2,
        values: data[i]
      });
    }
  }

  if (targetRows.length === 0) {
    throw new Error("找不到日期 " + targetDate + " 的資料");
  }

  if (targetRows.length !== fghList.length) {
    throw new Error(
      "FGH 筆數不一致：Sheet 當天資料 " +
      targetRows.length +
      " 筆，但 payload.fgh 是 " +
      fghList.length +
      " 筆。停止寫入，避免錯位。"
    );
  }

  const fghValues = fghList.map(function(triplet, index) {
    if (!Array.isArray(triplet) || triplet.length !== 3) {
      throw new Error("第 " + (index + 1) + " 筆 FGH 格式錯誤");
    }

    return [
      String(triplet[0] || "").trim(),
      String(triplet[1] || "").trim(),
      String(triplet[2] || "").trim()
    ];
  });

  const firstRow = targetRows[0].sheetRow;
  sheet.getRange(firstRow, 6, fghValues.length, 3).setValues(fghValues);

  const written = {};
  let iWritten = 0;

  for (let i = 0; i < targetRows.length; i++) {
    const tag = String(targetRows[i].values[1] || "").trim();

    if (!iMap[tag]) continue;
    if (written[tag]) continue;

    sheet.getRange(targetRows[i].sheetRow, 9).setValue(iMap[tag]);
    written[tag] = true;
    iWritten++;
  }

  Logger.log("✅ 寫入完成");
  Logger.log("日期：" + targetDate);
  Logger.log("FGH：" + fghValues.length + " 筆");
  Logger.log("I 欄：" + iWritten + " 個資產：" + Object.keys(written).join(", "));
}
