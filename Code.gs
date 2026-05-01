/**
 * News Fetcher - Final Stable (8 columns)
 * A 日期
 * B 標籤
 * C 標題
 * D 連結
 * E 發布時間
 * F 情緒（留空）
 * G 強度（留空）
 * H 類型（留空）
 */

function autoFetchNews() {
  const sheetName = "news";
  const headers = ["日期","標籤","標題","連結","發布時間","情緒","強度","類型"];

  const RSS_SOURCES = [
    // Google（掃描）
    "https://news.google.com/rss/search?q=0050 OR 台積電 OR TSMC OR Tesla OR TSLA OR Bitcoin OR BTC OR 比特幣&hl=zh-TW&gl=TW&ceid=TW:zh-Hant",

    // 核心來源
    "https://feeds.reuters.com/reuters/businessNews",
    "https://www.cna.com.tw/rss/aall.xml",

    // 市場來源
    "https://www.cnbc.com/id/100003114/device/rss/rss.html",
    "https://feeds.marketwatch.com/marketwatch/topstories/",

    // 資產來源
    "https://news.search.yahoo.com/rss?p=TSMC+Tesla+Bitcoin",
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

  // 多來源合併
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

    const label = classifyAsset_(cleanTitle);

    rows.push([
      Utilities.formatDate(it.pubDate, tz, "yyyy-MM-dd"),  // A 日期
      label,                                                // B 標籤
      it.title,                                             // C 標題
      link,                                                 // D 連結
      Utilities.formatDate(it.pubDate, tz, "yyyy-MM-dd HH:mm:ss"), // E 發布時間
      "", "", ""                                            // FGH 留空
    ]);

    processedTitles.add(cleanTitle);
    existingLinks[link] = true;

    if (rows.length >= 30) break;
  }

  if (rows.length) {
    rows.sort((a,b) => new Date(b[4]) - new Date(a[4]));
    sheet.getRange(sheet.getLastRow()+1,1,rows.length,headers.length).setValues(rows);
  }
}

// --- 工具函式 ---

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
    .split("｜")[0]
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

function classifyAsset_(title) {
  const t = title.toLowerCase();

  if (t.includes("台積") || t.includes("tsmc") || t.includes("2330")) return "TSMC";
  if (t.includes("0050")) return "0050";
  if (t.includes("00679b") || t.includes("債")) return "BOND";
  if (t.includes("tesla") || t.includes("tsla")) return "TSLA";
  if (t.includes("google") || t.includes("alphabet")) return "GOOGL";
  if (t.includes("bitcoin") || t.includes("btc") || t.includes("比特幣")) return "BTC";
  if (t.includes("vt") || t.includes("vwra")) return "GLOBAL";
  if (t.includes("gold") || t.includes("黃金")) return "GOLD";

  return "OTHER";
}

function getExistingLinks_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const values = sheet.getRange(2,4,lastRow-1,1).getValues();
  const map = {};

  values.forEach(v => {
    if (v[0]) map[String(v[0])] = true;
  });

  return map;
}