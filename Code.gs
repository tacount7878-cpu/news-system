/**
 * Google Apps Script - News fetcher V1.5 (Final)
 * - Event filtering (POS/NEG/NEUTRAL gate)
 * - Blacklist filtering
 * - Title deduplication
 * - Google News redirect normalization
 * - Append-only write (stable)
 */

const EVENT_RULES = {
  positive: [
    "財報","營收","創新高","上修","調升","獲利","eps","接單","訂單","擴產","投資","資本支出",
    "ai","先進製程","2奈米","3奈米","合作","合約","簽約","併購","etf流入","資金流入","升評"
  ],
  negative: [
    "下修","衰退","減產","裁員","虧損","需求疲弱","庫存","去庫存","降評","罰款",
    "監管","禁令","制裁","競爭加劇","良率問題","延遲","出貨不如預期","資金流出"
  ]
};

function autoFetchNews() {
  const sheetName = "news";
  const headers = ["日期", "標籤", "標題", "連結", "發布時間"];
  const BLACKLIST_SOURCES = ["facebook", "cmoney", "line today", "爆料", "ptt"];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const now = new Date();
  const sinceMs = now.getTime() - 24 * 60 * 60 * 1000;

  const query = "0050 OR 元大台灣50 OR 台積電 OR TSMC OR Tesla OR TSLA OR Bitcoin OR BTC OR 比特幣 OR VWRA OR 00679B";
  const rssUrl = "https://news.google.com/rss/search?q=" + encodeURIComponent(query) + "&hl=zh-TW&gl=TW&ceid=TW:zh-Hant";

  const items = fetchRssItems_(rssUrl);
  if (!items.length) return;

  const existingLinks = getExistingLinks_(sheet);
  const processedTitles = new Set();
  const tz = Session.getScriptTimeZone();
  let rows = [];

  for (let i = 0; i < items.length; i++) {
    let it = items[i];
    if (!it.pubDate || isNaN(it.pubDate.getTime()) || it.pubDate.getTime() < sinceMs) continue;

    let link = normalizeUrl_(it.link);
    if (BLACKLIST_SOURCES.some(b => link.toLowerCase().includes(b))) continue;

    let cleanTitle = it.title
      .split(" - ")[0]
      .split("｜")[0]
      .split(":")[0]
      .replace(/\s+/g, "")
      .trim();

    let event = classifyEvent_(cleanTitle);
    if (!event) continue;

    if (processedTitles.has(cleanTitle) || existingLinks[link]) continue;

    let label = classifyAsset(cleanTitle);
    let dateStr = Utilities.formatDate(it.pubDate, tz, "yyyy-MM-dd");
    let pubStr = Utilities.formatDate(it.pubDate, tz, "yyyy-MM-dd HH:mm:ss");

    rows.push([dateStr, label, it.title, link, pubStr]);

    processedTitles.add(cleanTitle);
    existingLinks[link] = true;

    if (rows.length >= 20) break;
  }

  if (rows.length) {
    rows.sort((a, b) => new Date(b[4]) - new Date(a[4]));
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }
}

function classifyEvent_(title) {
  const t = (title || "").toLowerCase();
  let pos = 0, neg = 0;

  EVENT_RULES.positive.forEach(k => { if (t.includes(k)) pos++; });
  EVENT_RULES.negative.forEach(k => { if (t.includes(k)) neg++; });

  if (pos === 0 && neg === 0) return null;
  if (pos > neg) return "POS";
  if (neg > pos) return "NEG";
  return "NEUTRAL";
}

function normalizeUrl_(url) {
  if (!url) return "";
  try {
    const decoded = decodeURIComponent(url);
    const m = decoded.match(/url=(https?:\/\/[^&]+)/);
    if (m) return m[1].split("?")[0];
  } catch (e) {}
  return url.split("?")[0];
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
  } catch (e) { return []; }
}

function classifyAsset(title) {
  const t = title.toLowerCase();
  if (t.includes("台積電") || t.includes("tsmc") || t.includes("2330")) return "TSMC";
  if (t.includes("0050")) return "0050";
  if (t.includes("00679b") || t.includes("美債") || t.includes("債")) return "BOND";
  if (t.includes("tesla") || t.includes("tsla")) return "TSLA";
  if (t.includes("google") || t.includes("alphabet")) return "GOOGL";
  if (t.includes("bitcoin") || t.includes("btc") || t.includes("比特幣")) return "BTC";
  if (t.includes("vt") || t.includes("vwra")) return "GLOBAL";
  if (t.includes("黃金") || t.includes("gold")) return "GOLD";
  return "OTHER";
}

function getExistingLinks_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const values = sheet.getRange(2, 4, Math.min(lastRow - 1, 200), 1).getValues();
  let map = {};
  values.forEach(v => { if (v[0]) map[String(v[0])] = true; });
  return map;
}