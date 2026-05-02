/**
 * News Fetcher - 重設版
 * 抓「會影響整體資產」的新聞
 * 欄位：A日期 B標籤 C標題 D連結 E發布時間 F情緒 G強度 H類型 I今日判讀
 */

function autoFetchNews() {
  const sheetName = "news";
  const headers = ["\u65e5\u671f","\u6a19\u7c64","\u6a19\u984c","\u9023\u7d50","\u767c\u5e03\u6642\u9593","\u60c5\u7dd2","\u5f37\u5ea6","\u985e\u578b","\u4eca\u65e5\u5224\u8b80"];

  const RSS_SOURCES = [
    // Google News：核心關鍵字（涵蓋台積電、0050、台股、Fed、利率、美債、美股大盤、特斯拉、Google、比特幣、黃金）
    "https://news.google.com/rss/search?q=\u53f0\u7a4d\u96fb+OR+TSMC+OR+0050+OR+\u53f0\u80a1&hl=zh-TW&gl=TW&ceid=TW:zh-Hant",
    "https://news.google.com/rss/search?q=Fed+OR+\u806f\u6e96\u6703+OR+\u5229\u7387+OR+\u7f8e\u50b5&hl=zh-TW&gl=TW&ceid=TW:zh-Hant",
    "https://news.google.com/rss/search?q=Tesla+OR+TSLA+OR+Google+OR+Alphabet&hl=zh-TW&gl=TW&ceid=TW:zh-Hant",
    "https://news.google.com/rss/search?q=Bitcoin+OR+BTC+OR+\u6bd4\u7279\u5e63+OR+Gold+OR+\u9ec3\u91d1&hl=zh-TW&gl=TW&ceid=TW:zh-Hant",

    // 國際財經主流（已驗證可用）
    "https://feeds.reuters.com/reuters/businessNews",
    "https://www.cna.com.tw/rss/aall.xml",
    "https://www.cnbc.com/id/100003114/device/rss/rss.html",
    "https://feeds.marketwatch.com/marketwatch/topstories/",

    // 加密貨幣
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
 * 分類規則：對應 UI 卡片的 7 個資產
 * GLOBAL（全球指數）/ TW0050（台股大盤）/ BOND（美債）
 * TSMC（台積電）/ GOOGL / TSLA / BTC / GOLD
 */
function classifyAsset_(title) {
  const t = title.toLowerCase();

  // 個股優先（避免被大盤蓋掉）
  if (t.includes("\u53f0\u7a4d") || t.includes("tsmc") || t.includes("2330")) return "TSMC";
  if (t.includes("tesla") || t.includes("tsla") || t.includes("\u7279\u65af\u62c9")) return "TSLA";
  if (t.includes("google") || t.includes("alphabet") || t.includes("googl")) return "GOOGL";

  // 加密 / 黃金
  if (t.includes("bitcoin") || t.includes("btc") || t.includes("\u6bd4\u7279\u5e63")) return "BTC";
  if (t.includes("gold") || t.includes("\u9ec3\u91d1") || t.includes("gldm")) return "GOLD";

  // 美債 / Fed / 利率（影響 BOND）
  if (t.includes("fed") || t.includes("\u806f\u6e96\u6703") || t.includes("\u5229\u7387") ||
      t.includes("\u6b96\u5229\u7387") || t.includes("treasury") || t.includes("\u7f8e\u50b5") ||
      t.includes("powell") || t.includes("\u9bee\u91d1\u6708")) return "BOND";

  // 全球指數（S&P / 那斯達克 / 美股大盤）
  if (t.includes("s&p") || t.includes("\u90a3\u65af\u9054\u514b") || t.includes("nasdaq") ||
      t.includes("\u9053\u743c") || t.includes("\u7f8e\u80a1") || t.includes("vt") ||
      t.includes("vwra")) return "GLOBAL";

  // 台股大盤
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