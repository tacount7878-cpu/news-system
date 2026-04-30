/**
 * Google Apps Script - News fetcher (Clean Pipe) V1.1
 * - Fetches Google News RSS
 * - Filters last 24 hours
 * - Dedup by link
 * - Output: 5 columns only
 */
function autoFetchNews() {
  const sheetName = "news";
  const headers = ["日期", "標籤", "標題", "連結", "發布時間"];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  // Header
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const now = new Date();
  const sinceMs = now.getTime() - 24 * 60 * 60 * 1000;

  const query = "0050 OR 元大台灣50 OR 台灣50 OR 台股ETF OR 台積電 OR TSMC OR 台積 OR 2330 OR 00679B OR 元大美債20年 OR 00719B OR 元大美債1-3年 OR VT OR VWRA OR 全球ETF OR Tesla OR 特斯拉 OR TSLA OR 馬斯克 OR Google OR Alphabet OR 谷歌 OR IBKR OR 盈透證券 OR GLDM OR SPDR Gold MiniShares Trust OR SPDR Gold MiniShares OR SPDR Gold OR Bitcoin OR BTC OR 比特幣 OR 加密貨幣";
  const rssUrl =
    "https://news.google.com/rss/search?q=" +
    encodeURIComponent(query) +
    "&hl=zh-TW&gl=TW&ceid=TW:zh-Hant";

  const items = fetchRssItems_(rssUrl);
  if (!items.length) return;

  const existingLinks = getExistingLinks_(sheet);
  const tz = Session.getScriptTimeZone();
  let rows = [];

  for (let i = 0; i < items.length; i++) {
    let it = items[i];

    if (!it.pubDate || isNaN(it.pubDate.getTime())) continue;
    if (it.pubDate.getTime() < sinceMs) continue;

    let link = normalizeUrl_(it.link);
    if (!link || existingLinks[link]) continue;

    let label = classifyLabel_(it.title);
    let dateStr = Utilities.formatDate(it.pubDate, tz, "yyyy-MM-dd");
    let pubStr = Utilities.formatDate(it.pubDate, tz, "yyyy-MM-dd HH:mm:ss");

    rows.push([dateStr, label, it.title || "", link, pubStr]);
    existingLinks[link] = true;

    if (rows.length >= 20) break;
  }

  if (rows.length) {
    rows.sort((a, b) => new Date(b[4]) - new Date(a[4]));
    sheet
      .getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length)
      .setValues(rows);
  }
}

function normalizeUrl_(url) {
  return url ? url.split("?")[0] : "";
}

function fetchRssItems_(url) {
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return [];

    const doc = XmlService.parse(resp.getContentText());
    const items = doc.getRootElement().getChild("channel").getChildren("item");

    return items.map(it => ({
      title: it.getChildText("title"),
      link: it.getChildText("link"),
      pubDate: new Date(it.getChildText("pubDate"))
    }));
  } catch (e) {
    return [];
  }
}

function classifyLabel_(title) {
  const t = (title || "").toUpperCase();

  if (t.includes("TSMC") || t.includes("台積")) return "TSMC";
  if (t.includes("0050")) return "0050";
  if (t.includes("TESLA") || t.includes("TSLA")) return "TSLA";
  if (t.includes("BITCOIN") || t.includes("BTC") || t.includes("比特幣")) return "BTC";

  return "OTHER";
}

function getExistingLinks_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const values = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  let map = {};

  values.forEach(v => {
    if (v[0]) map[String(v[0])] = true;
  });

  return map;
}
