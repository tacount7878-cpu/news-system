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
  const headers = ["日期","標籤","標題","連結","發布時間","情緒","強度","類型","今日判讀"];

  const PER_SOURCE_LIMIT = 5;
  const TOTAL_LIMIT = 50;
  const KEEP_OTHER = false;

  const RSS_SOURCES = [
    // 強綁定來源
    "https://www.federalreserve.gov/feeds/press_all.xml",
    "https://www.coindesk.com/arc/outboundfeeds/rss/",

    // 國際市場情緒
    "https://www.cnbc.com/id/100003114/device/rss/rss.html",
    "https://finance.yahoo.com/news/rssindex",

    // Google News 中文聚合：放最後，補廣度
    "https://news.google.com/rss/search?q=台積電+OR+TSMC+OR+0050+OR+台股&hl=zh-TW&gl=TW&ceid=TW:zh-Hant",
    "https://news.google.com/rss/search?q=Fed+OR+聯準會+OR+利率+OR+美債&hl=zh-TW&gl=TW&ceid=TW:zh-Hant",
    "https://news.google.com/rss/search?q=Tesla+OR+TSLA+OR+Google+OR+Alphabet&hl=zh-TW&gl=TW&ceid=TW:zh-Hant",
    "https://news.google.com/rss/search?q=Bitcoin+OR+BTC+OR+比特幣+OR+Gold+OR+黃金&hl=zh-TW&gl=TW&ceid=TW:zh-Hant"
  ];

  const BLACKLIST = [
    "facebook",
    "cmoney",
    "linetoday",
    "line today",
    "line.me",
    "ptt",
    "youtube",
    "詐騙",
    "車手",
    "車勢",
    "汽車",
    "車貸",
    "貸款",
    "保險",
    "中古車"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const sinceMs = Date.now() - 24 * 60 * 60 * 1000;
  const tz = Session.getScriptTimeZone();

  const existingLinks = getExistingLinks_(sheet);
  const processedTitles = new Set();
  const rows = [];

  for (let s = 0; s < RSS_SOURCES.length; s++) {
    const sourceUrl = RSS_SOURCES[s];
    const items = fetchRssItems_(sourceUrl);

    const candidates = items
      .filter(function(it) {
        return it.pubDate && !isNaN(it.pubDate.getTime());
      })
      .sort(function(a, b) {
        return b.pubDate.getTime() - a.pubDate.getTime();
      })
      .filter(function(it) {
        return it.pubDate.getTime() >= sinceMs;
      });

    let sourceCount = 0;

    for (let i = 0; i < candidates.length; i++) {
      if (sourceCount >= PER_SOURCE_LIMIT) break;
      if (rows.length >= TOTAL_LIMIT) break;

      const it = candidates[i];
      const link = normalizeUrl_(it.link);
      if (!link) continue;

      const lowerLink = link.toLowerCase();
      const lowerTitle = String(it.title || "").toLowerCase();

      if (BLACKLIST.some(function(b) {
        return lowerLink.includes(b) || lowerTitle.includes(b);
      })) continue;
      if (existingLinks[link]) continue;

      const cleanTitle = normalizeTitle_(it.title);
      if (!cleanTitle) continue;
      if (processedTitles.has(cleanTitle)) continue;

      const label = classifyAssetByUrl_(sourceUrl, it.title);

      if (!KEEP_OTHER && label === "OTHER") continue;

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
      sourceCount++;
    }

    if (rows.length >= TOTAL_LIMIT) break;
  }

  if (rows.length) {
    rows.sort(function(a, b) {
      return new Date(b[4]) - new Date(a[4]);
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }

  Logger.log("新增新聞：" + rows.length + " 筆");
}

function normalizeUrl_(url) {
  if (!url) return "";
  try {
    const decoded = decodeURIComponent(url);
    const m = decoded.match(/url=(https?:\/\/[^&]+)/);
    if (m) return m[1].split("?")[0];
  } catch(e) {}
  return String(url).split("?")[0];
}

function normalizeTitle_(title) {
  return (title || "")
    .split(" - ")[0]
    .split("｜")[0]
    .replace(/\s+/g, "")
    .toLowerCase()
    .trim();
}

function fetchRssItems_(url) {
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = resp.getResponseCode();

    if (code !== 200) {
      Logger.log("RSS HTTP 非 200：" + code + " → " + url);
      return [];
    }

    let doc;
    try {
      doc = XmlService.parse(resp.getContentText());
    } catch (e) {
      const preview = resp.getContentText().slice(0, 200);
      Logger.log("RSS XML 解析失敗：" + url + " → " + e.message + "｜preview: " + preview);
      return [];
    }

    const root = doc.getRootElement();
    const rootName = root.getName().toLowerCase();
    let items = [];

    if (rootName === "rss" || rootName === "rdf") {
      items = parseRssItems_(root);
    } else if (rootName === "feed") {
      items = parseAtomItems_(root);
    } else {
      Logger.log("RSS 未知格式：" + rootName + " → " + url);
      return [];
    }

    if (items.length === 0) {
      Logger.log("RSS 0 筆結果：" + url);
    }

    return items;
  } catch (e) {
    Logger.log("RSS 例外：" + url + " → " + e.message);
    return [];
  }
}

function parseRssItems_(root) {
  const channel = root.getChild("channel");
  let rawItems = [];

  if (channel) {
    rawItems = channel.getChildren("item");
  } else {
    // 少數 RDF 格式會直接把 item 放在 root 底下
    rawItems = root.getChildren("item");
  }

  return rawItems.map(function(it) {
    return {
      title: it.getChildText("title") || "",
      link: it.getChildText("link") || "",
      pubDate: parseNewsDate_(it.getChildText("pubDate") || it.getChildText("date") || "")
    };
  });
}

function parseAtomItems_(root) {
  const ns = root.getNamespace();
  const entries = root.getChildren("entry", ns);

  return entries.map(function(entry) {
    const linkElement = getAtomLinkElement_(entry, ns);
    const href = linkElement && linkElement.getAttribute("href")
      ? linkElement.getAttribute("href").getValue()
      : "";

    return {
      title: entry.getChildText("title", ns) || "",
      link: href,
      pubDate: parseNewsDate_(
        entry.getChildText("published", ns) ||
        entry.getChildText("updated", ns) ||
        ""
      )
    };
  });
}

function getAtomLinkElement_(entry, ns) {
  const links = entry.getChildren("link", ns);
  if (!links || links.length === 0) return null;

  for (let i = 0; i < links.length; i++) {
    const relAttr = links[i].getAttribute("rel");
    const rel = relAttr ? relAttr.getValue() : "alternate";
    if (rel === "alternate") return links[i];
  }

  return links[0];
}

function parseNewsDate_(value) {
  if (!value) return null;
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * 來源優先分類：
 * 強綁定來源直接分類；MoneyDJ 為弱綁定，標題分不出來才 fallback 到 TW0050。
 */
function classifyAssetByUrl_(url, title) {
  const u = String(url || "").toLowerCase();

  if (u.includes("federalreserve.gov")) return "BOND";
  if (u.includes("treasury.gov")) return "BOND";
  if (u.includes("kitco.com")) return "GOLD";
  if (u.includes("coindesk.com")) return "BTC";

  const label = classifyAsset_(title);

  if (label === "OTHER" && u.includes("moneydj.com")) return "TW0050";

  return label;
}

/**
 * 投資 UI 使用的 8 類資產：
 * GLOBAL：全球指數 / TW0050：台股與 0050 / BOND：美債與 Fed
 * TSMC / GOOGL / TSLA / BTC / GOLD
 */
function classifyAsset_(title) {
  const t = String(title || "").toLowerCase();

  // 明確排除：非投資訊號
  if (
    t.includes("車貸") ||
    t.includes("房貸") ||
    t.includes("貸款") ||
    t.includes("0利率") ||
    t.includes("零利率") ||
    t.includes("分期") ||
    t.includes("保險") ||
    t.includes("中古車") ||
    t.includes("車勢") ||
    t.includes("詐騙") ||
    t.includes("車手")
  ) {
    return "OTHER";
  }

  // 個股優先
  if (t.includes("台積") || t.includes("tsmc") || t.includes("2330")) return "TSMC";
  if (t.includes("tesla") || t.includes("tsla") || t.includes("特斯拉")) return "TSLA";
  if (t.includes("google") || t.includes("alphabet") || t.includes("googl")) return "GOOGL";

  // BTC：只抓比特幣，不抓其他幣
  if (t.includes("bitcoin") || t.includes("btc") || t.includes("比特幣")) return "BTC";
  if (
    t.includes("xrp") ||
    t.includes("chainlink") ||
    t.includes("solana") ||
    t.includes("ethereum") ||
    t.includes("eth") ||
    t.includes("狗狗幣") ||
    t.includes("迷因幣")
  ) {
    return "OTHER";
  }

  // GOLD：排除地名與詐騙類
  if (t.includes("黃金海岸")) return "OTHER";
  if (t.includes("gold") || t.includes("黃金") || t.includes("gldm")) return "GOLD";

  // BOND：只排除個別配息產品，不擋美債/公債/Treasury
  if (
    (t.includes("配息") || t.includes("月配息") || t.includes("遠期殖利率")) &&
    !t.includes("美債") &&
    !t.includes("公債") &&
    !t.includes("treasury")
  ) {
    return "OTHER";
  }

  if (t.includes("優先收益基金")) return "OTHER";

  if (
    t.includes("fed") ||
    t.includes("聯準會") ||
    t.includes("殖利率") ||
    t.includes("treasury") ||
    t.includes("美債") ||
    t.includes("公債") ||
    t.includes("powell") ||
    t.includes("鮑爾") ||
    (
      t.includes("利率") &&
      !/[毛淨營獲]利率/.test(t) &&
      !t.includes("邊際利率")
    )
  ) {
    return "BOND";
  }

  // 全球指數 / 美股 / VWRA
  if (
    t.includes("s&p") ||
    t.includes("標普") ||
    t.includes("那斯達克") ||
    t.includes("nasdaq") ||
    t.includes("道瓊") ||
    t.includes("美股") ||
    t.includes("vt") ||
    t.includes("vwra")
  ) {
    return "GLOBAL";
  }

  // 台股 / 0050
  if (
    t.includes("0050") ||
    t.includes("台股") ||
    t.includes("加權") ||
    t.includes("台灣加權")
  ) {
    return "TW0050";
  }

  return "OTHER";
}

function getExistingLinks_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const values = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  const map = {};
  values.forEach(function(v) {
    if (v[0]) map[String(v[0])] = true;
  });
  return map;
}

function runTodayPayload() {
  applyAnalysisPayloadByBlankRows({
    fgh: [
      ["中性", "弱", "雜訊"],
      ["負面", "中", "政策"],
      ["正面", "中", "技術"],
      ["負面", "中", "政策"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "資金"],
      ["正面", "中", "技術"],
      ["中性", "中", "政策"],
      ["負面", "中", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "弱", "雜訊"],
      ["正面", "中", "資金"],
      ["中性", "弱", "雜訊"],
      ["正面", "中", "技術"],
      ["正面", "強", "情緒"],
      ["中性", "中", "資金"],
      ["正面", "中", "技術"],
      ["正面", "強", "情緒"],
      ["負面", "中", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "中", "政策"],
      ["負面", "中", "政策"],
      ["負面", "中", "政策"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "技術"],
      ["中性", "弱", "雜訊"],
      ["負面", "強", "情緒"],
      ["正面", "中", "資金"],
      ["中性", "弱", "雜訊"],
      ["正面", "中", "技術"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "雜訊"],
      ["正面", "中", "資金"],
      ["正面", "中", "技術"],
      ["負面", "中", "情緒"],
      ["中性", "弱", "技術"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "技術"],
      ["正面", "中", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "中", "技術"],
      ["中性", "強", "雜訊"],
      ["負面", "強", "情緒"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "資金"],
      ["正面", "中", "技術"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "資金"],
      ["中性", "弱", "雜訊"],
      ["中性", "中", "技術"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "情緒"],
      ["負面", "中", "技術"],
      ["中性", "弱", "雜訊"],
      ["中性", "中", "政策"],
      ["正面", "強", "情緒"],
      ["負面", "中", "政策"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "情緒"],
      ["負面", "中", "政策"],
      ["中性", "弱", "資金"],
      ["中性", "中", "政策"],
      ["負面", "中", "政策"],
      ["正面", "強", "情緒"],
      ["中性", "中", "資金"],
      ["中性", "弱", "雜訊"],
      ["正面", "中", "資金"],
      ["中性", "弱", "雜訊"],
      ["正面", "中", "資金"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "資金"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "技術"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["正面", "中", "技術"],
      ["正面", "強", "技術"],
      ["負面", "中", "情緒"],
      ["正面", "中", "技術"],
      ["中性", "中", "政策"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "雜訊"],
      ["正面", "中", "資金"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "中", "政策"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "弱", "雜訊"],
      ["正面", "中", "資金"],
      ["中性", "弱", "雜訊"],
      ["負面", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "政策"],
      ["正面", "中", "資金"],
      ["負面", "強", "政策"],
      ["中性", "中", "政策"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "中", "政策"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "技術"],
      ["正面", "中", "資金"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["負面", "中", "政策"],
      ["正面", "強", "資金"],
      ["正面", "中", "情緒"],
      ["正面", "中", "政策"],
      ["中性", "中", "政策"],
      ["中性", "弱", "雜訊"],
      ["中性", "中", "政策"],
      ["正面", "強", "資金"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["正面", "中", "情緒"],
      ["中性", "中", "技術"],
      ["正面", "中", "技術"],
      ["負面", "中", "政策"],
      ["負面", "中", "情緒"],
      ["中性", "弱", "技術"],
      ["正面", "中", "技術"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "技術"],
      ["中性", "強", "雜訊"],
      ["負面", "中", "技術"],
      ["正面", "中", "技術"],
      ["正面", "中", "資金"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "資金"],
      ["負面", "強", "政策"],
      ["正面", "中", "技術"],
      ["正面", "中", "情緒"],
      ["負面", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "中", "技術"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["正面", "強", "情緒"],
      ["正面", "強", "技術"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "技術"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "技術"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "情緒"],
      ["負面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "弱", "技術"],
      ["中性", "弱", "技術"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "技術"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "中", "資金"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["正面", "中", "技術"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["負面", "中", "政策"],
      ["中性", "中", "資金"],
      ["正面", "中", "技術"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["負面", "強", "政策"],
      ["正面", "中", "技術"],
      ["中性", "中", "情緒"],
      ["負面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "中", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["正面", "中", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "中", "資金"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "資金"],
      ["正面", "中", "技術"],
      ["中性", "中", "政策"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "技術"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "強", "情緒"],
      ["正面", "強", "技術"],
      ["正面", "中", "資金"],
      ["正面", "中", "技術"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "資金"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["正面", "中", "技術"],
      ["中性", "弱", "資金"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "雜訊"],
      ["正面", "中", "技術"],
      ["正面", "中", "技術"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "技術"],
      ["中性", "中", "政策"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "技術"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "資金"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "資金"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "中", "情緒"],
      ["正面", "中", "資金"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "中", "情緒"],
      ["中性", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["負面", "中", "政策"],
      ["正面", "強", "資金"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "資金"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "資金"],
      ["正面", "強", "技術"],
      ["中性", "中", "情緒"],
      ["負面", "中", "情緒"],
      ["正面", "中", "技術"],
      ["中性", "弱", "雜訊"],
      ["中性", "中", "政策"],
      ["中性", "弱", "雜訊"],
      ["正面", "中", "情緒"],
      ["負面", "中", "政策"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "技術"],
      ["負面", "中", "技術"],
      ["中性", "中", "情緒"],
      ["中性", "中", "技術"],
      ["中性", "中", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["正面", "中", "資金"],
      ["正面", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "技術"],
      ["正面", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "技術"],
      ["正面", "強", "技術"],
      ["正面", "中", "情緒"],
      ["正面", "中", "資金"],
      ["正面", "中", "資金"],
      ["正面", "強", "情緒"],
      ["正面", "強", "技術"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "技術"],
      ["中性", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "情緒"],
      ["負面", "中", "政策"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "技術"],
      ["正面", "中", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "強", "資金"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "技術"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "技術"],
      ["正面", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "資金"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "資金"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "資金"],
      ["正面", "強", "情緒"],
      ["中性", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["負面", "強", "情緒"],
      ["正面", "中", "技術"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "技術"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "資金"],
      ["中性", "中", "政策"],
      ["正面", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "中", "技術"],
      ["正面", "中", "技術"],
      ["中性", "中", "情緒"],
      ["負面", "中", "政策"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "資金"],
      ["正面", "中", "技術"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "技術"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "資金"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "中", "資金"],
      ["中性", "弱", "資金"],
      ["正面", "強", "情緒"],
      ["正面", "中", "資金"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "技術"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "資金"],
      ["中性", "弱", "雜訊"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["中性", "中", "情緒"],
      ["中性", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["正面", "中", "技術"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["正面", "中", "技術"],
      ["正面", "中", "技術"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "資金"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "情緒"],
      ["正面", "中", "技術"],
      ["正面", "強", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "資金"],
      ["中性", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "資金"],
      ["正面", "中", "技術"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "資金"],
      ["中性", "中", "政策"],
      ["正面", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["負面", "中", "政策"],
      ["中性", "中", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "資金"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "資金"],
      ["正面", "強", "情緒"],
      ["中性", "中", "情緒"],
      ["正面", "中", "資金"],
      ["正面", "強", "情緒"],
      ["中性", "中", "情緒"],
      ["正面", "中", "技術"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "中", "政策"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["負面", "中", "政策"],
      ["中性", "弱", "資金"],
      ["中性", "中", "情緒"],
      ["正面", "中", "技術"],
      ["中性", "中", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "弱", "技術"],
      ["正面", "中", "技術"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "強", "資金"],
      ["正面", "強", "資金"],
      ["負面", "中", "政策"],
      ["中性", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "技術"],
      ["正面", "強", "情緒"],
      ["正面", "強", "技術"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "雜訊"],
      ["中性", "中", "政策"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "資金"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "技術"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "情緒"],
      ["負面", "中", "政策"],
      ["中性", "中", "政策"],
      ["中性", "弱", "雜訊"],
      ["中性", "中", "政策"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["中性", "中", "情緒"],
      ["正面", "中", "資金"],
      ["中性", "弱", "雜訊"],
      ["正面", "中", "技術"],
      ["中性", "弱", "雜訊"],
      ["中性", "弱", "雜訊"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "中", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "中", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "資金"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["負面", "中", "情緒"],
      ["中性", "弱", "資金"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "弱", "資金"],
      ["中性", "中", "政策"],
      ["正面", "中", "資金"],
      ["中性", "弱", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["中性", "中", "政策"],
      ["正面", "強", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "弱", "情緒"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["中性", "中", "政策"],
      ["正面", "中", "技術"],
      ["負面", "中", "情緒"],
      ["負面", "中", "情緒"],
      ["中性", "中", "情緒"],
      ["負面", "強", "情緒"],
      ["負面", "中", "政策"],
      ["正面", "中", "情緒"],
      ["正面", "中", "技術"],
      ["中性", "弱", "技術"],
      ["正面", "中", "資金"],
      ["中性", "中", "資金"],
      ["負面", "中", "情緒"],
      ["正面", "強", "情緒"],
      ["正面", "中", "技術"],
      ["正面", "中", "資金"],
      ["中性", "弱", "情緒"],
      ["正面", "中", "技術"],
      ["中性", "中", "情緒"],
      ["負面", "中", "技術"],
      ["正面", "中", "技術"],
      ["負面", "中", "政策"],
      ["正面", "中", "情緒"],
      ["正面", "中", "資金"],
      ["負面", "強", "情緒"],
      ["正面", "中", "情緒"],
      ["負面", "中", "情緒"]
    ],
    iMap: {
      "2026-05-03|BOND": "Fed人事與信用風險雜音升高 利率路徑仍偏觀望",
      "2026-05-03|GOOGL": "Alphabet雲端與AI動能延續 但新聞雜訊偏多",
      "2026-05-03|TSMC": "台積電基本面與AI封裝題材續強 短線仍看資金動能",
      "2026-05-03|TW0050": "連假前台股轉保守 風險意識高於追價動能",
      "2026-05-03|GLOBAL": "美股週前觀望氣氛升高 聚焦AI支出與信用風險",
      "2026-05-04|TW0050": "台股突破4萬點情緒強烈 但融資升溫與過熱風險同步上升",
      "2026-05-04|GOOGL": "Google AI與雲端題材延續 短線評價與產能壓力需觀察",
      "2026-05-04|TSLA": "Tesla低價與進口題材偏中性 需求結構仍需驗證",
      "2026-05-04|TSMC": "台積電領軍台股創高 龍潭擴建與先進製程題材強勢",
      "2026-05-04|BOND": "Fed人事、通膨與美債殖利率交錯 債券訊號偏震盪",
      "2026-05-04|BTC": "比特幣站上8萬美元附近 ETF流入支撐但地緣消息放大波動",
      "2026-05-04|GOLD": "黃金受通膨、美元與避險需求拉扯 訊號偏震盪整理",
      "2026-05-04|GLOBAL": "美股財報與AI支出仍支撐市場 但荷姆茲與油價風險壓抑情緒"
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
