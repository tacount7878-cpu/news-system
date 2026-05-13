// 標籤標準化
function normalizeTag_(rawTag) {
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
    return [normalizeTag_(row[0])];
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
  const payload = {
    rowMap: {
      "2026-05-13|TW0050|4447": ["中性", "中", "市場"],
      "2026-05-13|TW0050|4448": ["中性", "中", "市場"],
      "2026-05-13|GOLD|4449": ["正面", "中", "市場"],
      "2026-05-13|TSLA|4450": ["正面", "中", "技術"],
      "2026-05-13|BTC|4451": ["中性", "中", "市場"],
      "2026-05-13|GOOGL|4452": ["正面", "中", "技術"],
      "2026-05-13|GOOGL|4453": ["正面", "中", "技術"],
      "2026-05-13|GOOGL|4454": ["正面", "中", "技術"],
      "2026-05-13|GOOGL|4455": ["正面", "中", "技術"],
      "2026-05-12|TW0050|4456": ["正面", "中", "資金"],
      "2026-05-12|GOLD|4457": ["中性", "低", "雜訊"],
      "2026-05-12|BTC|4458": ["正面", "中", "市場"],
      "2026-05-12|GOLD|4459": ["正面", "中", "財報"],
      "2026-05-12|BOND|4460": ["負面", "中", "政策"],
      "2026-05-13|TSMC|4461": ["正面", "中", "財報"],
      "2026-05-13|BOND|4462": ["負面", "中", "市場"],
      "2026-05-13|BTC|4463": ["正面", "中", "技術"],
      "2026-05-13|TSMC|4464": ["正面", "中", "財報"],
      "2026-05-13|TSMC|4465": ["負面", "中", "競爭"],
      "2026-05-13|TW0050|4466": ["中性", "中", "市場"],
      "2026-05-13|BOND|4467": ["負面", "高", "政策"],
      "2026-05-13|GOOGL|4468": ["正面", "中", "技術"],
      "2026-05-13|GLOBAL|4469": ["負面", "中", "市場"],
      "2026-05-13|GOOGL|4470": ["正面", "中", "技術"],
      "2026-05-13|GOOGL|4471": ["正面", "中", "技術"],
      "2026-05-13|GOOGL|4472": ["正面", "中", "技術"],
      "2026-05-13|GOOGL|4473": ["正面", "中", "技術"],
      "2026-05-12|TSMC|4474": ["正面", "高", "財報"],
      "2026-05-12|BTC|4475": ["正面", "中", "市場"],
      "2026-05-12|GOLD|4476": ["正面", "中", "市場"],
      "2026-05-12|GOLD|4477": ["中性", "中", "市場"],
      "2026-05-12|GOLD|4478": ["正面", "中", "市場"],
      "2026-05-13|GLOBAL|4479": ["負面", "中", "市場"],
      "2026-05-13|GLOBAL|4480": ["負面", "中", "市場"],
      "2026-05-13|GOLD|4481": ["中性", "低", "技術"],
      "2026-05-13|GOLD|4482": ["中性", "中", "政策"],
      "2026-05-13|TSMC|4483": ["負面", "高", "政策"],
      "2026-05-13|GOLD|4484": ["負面", "中", "市場"],
      "2026-05-13|TSMC|4485": ["負面", "中", "市場"],
      "2026-05-13|TSMC|4486": ["負面", "中", "市場"],
      "2026-05-13|TW0050|4487": ["中性", "中", "市場"],
      "2026-05-13|BTC|4488": ["負面", "中", "財報"],
      "2026-05-13|GLOBAL|4489": ["負面", "中", "市場"],
      "2026-05-13|GLOBAL|4490": ["負面", "中", "市場"],
      "2026-05-13|GLOBAL|4491": ["負面", "中", "市場"],
      "2026-05-13|BOND|4492": ["負面", "中", "市場"],
      "2026-05-13|GOLD|4493": ["負面", "中", "市場"],
      "2026-05-13|GOOGL|4494": ["中性", "低", "市場"],
      "2026-05-13|GOOGL|4495": ["正面", "中", "技術"],
      "2026-05-12|GLOBAL|4496": ["負面", "中", "市場"],
      "2026-05-12|GOOGL|4497": ["正面", "中", "技術"],
      "2026-05-12|TSLA|4498": ["中性", "中", "政策"],
      "2026-05-12|GOOGL|4499": ["負面", "中", "技術"],
      "2026-05-13|GLOBAL|4500": ["負面", "中", "市場"],
      "2026-05-13|BOND|4501": ["負面", "中", "政策"],
      "2026-05-13|GOOGL|4502": ["正面", "中", "技術"],
      "2026-05-13|TSMC|4503": ["負面", "高", "競爭"],
      "2026-05-13|TSMC|4504": ["正面", "中", "資金"],
      "2026-05-13|TSMC|4505": ["負面", "中", "市場"],
      "2026-05-13|BOND|4506": ["負面", "中", "政策"],
      "2026-05-13|BOND|4507": ["負面", "高", "政策"],
      "2026-05-13|BOND|4508": ["負面", "高", "市場"],
      "2026-05-13|BOND|4509": ["負面", "中", "市場"],
      "2026-05-13|TSMC|4510": ["負面", "中", "市場"],
      "2026-05-13|GOOGL|4511": ["正面", "低", "技術"],
      "2026-05-13|GOLD|4512": ["中性", "中", "政策"],
      "2026-05-13|GOOGL|4513": ["正面", "中", "技術"],
      "2026-05-13|TSLA|4514": ["中性", "中", "政策"],
      "2026-05-13|TSMC|4515": ["負面", "中", "市場"],
      "2026-05-13|GOLD|4516": ["負面", "中", "市場"],
      "2026-05-13|GOOGL|4517": ["負面", "中", "政策"],
      "2026-05-13|GLOBAL|4518": ["負面", "高", "市場"],
      "2026-05-12|GOLD|4519": ["中性", "低", "雜訊"],
      "2026-05-12|GOLD|4520": ["中性", "低", "雜訊"],
      "2026-05-13|GOLD|4521": ["正面", "中", "市場"],
      "2026-05-13|TW0050|4522": ["正面", "中", "資金"],
      "2026-05-13|GOOGL|4523": ["正面", "中", "技術"],
      "2026-05-13|GOLD|4524": ["中性", "低", "雜訊"],
      "2026-05-13|TSMC|4525": ["正面", "高", "財報"],
      "2026-05-13|GOLD|4526": ["中性", "中", "政策"],
      "2026-05-13|TW0050|4527": ["正面", "中", "市場"],
      "2026-05-13|TW0050|4528": ["中性", "中", "市場"],
      "2026-05-13|TW0050|4529": ["中性", "中", "市場"],
      "2026-05-13|BTC|4530": ["負面", "中", "市場"],
      "2026-05-13|GOOGL|4531": ["正面", "中", "技術"],
      "2026-05-13|GOLD|4532": ["中性", "低", "雜訊"],
      "2026-05-13|BOND|4533": ["負面", "高", "政策"],
      "2026-05-13|BOND|4534": ["負面", "高", "政策"],
      "2026-05-13|BOND|4535": ["負面", "中", "政策"],
      "2026-05-13|BOND|4536": ["負面", "高", "市場"],
      "2026-05-13|BOND|4537": ["負面", "中", "市場"],
      "2026-05-13|TSLA|4538": ["中性", "中", "政策"],
      "2026-05-13|GOOGL|4539": ["正面", "中", "技術"],
      "2026-05-13|GOOGL|4540": ["正面", "中", "技術"],
      "2026-05-13|GOOGL|4541": ["正面", "中", "技術"],
      "2026-05-13|GOLD|4542": ["負面", "中", "市場"],
      "2026-05-13|GOOGL|4543": ["正面", "中", "技術"],
      "2026-05-13|BOND|4544": ["負面", "高", "政策"],
      "2026-05-13|TSMC|4545": ["中性", "中", "市場"],
      "2026-05-13|TSMC|4546": ["負面", "中", "市場"],
      "2026-05-13|GLOBAL|4547": ["負面", "中", "市場"],
      "2026-05-13|TSMC|4548": ["中性", "中", "市場"],
      "2026-05-13|TW0050|4549": ["負面", "中", "市場"],
      "2026-05-13|BTC|4550": ["負面", "高", "市場"],
      "2026-05-13|TSMC|4551": ["負面", "中", "市場"],
      "2026-05-13|GOOGL|4552": ["負面", "中", "技術"],
      "2026-05-13|GOOGL|4553": ["正面", "低", "技術"],
      "2026-05-13|BOND|4554": ["負面", "中", "市場"],
      "2026-05-13|GOLD|4555": ["負面", "高", "市場"],
      "2026-05-13|BOND|4556": ["負面", "中", "市場"],
      "2026-05-13|GOOGL|4557": ["正面", "中", "技術"],
      "2026-05-13|GLOBAL|4558": ["負面", "中", "市場"],
      "2026-05-13|BOND|4559": ["中性", "低", "市場"],
      "2026-05-13|GOLD|4560": ["中性", "低", "雜訊"],
      "2026-05-13|GOLD|4561": ["正面", "中", "市場"],
      "2026-05-13|BTC|4562": ["負面", "中", "市場"],
      "2026-05-13|GLOBAL|4563": ["負面", "中", "市場"],
    },

    iMap: {
      "2026-05-12|TW0050": "ETF吸金與內資買盤支撐台股，資金面仍偏強，但高檔位置不宜追價。",
      "2026-05-12|GOLD": "黃金與白銀消息偏分歧，白銀強勢與礦商獲利支撐貴金屬題材，但部分消費與雜訊新聞需排除。",
      "2026-05-12|BTC": "比特幣資金費率翻正與AI泡沫題材帶動偏多情緒，但仍受8萬美元關卡與宏觀利率變數牽動。",
      "2026-05-12|BOND": "市場對Fed降息預期轉弱，美債與利率環境偏逆風。",
      "2026-05-12|TSMC": "台積電EPS、配息與先進製程擴產皆創高，基本面仍強。",
      "2026-05-12|GLOBAL": "美國通膨數據壓抑科技股與大盤，全球股市短線偏震盪。",
      "2026-05-12|GOOGL": "Google智財協議與AI布局有支撐，但短線也出現服務當機雜音。",
      "2026-05-12|TSLA": "Tesla隨川普訪中帶來政策想像，但變數仍高，短線偏中性。",

      "2026-05-13|TW0050": "台股高檔震盪加劇，ETF資金仍熱，但乖離、泡沫、融資與追高風險升溫。",
      "2026-05-13|GOLD": "黃金受通膨、美元與殖利率壓抑短線回檔，但外銀與年底目標價仍支撐中期偏多格局。",
      "2026-05-13|TSLA": "Tesla Semi、訪中與FSD題材仍有支撐，但漲多回落與AI晶片轉單雜音讓短線偏中性。",
      "2026-05-13|BTC": "美國CPI高於預期壓縮降息預期，比特幣一度跌近8萬美元，短線風險升高。",
      "2026-05-13|GOOGL": "Googlebook、Gemini Intelligence、Android與太空資料中心題材密集發酵，AI產品線動能偏強，但零日漏洞與地圖國安爭議帶來雜音。",
      "2026-05-13|TSMC": "台積電配息、EPS與擴產維持強勢，但費半下跌、ADR回檔、英特爾搶單與地緣封鎖題材壓抑短線評價。",
      "2026-05-13|BOND": "CPI升溫、油價與Fed主席變數推升利率壓力，市場降息預期明顯後退，債券短線偏逆風。",
      "2026-05-13|GLOBAL": "美股受通膨、油價與科技股賣壓拖累，S&P與Nasdaq回落，全球風險偏好降溫。"
    }
  };

  applyAnalysisPayloadByRowMap_(payload);
}
function applyAnalysisPayloadByBlankRows_(payload) {
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

function applyAnalysisPayloadByRowMap_(payload) {
  if (!payload || !payload.rowMap || !payload.iMap) {
    throw new Error("payload 缺少必要欄位：rowMap / iMap");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("news");
  const lastRow = sheet.getLastRow();
  const tz = "Asia/Taipei";

  if (lastRow < 2) {
    throw new Error("news 工作表沒有資料");
  }

  const keys = Object.keys(payload.rowMap);
  if (keys.length === 0) {
    throw new Error("rowMap 是空的");
  }

  const writtenRows = {};
  const iWriteTarget = {};
  let fghWritten = 0;

  keys.forEach(function(key) {
    const parts = key.split("|");
    if (parts.length !== 3) {
      throw new Error("rowMap key 格式錯誤：" + key + "，正確格式為 日期|標籤|行號");
    }

    const dateKey = String(parts[0] || "").trim();
    const tagKey = String(parts[1] || "").trim();
    const rowNumber = Number(parts[2]);

    if (!rowNumber || rowNumber < 2 || rowNumber > lastRow) {
      throw new Error("行號超出範圍：" + key);
    }

    if (writtenRows[rowNumber]) {
      throw new Error("rowMap 重複指定同一列：" + rowNumber);
    }

    const triplet = payload.rowMap[key];
    if (!Array.isArray(triplet) || triplet.length !== 3) {
      throw new Error("rowMap 內容格式錯誤：" + key);
    }

    const row = sheet.getRange(rowNumber, 1, 1, 9).getValues()[0];
    const sheetDate = normalizePayloadDate_(row[0], tz);
    const sheetTag = String(row[1] || "").trim();

    if (sheetDate !== dateKey || sheetTag !== tagKey) {
      throw new Error(
        "key 與 Sheet 實際資料不一致：" +
        key +
        "，Sheet=" +
        sheetDate +
        "|" +
        sheetTag +
        "|" +
        rowNumber
      );
    }

    sheet.getRange(rowNumber, 6, 1, 3).setValues([[
      String(triplet[0] || "").trim(),
      String(triplet[1] || "").trim(),
      String(triplet[2] || "").trim()
    ]]);

    writtenRows[rowNumber] = true;
    fghWritten++;

    if (tagKey !== "OTHER") {
      const iKey = dateKey + "|" + tagKey;
      if (!iWriteTarget[iKey] || rowNumber < iWriteTarget[iKey]) {
        iWriteTarget[iKey] = rowNumber;
      }
    }
  });

  let iWritten = 0;
  const missingIKeys = [];

  Object.keys(payload.iMap).forEach(function(iKey) {
    const parts = iKey.split("|");
    if (parts.length !== 2) {
      throw new Error("iMap key 格式錯誤：" + iKey + "，正確格式為 日期|標籤");
    }

    if (parts[1] === "OTHER") return;

    const rowNumber = iWriteTarget[iKey];
    if (!rowNumber) {
      missingIKeys.push(iKey);
      return;
    }

    sheet.getRange(rowNumber, 9).setValue(String(payload.iMap[iKey] || "").trim());
    iWritten++;
  });

  Logger.log("✅ rowMap 寫入完成");
  Logger.log("FGH：" + fghWritten + " 筆");
  Logger.log("I 欄：" + iWritten + " 筆");
  if (missingIKeys.length) {
    Logger.log("未寫入 iMap，因 rowMap 本次沒有對應列：" + missingIKeys.join(", "));
  }
}
