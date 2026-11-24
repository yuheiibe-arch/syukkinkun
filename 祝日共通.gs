// === 【共通基盤】内閣府 祝日CSV自動取得 ＆ 休日・時間判定 ===
const CAO_GAIYOU_URL = 'https://www8.cao.go.jp/chosei/shukujitsu/gaiyou.html';
const CAO_DEFAULT_CSV_URL = 'https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv';
const CAO_CACHE_KEY = 'CAO_HOLIDAY_MAP_V1';
const CAO_BACKUP_PROP_KEY = 'CAO_HOLIDAY_BACKUP_V1';
const CAO_CACHE_TTL_SEC = 21600; // 6時間キャッシュ

// 年末年始の休日期間（12/29〜1/3）
const COMMON_YEAR_END_START_DAY = 29;
const COMMON_NEW_YEAR_END_DAY = 3;

/**
 * 内閣府の公式CSVから祝日マップ { 'yyyy/MM/dd': '祝日名' } を取得します。
 * @param {boolean} forceRefresh trueの場合はキャッシュを無視して再取得
 */
function getCaoHolidayMap_common(forceRefresh) {
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    const cached = cache.get(CAO_CACHE_KEY);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        Logger.log(`祝日キャッシュ解析エラー: ${e.toString()}`);
      }
    }
  }

  try {
    const csvUrl = resolveCaoCsvUrl_common();
    const response = UrlFetchApp.fetch(csvUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      throw new Error(`HTTP ${response.getResponseCode()}`);
    }

    // 内閣府CSVは Shift_JIS のため文字コードを指定してデコード
    const csvText = response.getContentText('Shift_JIS');
    const rows = Utilities.parseCsv(csvText);

    // 容量節約のため前年以降の祝日データを保持（過去数十年分を除外）
    const minYear = new Date().getFullYear() - 1;
    const holidayMap = {};

    for (let i = 1; i < rows.length; i++) {
      if (!rows[i] || rows[i].length < 2) continue;
      const rawDate = rows[i][0].trim();
      const holidayName = rows[i][1].trim();

      const match = rawDate.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
      if (!match) continue;

      const year = parseInt(match[1], 10);
      if (year < minYear) continue;

      const mm = ('0' + parseInt(match[2], 10)).slice(-2);
      const dd = ('0' + parseInt(match[3], 10)).slice(-2);
      holidayMap[`${year}/${mm}/${dd}`] = holidayName;
    }

    if (Object.keys(holidayMap).length === 0) {
      throw new Error('CSVから有効な祝日データが取得できませんでした。');
    }

    const jsonStr = JSON.stringify(holidayMap);
    cache.put(CAO_CACHE_KEY, jsonStr, CAO_CACHE_TTL_SEC);
    PropertiesService.getScriptProperties().setProperty(CAO_BACKUP_PROP_KEY, jsonStr);
    return holidayMap;

  } catch (e) {
    Logger.log(`内閣府CSV取得エラー（バックアップを使用します）: ${e.toString()}`);
    const backup = PropertiesService.getScriptProperties().getProperty(CAO_BACKUP_PROP_KEY);
    if (backup) {
      cache.put(CAO_CACHE_KEY, backup, 1800); // 30分後に再試行
      return JSON.parse(backup);
    }
    return {};
  }
}

/**
 * gaiyou.html 内のCSVリンクを動的に解決（見つからない場合は標準URLを使用）
 */
function resolveCaoCsvUrl_common() {
  try {
    const res = UrlFetchApp.fetch(CAO_GAIYOU_URL, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const html = res.getContentText('UTF-8');
      const match = html.match(/href=["']([^"']*syukujitsu[^"']*\.csv)["']/i);
      if (match && match[1]) {
        const href = match[1];
        if (/^https?:\/\//i.test(href)) return href;
        if (href.startsWith('/')) return `https://www8.cao.go.jp${href}`;
        return `https://www8.cao.go.jp/chosei/shukujitsu/${href.replace(/^\.\//, '')}`;
      }
    }
  } catch (e) {
    Logger.log(`gaiyou.html 参照スキップ: ${e.toString()}`);
  }
  return CAO_DEFAULT_CSV_URL;
}

/**
 * 内閣府CSV上の「国民の祝日・休日」かどうかを判定
 */
function isNationalHoliday_common(date, holidayMapOpt) {
  const map = holidayMapOpt || getCaoHolidayMap_common(false);
  const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
  return Boolean(map[dateStr]);
}

/**
 * 年末年始（12/29〜1/3）かどうかを判定
 */
function isYearEndNewYear_common(date) {
  const m = date.getMonth() + 1;
  const d = date.getDate();
  if (m === 12 && d >= COMMON_YEAR_END_START_DAY && d <= 31) return true;
  if (m === 1 && d >= 1 && d <= COMMON_NEW_YEAR_END_DAY) return true;
  return false;
}

/**
 * パターンB共通の休日判定（土日 または 内閣府の祝日・休日 または 年末年始）
 */
function isHoliday_common(date, holidayMapOpt) {
  const day = date.getDay();
  if (day === 0 || day === 6) return true;
  if (isYearEndNewYear_common(date)) return true;
  return isNationalHoliday_common(date, holidayMapOpt);
}

/**
 * シフト時間文字列（例: "8-17", "9.5-18", "08:30-17:30"）を解析して分単位とHH:mm表記を返す
 */
function parseShiftTime_common(rawTimeStr) {
  if (!rawTimeStr) return null;
  const match = rawTimeStr.toString().trim().match(/(\d{1,2}(?::\d{2}|\.\d+)?)\s*[-ー〜～]\s*(\d{1,2}(?::\d{2}|\.\d+)?)/);
  if (!match) return null;

  const toMinutes = (token) => {
    if (token.includes(':')) {
      const [h, m] = token.split(':').map(Number);
      return h * 60 + m;
    }
    return Math.round(parseFloat(token) * 60);
  };

  const toHHMM = (mins) => {
    const h = Math.floor(mins / 60);
    const m = mins % 60;
    return Utilities.formatString('%02d:%02d', h, m);
  };

  const startMinutes = toMinutes(match[1]);
  const endMinutes = toMinutes(match[2]);
  if (isNaN(startMinutes) || isNaN(endMinutes)) return null;

  return {
    startMinutes: startMinutes,
    endMinutes: endMinutes,
    formattedTime: `${toHHMM(startMinutes)}-${toHHMM(endMinutes)}`
  };
}