// === 【祝日書き出し ＆ 内閣府CSV共通基盤】 ===
const HOLIDAY_OUTPUT_SPREADSHEET_ID = '1bKbpLbiDZ7V_pDgnUcRApiSA9jEIQ05Q6Msd1AIi3IU';

const CAO_GAIYOU_URL = 'https://www8.cao.go.jp/chosei/shukujitsu/gaiyou.html';
const CAO_DEFAULT_CSV_URL = 'https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv';
const CAO_CACHE_KEY = 'CAO_HOLIDAY_MAP_V1';
const CAO_BACKUP_PROP_KEY = 'CAO_HOLIDAY_BACKUP_V1';
const CAO_CACHE_TTL_SEC = 21600; // 6時間キャッシュ

// 年末年始の休日期間（12/29〜1/3）
const COMMON_YEAR_END_START_DAY = 29;
const COMMON_NEW_YEAR_END_DAY = 3;

/**
 * 内閣府の最新祝日CSVを取得し、「休日当番拠点長」シートの末尾に未登録の土日祝を自動追記します。
 * （既存の入力済み拠点長データがズレないよう、削除せず末尾へ差分追記します）
 */
function writeWeekendAndHolidays() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(HOLIDAY_OUTPUT_SPREADSHEET_ID);
  const targetSheet = ss.getSheetByName('休日当番拠点長');
  if (!targetSheet) {
    Logger.log('エラー: 「休日当番拠点長」シートが見つかりません。');
    return;
  }

  // 内閣府CSVを最新状態で取得
  const holidayMap = getCaoHolidayMap_common(true);
  const holidayDates = Object.keys(holidayMap).sort();
  if (holidayDates.length === 0) {
    Logger.log('エラー: 内閣府の祝日データが取得できなかったため中断します。');
    return;
  }

  // 「祝日」シートが存在する場合は最新データに同期
  syncHolidaySheet_internal(ss, holidayMap, holidayDates);

  // 内閣府CSVで公表されている最大年の12月31日までを対象にする
  const maxYear = Math.max(...holidayDates.map(d => parseInt(d.slice(0, 4), 10)));
  const endDate = new Date(maxYear, 11, 31);

  const startRow = 14;
  const lastRow = targetSheet.getLastRow();

  let writeRow = startRow;
  let startDate = new Date(2026, 0, 12); // シートが空の場合の初期開始日

  // すでに日付が入っている場合は、最終日の「翌日」から末尾に追記
  if (lastRow >= startRow) {
    const existingColA = targetSheet.getRange(startRow, 1, lastRow - startRow + 1, 1).getDisplayValues().flat();
    let lastFilledOffset = -1;
    let lastParsedDate = null;

    for (let i = 0; i < existingColA.length; i++) {
      const parsed = parseJapaneseDateCell_internal(existingColA[i]);
      if (parsed) {
        lastFilledOffset = i;
        lastParsedDate = parsed;
      }
    }

    if (lastParsedDate) {
      writeRow = startRow + lastFilledOffset + 1;
      startDate = new Date(lastParsedDate);
      startDate.setDate(startDate.getDate() + 1);
    }
  }

  if (startDate > endDate) {
    Logger.log(`すでに公表済みの最終年（${maxYear}年末）まで書き出し済みです。追記はありません。`);
    return;
  }

  const results = [];
  const dayNames = ['日', '月', '火', '水', '木', '金', '土'];

  for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    const day = d.getDay();
    const dateStr = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
    const isHoliday = Boolean(holidayMap[dateStr]);
    const isWeekend = day === 0 || day === 6;

    if (isWeekend || isHoliday) {
      const y = d.getFullYear();
      const m = d.getMonth() + 1;
      const da = d.getDate();
      const weekday = dayNames[day];

      const label = isHoliday
        ? `${y}年${m}月${da}日(${weekday}祝)`
        : `${y}年${m}月${da}日(${weekday})`;

      results.push([label]);
    }
  }

  if (results.length > 0) {
    targetSheet.getRange(writeRow, 1, results.length, 1).setValues(results);
    Logger.log(`${results.length}件の土日祝を ${writeRow}行目から追記しました（〜${maxYear}年末）。`);
  }
}

function parseJapaneseDateCell_internal(str) {
  if (!str) return null;
  const m = str.toString().match(/(\d{4})[年\/\-](\d{1,2})[月\/\-](\d{1,2})/);
  if (!m) return null;
  return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
}

function syncHolidaySheet_internal(ss, holidayMap, sortedDates) {
  const holidaySheet = ss.getSheetByName('祝日');
  if (!holidaySheet) return;
  holidaySheet.clearContents();
  const rows = sortedDates.map(d => [d, holidayMap[d]]);
  if (rows.length > 0) {
    holidaySheet.getRange(1, 1, rows.length, 2).setValues(rows);
  }
}

// --- 以下、他スクリプトから呼び出される共通関数（祝日・休日・時間解析） ---

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

    // 内閣府CSV（Shift_JIS）をデコードして解析
    const csvText = response.getContentText('Shift_JIS');
    const rows = Utilities.parseCsv(csvText);

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
      cache.put(CAO_CACHE_KEY, backup, 1800);
      return JSON.parse(backup);
    }
    return {};
  }
}

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

function isNationalHoliday_common(date, holidayMapOpt) {
  const map = holidayMapOpt || getCaoHolidayMap_common(false);
  const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
  return Boolean(map[dateStr]);
}

function isYearEndNewYear_common(date) {
  const m = date.getMonth() + 1;
  const d = date.getDate();
  if (m === 12 && d >= COMMON_YEAR_END_START_DAY && d <= 31) return true;
  if (m === 1 && d >= 1 && d <= COMMON_NEW_YEAR_END_DAY) return true;
  return false;
}

function isHoliday_common(date, holidayMapOpt) {
  const day = date.getDay();
  if (day === 0 || day === 6) return true;
  if (isYearEndNewYear_common(date)) return true;
  return isNationalHoliday_common(date, holidayMapOpt);
}

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