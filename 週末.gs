// === 【連休体制通知】設定項目 ===
const MASTER_SPREADSHEET_ID = '1yj4dG9BzNnUU85C-CdYCwPXougm4XJv3cC2JxwMacK0'; 
const OUTPUT_SPREADSHEET_ID = '1bKbpLbiDZ7V_pDgnUcRApiSA9jEIQ05Q6Msd1AIi3IU'; 
const OUTPUT_SHEET_NAME = '週末勤務者';

// カレンダー（スケジュール表）のID
const CALENDAR_SPREADSHEET_ID = '1WlmirSDOPnIcV2cwY5ClXkMWBFMM-4zrAw4XFDNUPGw';

// Chatwork投稿用の設定
const CW_ROOM_ID = '325431839'; 
const CW_API_TOKEN = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');

const MASTER_EXCLUDED_SHEET_NAMES = ["祝日", "土日祝業務内容", "参考1", "シフト構成パターン（小谷編集中）", "メンバーTODO", "携帯番号"];
const MASTER_NON_WORK_SET = new Set(["休", "有給", "0", ""]); 

// === 年末年始の設定（ハードコード） ===
const YEAR_END_START_MONTH = 12;
const YEAR_END_START_DAY = 29;
const YEAR_END_END_MONTH = 1;
const YEAR_END_END_DAY = 3;

/**
 * [メイン関数] 週末・連休の勤務体制メッセージを生成し、シート書き出し＆Chatwork投稿します。
 * トリガー：毎日実行（明日が平日の場合は自動でスキップされます）
 */
function generateAndPostWeekendDutyMessage() {
  const upcomingHolidays = getUpcomingHolidaysFromCalendar_internal();
  
  if (upcomingHolidays.length === 0) {
    Logger.log('明日は平日（または休日データなし）のため、処理をスキップします。');
    return;
  }
  
  Logger.log(`通知対象の休日が見つかりました（${upcomingHolidays.length}連休）。メッセージの作成を開始します。`);

  const dutyInfo = getDutyInfoForDates_internal(upcomingHolidays);
  const outputSheet = SpreadsheetApp.openById(OUTPUT_SPREADSHEET_ID).getSheetByName(OUTPUT_SHEET_NAME);
  
  if (!outputSheet) {
    Logger.log(`エラー: 書き出し先シート「${OUTPUT_SHEET_NAME}」が見つかりません。`);
    return;
  }
  
  const mention = outputSheet.getRange('C2').getDisplayValue().trim();
  const messageContent = buildMessageContent_internal(dutyInfo, upcomingHolidays.length);

  outputSheet.getRange('B1').setValue(messageContent);
  outputSheet.getRange('B2').setValue(mention);
  Logger.log(`「${OUTPUT_SHEET_NAME}」シートへの書き出しが完了しました。`);

  if (!CW_API_TOKEN || !CW_ROOM_ID) {
    Logger.log('ChatworkのAPIトークンまたはルームIDが未設定のため、投稿はスキップします。');
    return;
  }
  
  const finalChatMessage = mention ? `${mention}\n${messageContent}` : messageContent;
  postToChatwork_internal(CW_ROOM_ID, finalChatMessage);
}

// --- ▼▼▼ カレンダー連動 ＆ 年末年始ロジック ▼▼▼ ---

function isYearEndHoliday_internal(date) {
    const month = date.getMonth() + 1;
    const day = date.getDate();

    // 12月29日～12月31日
    if (month === YEAR_END_START_MONTH && day >= YEAR_END_START_DAY && day <= 31) {
        return true;
    }
    // 1月1日～1月3日
    if (month === YEAR_END_END_MONTH && day >= 1 && day <= YEAR_END_END_DAY) {
        return true;
    }
    return false;
}

function getUpcomingHolidaysFromCalendar_internal() {
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  const year = tomorrow.getFullYear();
  const month = tomorrow.getMonth() + 1;
  const fiscalYear = month <= 3 ? year - 1 : year; 
  const sheetName = `${fiscalYear}年度`;
  
  const ss = SpreadsheetApp.openById(CALENDAR_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  
  const data = sheet.getRange(3, 1, lastRow - 2, 3).getValues();
  const targetDateStr = Utilities.formatDate(tomorrow, 'JST', 'yyyy/MM/dd');
  let startIndex = -1;
  
  for (let i = 0; i < data.length; i++) {
    const cellDate = data[i][0];
    if (!cellDate) continue; 
    
    const cellDateStr = cellDate instanceof Date 
        ? Utilities.formatDate(cellDate, 'JST', 'yyyy/MM/dd') 
        : Utilities.formatDate(new Date(cellDate), 'JST', 'yyyy/MM/dd');

    if (cellDateStr === targetDateStr) {
      startIndex = i;
      break;
    }
  }

  if (startIndex === -1) return []; 

  const holidayDates = [];
  
  // 明日を起点に、連続する休日を取得
  for (let i = startIndex; i < data.length; i++) {
    const dateVal = data[i][0];
    if (!dateVal) continue;
    
    const checkDate = new Date(dateVal);
    const isYearEnd = isYearEndHoliday_internal(checkDate);
    const isCalendarWeekday = data[i][2] === "平日";

    // 「年末年始の期間ではない」かつ「カレンダーが平日」なら連休終了
    if (!isYearEnd && isCalendarWeekday) {
      break; 
    }
    
    holidayDates.push(checkDate);
  }
  
  return holidayDates;
}

// --- その他のヘルパー関数 ---

function buildMessageContent_internal(dutyInfo, holidayLength) {
  const dutyLines = formatDutyLines_internal(dutyInfo);
  
  let closingRemark = 'よろしくお願いいたします(bow)';
  if (holidayLength >= 3) {
    closingRemark = `${holidayLength}連休、よろしくお願いいたします(bow)`;
  } else if (holidayLength === 2) {
    closingRemark = `週末、よろしくお願いいたします(bow)`;
  }
  
  return `お疲れさまです！\n[info][title]週末のDS部体制[/title]\n${dutyLines}[/info]\n${closingRemark}`;
}

function formatDutyLines_internal(dutyInfo) {
  const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
  let bodyLines = [];
  for (const info of dutyInfo) {
    const month = info.date.getMonth() + 1;
    const day = info.date.getDate();
    const dayOfWeek = dayNames[info.date.getDay()];
    let namesPart = info.names.length > 0 ? info.names.join('さん・') + 'さん' : '不在';
    let line = `${month}月${day}日(${dayOfWeek}) ${namesPart}`;
    if (info.names.includes('小谷')) {
      line += ' ※みんなに電話転送OFF';
    }
    bodyLines.push(line);
  }
  return bodyLines.join('\n');
}

function getDutyInfoForDates_internal(dates) {
  const sourceSpreadsheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const dutyInfo = [];
  for (const date of dates) {
    const yearShort = date.getFullYear() % 100;
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const yymm = `${yearShort}${('0' + month).slice(-2)}`;
    const targetSheet = findShiftSheet_internal(sourceSpreadsheet, yymm);
    if (targetSheet) {
      const names = extractNamesForDate_internal(targetSheet, day, month);
      dutyInfo.push({ date: date, names: names });
    } else {
      dutyInfo.push({ date: date, names: ["シフト表なし"] });
    }
  }
  return dutyInfo;
}

function findShiftSheet_internal(spreadsheet, yymm) {
  const allSheets = spreadsheet.getSheets();
  const targetSheetName = `確定_${yymm}`;
  let foundSheet = allSheets.find(sheet => sheet.getName() === targetSheetName);
  if (foundSheet) return foundSheet;
  return allSheets.find(sheet => {
    const sheetName = sheet.getName();
    return !MASTER_EXCLUDED_SHEET_NAMES.includes(sheetName) && sheetName.includes(yymm);
  });
}

// ★追加：データ開始行の動的検索（A列の「除外」の次）
function findDataStartRow_weekend(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 13;
  const colA = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();
  for (let i = 0; i < colA.length; i++) {
    if (colA[i][0].trim() === "除外") {
      return i + 2;
    }
  }
  return 13;
}

function extractNamesForDate_internal(sheet, targetDay, targetMonth) {
  const names = [];
  const headerRow = 10;
  
  // ★修正: 固定の13行目から動的取得に変更
  const dataStartRow = findDataStartRow_weekend(sheet);
  
  const DATE_START_COLUMN = 9; 
  const DATE_END_COLUMN = 39;

  const maxColumnToRead = Math.min(sheet.getLastColumn(), DATE_END_COLUMN);
  const headerValues = sheet.getRange(headerRow, 1, 2, maxColumnToRead).getDisplayValues();
  const row10 = headerValues[0];
  const row11 = headerValues[1];

  let nameColumnIndex = -1;
  let dateColumnIndex = -1;

  for (let i = 0; i < row10.length; i++) {
    if (row10[i].toString().trim() === "氏名" || row11[i].toString().trim() === "氏名") {
      nameColumnIndex = i;
    }
  }

  for (let i = DATE_START_COLUMN - 1; i < DATE_END_COLUMN; i++) {
    if (i >= row10.length) break;
    const cellValue = row10[i].toString().trim();
    const parts = cellValue.split('/');
    if (parts.length === 2 && parseInt(parts[0], 10) === targetMonth && parseInt(parts[1], 10) === targetDay) {
      dateColumnIndex = i;
      break;
    }
  }

  if (nameColumnIndex === -1 || dateColumnIndex === -1) {
    Logger.log('エラー: 「氏名」列または対象日付が見つかりませんでした。');
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return [];
  const values = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, maxColumnToRead).getDisplayValues();

  for (const rowData of values) {
    if (rowData[0].toString().toUpperCase() === 'TRUE') continue; 

    const name = rowData[nameColumnIndex]?.trim() || '';
    
    // ★追加: 集計行（「時間」「合計」や数字）に到達したら読み込み終了
    if (name === "時間" || name === "合計" || /^\d+$/.test(name)) {
      break; 
    }

    const shift = rowData[dateColumnIndex]?.trim() || '';
    
    if (name && shift && !MASTER_NON_WORK_SET.has(shift) && /\d/.test(shift)) {
      names.push(name);
    }
  }
  return names;
}

function postToChatwork_internal(roomId, message) {
  const apiUrl = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;
  const options = {
    method: 'post',
    headers: { 'X-ChatWorkToken': CW_API_TOKEN },
    payload: { body: message },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      Logger.log(`Chatworkへのメッセージ送信に成功しました。ルームID: ${roomId}`);
    } else {
      Logger.log(`Chatwork APIエラー (HTTP ${responseCode}) - ルームID: ${roomId}：${response.getContentText()}`);
    }
  } catch (e) {
    Logger.log(`Chatwork API呼び出し中に例外が発生しました - ルームID: ${roomId}：${e.toString()}`);
  }
}