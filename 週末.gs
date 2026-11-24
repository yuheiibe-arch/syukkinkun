// === 【連休体制通知】設定項目 ===
const MASTER_SPREADSHEET_ID = '1yj4dG9BzNnUU85C-CdYCwPXougm4XJv3cC2JxwMacK0'; 
const OUTPUT_SPREADSHEET_ID = '1bKbpLbiDZ7V_pDgnUcRApiSA9jEIQ05Q6Msd1AIi3IU'; 
const OUTPUT_SHEET_NAME = '週末勤務者';

// Chatwork投稿用の設定
const CW_ROOM_ID = '325431839'; 
const CW_API_TOKEN = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');

const MASTER_EXCLUDED_SHEET_NAMES = ["祝日", "土日祝業務内容", "参考1", "シフト構成パターン（小谷編集中）", "メンバーTODO", "携帯番号"];
const MASTER_NON_WORK_SET = new Set(["休", "有給", "0", ""]); 

/**
 * [メイン関数] 週末・連休の勤務体制メッセージを生成し、シート書き出し＆Chatwork投稿します。
 * トリガー：毎日実行（連休前の最終平日のみ発火し、通常平日や連休中は自動でスキップされます）
 */
function generateAndPostWeekendDutyMessage() {
  const upcomingHolidays = getUpcomingHolidaysFromCalendar_internal();
  
  if (upcomingHolidays.length === 0) {
    Logger.log('連休前日ではないため、処理をスキップします。');
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

// --- ▼▼▼ 内閣府CSV連動 ＆ 年末年始ロジック ▼▼▼ ---

function getUpcomingHolidaysFromCalendar_internal() {
  const holidayMap = getCaoHolidayMap_common(false);
  const today = new Date();

  // 本日がすでに休日の場合は連休中のためスキップ（連休前の最終平日のみ通知する）
  if (isHoliday_common(today, holidayMap)) {
    return [];
  }

  const holidayDates = [];
  const checkDate = new Date(today);

  // 明日を起点に、連続する休日（土日・内閣府祝日・年末年始）を取得
  for (let i = 0; i < 30; i++) {
    checkDate.setDate(checkDate.getDate() + 1);
    if (!isHoliday_common(checkDate, holidayMap)) {
      break;
    }
    holidayDates.push(new Date(checkDate));
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