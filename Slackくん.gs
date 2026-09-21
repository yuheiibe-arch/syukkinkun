// === 【Webhook版：開始時間優先】設定項目 ===
const MASTER_SPREADSHEET_ID_SLACK = '1yj4dG9BzNnUU85C-CdYCwPXougm4XJv3cC2JxwMacK0'; 
const OUTPUT_SPREADSHEET_ID_SLACK = '1bKbpLbiDZ7V_pDgnUcRApiSA9jEIQ05Q6Msd1AIi3IU'; 
const MENTION_SHEET_NAME_SLACK = '追加メンション';

// ★修正：スクリカプトプロパティから取得するように変更
const SLACK_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL'); 

// ★追加：Slackの送信先チャンネルID
const SLACK_CHANNEL_ID = 'C0A0RK9M8T1';

const EXCLUDED_SHEET_NAMES_SLACK = ["祝日", "土日祝業務内容", "参考1", "シフト構成パターン（小谷編集中）", "メンバーTODO", "携帯番号"];
const NON_WORK_STRINGS_SLACK = ["休", "有給", "0", ""];

/**
 * 【本番実行用トリガー関数】
 * Webhookを使用してSlackへ通知します。
 * 並び替えロジック：開始時間優先（昇順）
 */
function main_sendTomorrowShiftToSlack() {
  if (!SLACK_WEBHOOK_URL) {
    Logger.log("エラー: スクリプトプロパティに SLACK_WEBHOOK_URL が設定されていません。");
    return;
  }

  // 1. 文面の生成（ソート処理込み）
  const messageBody = generateMessage_Webhook();
  
  if (!messageBody) {
    Logger.log("データ取得エラー、または予期せぬエラーのため中断します。");
    return;
  }

  // 2. Slackへ送信
  Logger.log("メッセージを送信中...");
  postToSlackWebhook_internal(messageBody);
}

// --- Webhook送信関数 ---

function postToSlackWebhook_internal(text) {
  const payload = {
    channel: SLACK_CHANNEL_ID, // ★修正: チャンネルIDを明示的に指定して投稿先を固定
    text: text
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
    const responseText = response.getContentText();
    
    if (responseText === 'ok') {
      Logger.log("Slack送信成功");
    } else {
      Logger.log(`Slack送信エラー: ${responseText}`);
    }
  } catch (e) {
    Logger.log(`Slack通信エラー: ${e.toString()}`);
  }
}

// --- 文面生成＆ソートロジック ---

function generateMessage_Webhook() {
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  // タイトル行
  const dayOfWeekStr = ['日', '月', '火', '水', '木', '金', '土'][tomorrow.getDay()];
  const dateStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy年M月d日');
  const titleLine = `明日（${dateStr} ${dayOfWeekStr}）のDS部 出勤体制`;

  // 勤務者データ取得
  let shiftList = getShiftDataForDate_forSlack(tomorrow);
  
  // Slack用追加メンションの取得（B3セル）
  const slackMention = getSlackMention_internal();
  
  const closingText = "明日もよろしくお願いします。\nお休みなどの連絡は社員はギブランさん、清野さん。アルバイトは島津さんにメンションをしてください。";

  if (shiftList.length === 0) {
    // 誰もいない場合
    return `${titleLine}\n\n明日の勤務者は不在です。\n\n${closingText}\n\n${slackMention}`;
  }

  // === 並び替え処理（開始時間優先） ===
  const listWithValues = shiftList.map(item => {
    const times = parseStartEndVal(item.time);
    return { ...item, startVal: times.start, endVal: times.end };
  });

  listWithValues.sort((a, b) => {
    if (a.startVal !== b.startVal) {
      return a.startVal - b.startVal;
    }
    return a.endVal - b.endVal;
  });

  // === テキスト整形 ===
  const listText = listWithValues.map(item => {
    const formattedTime = formatTimeStr_Slack(item.time);
    return `${item.name}さん：${formattedTime}`;
  }).join("\n");

  // ★修正：コードブロック（```）でリストを囲む
  return `${titleLine}\n\n\`\`\`\n${listText}\n\`\`\`\n\n${closingText}\n\n${slackMention}`;
}

// メンションシートからSlack用メンションを取得する関数
function getSlackMention_internal() {
  try {
    const ss = SpreadsheetApp.openById(OUTPUT_SPREADSHEET_ID_SLACK);
    const sheet = ss.getSheetByName(MENTION_SHEET_NAME_SLACK);
    if (!sheet) return "";
    return sheet.getRange('B3').getDisplayValue().trim();
  } catch (e) {
    Logger.log(`メンション取得エラー: ${e.toString()}`);
    return "";
  }
}

// --- 時間処理ヘルパー関数 ---

function parseStartEndVal(timeStr) {
  if (!timeStr) return { start: 999, end: 0 };
  const parts = timeStr.split('-');
  const s = parseFloat(parts[0].trim());
  const e = parts.length > 1 ? parseFloat(parts[1].trim()) : s;
  return { start: isNaN(s) ? 999 : s, end: isNaN(e) ? 0 : e };
}

function formatTimeStr_Slack(rawTimeStr) {
  if (!rawTimeStr) return "";
  return rawTimeStr.split('-').map(part => {
    const val = parseFloat(part.trim());
    if (isNaN(val)) return part;
    const hour = Math.floor(val);
    const minute = (val - hour) * 60;
    return Utilities.formatString("%02d:%02d", hour, Math.round(minute));
  }).join('-');
}

// --- データ取得ヘルパー関数 ---

function getShiftDataForDate_forSlack(targetDate) {
  const yearShort = targetDate.getFullYear() % 100;
  const month = targetDate.getMonth() + 1;
  const day = targetDate.getDate();
  const yymm = `${yearShort}${('0' + month).slice(-2)}`;

  const spreadsheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID_SLACK);
  const targetSheet = findShiftSheet_forSlack(spreadsheet, yymm);

  if (!targetSheet) {
    Logger.log(`エラー: ${yymm} のシフトシートが見つかりませんでした。`);
    return [];
  }
  return extractNameAndTime_forSlack(targetSheet, day, month);
}

function findShiftSheet_forSlack(spreadsheet, yymm) {
  const allSheets = spreadsheet.getSheets();
  const targetSheetName = `確定_${yymm}`;
  let foundSheet = allSheets.find(sheet => sheet.getName() === targetSheetName);
  if (foundSheet) return foundSheet;
  return allSheets.find(sheet => {
    const sheetName = sheet.getName();
    return !EXCLUDED_SHEET_NAMES_SLACK.includes(sheetName) && sheetName.includes(yymm);
  });
}

// データ開始行の動的検索（A列の「除外」の次）
function findDataStartRow_forSlack(sheet) {
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

function extractNameAndTime_forSlack(sheet, targetDay, targetMonth) {
  const results = [];
  const headerRow = 10;
  
  const dataStartRow = findDataStartRow_forSlack(sheet);
  
  const DATE_START_COLUMN = 9; 
  const DATE_END_COLUMN = 39; 

  const maxColumnToRead = Math.min(sheet.getLastColumn(), DATE_END_COLUMN);
  const headerValues = sheet.getRange(headerRow, 1, 2, maxColumnToRead).getDisplayValues();
  const row10 = headerValues[0];
  const row11 = headerValues[1];

  let nameColumnIndex = -1;
  let dateColumnIndex = -1;

  for (let i = 0; i < row10.length; i++) {
    const cell10 = row10[i].toString().trim();
    const cell11 = row11[i].toString().trim();
    if (cell10 === "氏名" || cell11 === "氏名") {
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
    Logger.log(`エラー: 列が見つかりません (${targetMonth}/${targetDay})`);
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return [];

  const values = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, maxColumnToRead).getDisplayValues();

  for (const rowData of values) {
    // A列にチェック(TRUE)がある行は除外
    if (rowData[0].toString().toUpperCase() === 'TRUE') continue;
    
    const name = rowData[nameColumnIndex]?.trim() || "";
    
    // 集計行に到達したらループを抜ける
    if (name === "時間" || name === "合計" || /^\d+$/.test(name)) {
      break; 
    }
    
    if (!name) continue;

    const shift = rowData[dateColumnIndex]?.trim() || ""; 

    if (shift && !NON_WORK_STRINGS_SLACK.includes(shift) && /\d/.test(shift)) {
       results.push({ name: name, time: shift });
    }
  }
  return results;
}