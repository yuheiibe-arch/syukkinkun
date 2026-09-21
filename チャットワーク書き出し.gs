// === 設定項目 (スクリプト2) ===
const CHATWORK_TARGET_SPREADSHEET_ID = '1bKbpLbiDZ7V_pDgnUcRApiSA9jEIQ05Q6Msd1AIi3IU';
const CHATWORK_TARGET_SHEET_NAME = 'DS部';
const CHATWORK_FIXED_MENTION_PERSON = "Gibranさん";
const CHATWORK_POST_ROOM_ID = '49055690'; 
const CHATWORK_API_TOKEN = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');

// ★追加：カレンダーシートと年末年始設定
const DAILY_CALENDAR_ID = '1WlmirSDOPnIcV2cwY5ClXkMWBFMM-4zrAw4XFDNUPGw';
const DAILY_YEAR_END_START_MONTH = 12;
const DAILY_YEAR_END_START_DAY = 29;
const DAILY_YEAR_END_END_MONTH = 1;
const DAILY_YEAR_END_END_DAY = 3;

/**
 * 「DS部」シートの内容に基づき、Chatworkへ当番通知を送信します。
 * ★本日が平日（カレンダーで休日判定されていない日）の場合は自動でスキップします。
 */
function sendDsBuDutyNotificationToChatwork() {
  const today = new Date();
  
  // 1. ブレーキ処理：本日が休日か判定する
  if (!isTodayHoliday_internal(today)) {
    Logger.log("本日は平日（カレンダーの休日指定なし）のため、通知をスキップします。");
    return;
  }

  if (!CHATWORK_API_TOKEN || !CHATWORK_POST_ROOM_ID) {
    Logger.log('エラー: Chatwork APIトークンまたはルームIDが未設定です。');
    return;
  }
  
  const messageContent = generateDutyMessageForChatwork_internal(today);
  if (messageContent) {
    sendMessageToChatworkAPI_internal(CHATWORK_POST_ROOM_ID, messageContent);
  }
}

// --- 以下は sendDsBuDutyNotificationToChatwork のためのヘルパー関数です ---

function generateDutyMessageForChatwork_internal(todayDate) {
  const ss = SpreadsheetApp.openById(CHATWORK_TARGET_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CHATWORK_TARGET_SHEET_NAME);
  if (!sheet) return "[info]\n＜ご連絡＞\nエラー: 担当者情報を記載したシートが見つかりませんでした。\n[/info]";

  let namesOnDuty = [], isAbsent = false, name_column = 2, data_start_row = 2;
  if (sheet.getLastRow() >= data_start_row) {
    const values = sheet.getRange(data_start_row, name_column, sheet.getLastRow() - data_start_row + 1, 1).getDisplayValues();
    if (values.length === 1 && values[0][0] === "不在") isAbsent = true;
    else namesOnDuty = values.map(row => row[0].trim()).filter(name => name);
  }

  // --- 拠点長情報の取得とテンプレ反映 (当日用) ---
  const leaderInfo = getKyujitsuLeaderInfo_internal(ss, todayDate);
  let leaderSection = "";
  let leaderDetailInfo = "";

  // 拠点長名(name)が取得できている場合のみ追加
  if (leaderInfo && leaderInfo.name) {
    leaderSection = `\n\n本日の休日当番拠点長：${leaderInfo.name}さん\nTel:${leaderInfo.tel}`;
    leaderDetailInfo = getDynamicTemplate_internal(ss);
  }
  // ------------------------------------

  let dutyLine, contactLine;
  if (isAbsent || namesOnDuty.length === 0) {
    dutyLine = `本日DS部・休日当番は不在です。`;
    contactLine = `急ぎ案件等のチャットは${CHATWORK_FIXED_MENTION_PERSON}にてご連絡いただけますと幸いです！`;
  } else {
    dutyLine = `本日DS部・休日当番は${namesOnDuty.join('、')}です。`;
    contactLine = `急ぎ案件等のチャットは単独メンション（＋${CHATWORK_FIXED_MENTION_PERSON}）にてご連絡いただけますと幸いです！`;
  }

  // leaderSection と leaderDetailInfo が空なら、従来通りの表示になります
  return `[info][title]本日の勤務体制[/title]＜ご連絡＞\n${dutyLine}\n${contactLine}${leaderSection}\nよろしくお願いいたします。[/info]${leaderDetailInfo}`;
}

/**
 * 本日がカレンダー上で休日、または年末年始かどうかを判定する関数
 */
function isTodayHoliday_internal(targetDate) {
  // ① 年末年始の判定
  const month = targetDate.getMonth() + 1;
  const day = targetDate.getDate();
  if (month === DAILY_YEAR_END_START_MONTH && day >= DAILY_YEAR_END_START_DAY && day <= 31) return true;
  if (month === DAILY_YEAR_END_END_MONTH && day >= 1 && day <= DAILY_YEAR_END_END_DAY) return true;

  // ② カレンダーシートの判定
  const year = targetDate.getFullYear();
  const fiscalYear = month <= 3 ? year - 1 : year; // 1〜3月は前年度シート
  const sheetName = `${fiscalYear}年度`;
  
  try {
    const ss = SpreadsheetApp.openById(DAILY_CALENDAR_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return false; 

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return false;
    
    const data = sheet.getRange(3, 1, lastRow - 2, 3).getValues();
    const targetDateStr = Utilities.formatDate(targetDate, 'JST', 'yyyy/MM/dd');
    
    for (let i = 0; i < data.length; i++) {
      const cellDate = data[i][0];
      if (!cellDate) continue; 
      
      const cellDateStr = cellDate instanceof Date 
          ? Utilities.formatDate(cellDate, 'JST', 'yyyy/MM/dd') 
          : Utilities.formatDate(new Date(cellDate), 'JST', 'yyyy/MM/dd');

      if (cellDateStr === targetDateStr) {
        // C列(インデックス2)が「平日」以外なら休日（true）として扱う
        return data[i][2] !== "平日";
      }
    }
  } catch(e) {
    Logger.log(`カレンダー判定エラー: ${e.toString()}`);
  }
  
  return false;
}

/**
 * 拠点長シートから当日の情報を取得
 */
function getKyujitsuLeaderInfo_internal(ss, targetDate) {
  const sheet = ss.getSheetByName("休日当番拠点長");
  if (!sheet) return null;
  const values = sheet.getDataRange().getDisplayValues();
  
  const y = targetDate.getFullYear();
  const m = targetDate.getMonth() + 1;
  const d = targetDate.getDate();
  const p1 = `${y}/${m}/${d}`;
  const p2 = `${y}/${('0'+m).slice(-2)}/${('0'+d).slice(-2)}`;
  const p3 = `${y}年${m}月${d}日`;
  
  for (let i = 1; i < values.length; i++) {
    const cellValue = values[i][0]; // A列: 日付
    const nameValue = values[i][1]; // B列: 氏名
    
    // 日付セルが空、または氏名が空の場合はスキップ
    if (!cellValue || !nameValue) continue;
    
    // 日付が一致するか判定
    if (cellValue.includes(p1) || cellValue.includes(p2) || cellValue.includes(p3)) {
      return { name: nameValue.trim(), tel: values[i][3].trim() };
    }
  }
  return null;
}

/**
 * テンプレシートから内容を読み込み[info][title]を付与して返す
 */
function getDynamicTemplate_internal(ss) {
  const templateSheet = ss.getSheetByName("テンプレ");
  if (!templateSheet) return "";
  const title = templateSheet.getRange("A1").getDisplayValue().trim();
  const body = templateSheet.getRange("A2").getDisplayValue().trim();
  // タイトルも本文も空なら何も返さない
  if (!title && !body) return "";
  return `\n[info][title]${title}[/title]${body}[/info]`;
}

function sendMessageToChatworkAPI_internal(roomId, message) {
  const apiUrl = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;
  const options = {
    method: 'post',
    // 修正: CW_API_TOKEN を設定項目にある CHATWORK_API_TOKEN に統一しました
    headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN }, 
    payload: { body: message },
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(apiUrl, options);
}