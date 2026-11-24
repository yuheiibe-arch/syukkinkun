// === 設定項目 (休日朝当番 ＆ 休日退勤連絡) ===
const CHATWORK_TARGET_SPREADSHEET_ID = '1bKbpLbiDZ7V_pDgnUcRApiSA9jEIQ05Q6Msd1AIi3IU';
const CHATWORK_TARGET_SHEET_NAME = 'DS部';
const CHATWORK_FIXED_MENTION_PERSON = "Gibranさん";
const CHATWORK_POST_ROOM_ID = '49055690'; 
const CHATWORK_API_TOKEN = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');

// 休日退勤連絡用の設定
const LEAVE_MASTER_SPREADSHEET_ID = '1yj4dG9BzNnUU85C-CdYCwPXougm4XJv3cC2JxwMacK0';
const LEAVE_SENT_DATE_PROP_KEY = 'LAST_HOLIDAY_LEAVE_SENT_DATE';
const LEAVE_EXCLUDED_SHEETS = ["祝日", "土日祝業務内容", "参考1", "シフト構成パターン（小谷編集中）", "メンバーTODO", "携帯番号"];
const LEAVE_NON_WORK_STRINGS = ["休", "有給", "0", ""];

/**
 * 【休日朝の当番通知】
 * 「DS部」シートの内容に基づき、Chatworkへ当番通知（勤務時間付き）を送信します。
 * ★本日が平日（土日・内閣府祝日・年末年始以外）の場合は自動でスキップします。
 */
function sendDsBuDutyNotificationToChatwork() {
  const today = new Date();
  
  // 1. ブレーキ処理：本日が休日か判定する（内閣府CSV＋土日＋年末年始）
  if (!isHoliday_common(today)) {
    Logger.log("本日は平日のため、通知をスキップします。");
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

// --- 以下は休日朝通知のヘルパー関数です ---

function generateDutyMessageForChatwork_internal(todayDate) {
  const ss = SpreadsheetApp.openById(CHATWORK_TARGET_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CHATWORK_TARGET_SHEET_NAME);
  if (!sheet) return "[info]\n＜ご連絡＞\nエラー: 担当者情報を記載したシートが見つかりませんでした。\n[/info]";

  // 万が一「DS部」シートが当日分に未更新だった場合は自動同期する
  const expectedDateStr = Utilities.formatDate(todayDate, Session.getTimeZone(), 'M/d');
  const currentSheetDate = sheet.getRange('A2').getDisplayValue().trim();
  if (currentSheetDate !== expectedDateStr && typeof updateTodaysWorkList === 'function') {
    Logger.log('「DS部」シートが当日未更新のため、最新シフトを同期します。');
    updateTodaysWorkList();
  }

  let dutyEntries = [];
  let isAbsent = false;
  const data_start_row = 2;

  if (sheet.getLastRow() >= data_start_row) {
    const values = sheet.getRange(data_start_row, 2, sheet.getLastRow() - data_start_row + 1, 2).getDisplayValues();
    if (values.length === 1 && values[0][0].trim() === "不在") {
      isAbsent = true;
    } else {
      dutyEntries = values
        .filter(row => row[0] && row[0].trim() !== "")
        .map(row => {
          const name = row[0].trim();
          const time = (row[1] || "").trim();
          return time ? `${name}：${time}` : name;
        });
    }
  }

  // --- 拠点長情報の取得とテンプレ反映 (当日用) ---
  const leaderInfo = getKyujitsuLeaderInfo_internal(ss, todayDate);
  let leaderSection = "";
  let leaderDetailInfo = "";

  if (leaderInfo && leaderInfo.name) {
    leaderSection = `\n\n本日の休日当番拠点長：${leaderInfo.name}さん\nTel:${leaderInfo.tel}`;
    leaderDetailInfo = getDynamicTemplate_internal(ss);
  }

  let dutyLine, contactLine;
  if (isAbsent || dutyEntries.length === 0) {
    dutyLine = `本日DS部・休日当番は不在です。`;
    contactLine = `急ぎ案件等のチャットは${CHATWORK_FIXED_MENTION_PERSON}にてご連絡いただけますと幸いです！`;
  } else {
    dutyLine = `本日DS部・休日当番は以下の通りです。\n${dutyEntries.join('\n')}`;
    contactLine = `急ぎ案件等のチャットは単独メンション（＋${CHATWORK_FIXED_MENTION_PERSON}）にてご連絡いただけますと幸いです！`;
  }

  return `[info][title]本日の勤務体制[/title]＜ご連絡＞\n${dutyLine}\n${contactLine}${leaderSection}\nよろしくお願いいたします。[/info]${leaderDetailInfo}`;
}

/**
 * 拠点長シートから当日の情報を取得（誤一致防止版）
 */
function getKyujitsuLeaderInfo_internal(ss, targetDate) {
  const sheet = ss.getSheetByName("休日当番拠点長");
  if (!sheet) return null;
  const values = sheet.getDataRange().getDisplayValues();
  
  const y = targetDate.getFullYear();
  const m = targetDate.getMonth() + 1;
  const d = targetDate.getDate();
  
  for (let i = 1; i < values.length; i++) {
    const cellValue = values[i][0]; // A列: 日付
    const nameValue = values[i][1]; // B列: 氏名
    if (!cellValue || !nameValue) continue;
    
    const parsed = parseLeaderSheetDate_internal(cellValue);
    if (parsed && parsed.y === y && parsed.m === m && parsed.d === d) {
      return { name: nameValue.trim(), tel: (values[i][3] || "").trim() };
    }
  }
  return null;
}

function parseLeaderSheetDate_internal(str) {
  const m = str.toString().match(/(\d{4})[年\/\-](\d{1,2})[月\/\-](\d{1,2})/);
  if (!m) return null;
  return { y: parseInt(m[1], 10), m: parseInt(m[2], 10), d: parseInt(m[3], 10) };
}

function getDynamicTemplate_internal(ss) {
  const templateSheet = ss.getSheetByName("テンプレ");
  if (!templateSheet) return "";
  const title = templateSheet.getRange("A1").getDisplayValue().trim();
  const body = templateSheet.getRange("A2").getDisplayValue().trim();
  if (!title && !body) return "";
  return `\n[info][title]${title}[/title]${body}[/info]`;
}

function sendMessageToChatworkAPI_internal(roomId, message) {
  const apiUrl = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;
  const options = {
    method: 'post',
    headers: { 'X-ChatWorkToken': CHATWORK_API_TOKEN }, 
    payload: { body: message },
    muteHttpExceptions: true
  };
  try {
    const res = UrlFetchApp.fetch(apiUrl, options);
    if (res.getResponseCode() === 200) {
      return true;
    }
    Logger.log(`Chatwork送信エラー (${res.getResponseCode()}): ${res.getContentText()}`);
    return false;
  } catch (e) {
    Logger.log(`Chatwork通信エラー: ${e.toString()}`);
    return false;
  }
}

// ============================================================================
// === 【追加機能】土日祝・年末年始の退勤連絡（15分おき固定トリガー監視用） ===
// ============================================================================

/**
 * 【トリガー設定用：15分おき実行】
 * 土日祝・年末年始において、本日のDS部勤務者全員の退勤時刻を過ぎたら1日1回Chatworkへ通知します。
 * （動的トリガーは一切生成しないため、トリガーが増殖することはありません）
 */
function checkAndSendHolidayLeaveNotification() {
  const now = new Date();

  // 1. 早朝ガード（8時前は即終了）
  if (now.getHours() < 8) return;

  // 2. 送信済みガード（本日すでに送信完了していれば即終了）
  const todayKey = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(LEAVE_SENT_DATE_PROP_KEY) === todayKey) {
    return;
  }

  // 3. 休日ガード（平日は即終了）
  if (!isHoliday_common(now)) {
    return;
  }

  // 4. 排他ロック（重複送信防止）
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    if (props.getProperty(LEAVE_SENT_DATE_PROP_KEY) === todayKey) return;

    if (!CHATWORK_API_TOKEN || !CHATWORK_POST_ROOM_ID) {
      Logger.log('エラー: CHATWORK_API_TOKEN または ルームID が未設定です。');
      return;
    }

    const dutyList = getTodayShiftsForLeave_internal(now);

    // 休日にもともと勤務者0人（終日不在）の日は退勤連絡をスキップ
    if (dutyList.length === 0) {
      Logger.log('本日は休日勤務者が0名（終日不在）のため、退勤連絡はスキップします。');
      return;
    }

    const validEndTimes = dutyList
      .filter(item => item.endMinutes !== null)
      .map(item => item.endMinutes);

    if (validEndTimes.length === 0) {
      Logger.log('退勤時刻を判定できるシフトデータがないためスキップします。');
      return;
    }

    const maxEndMinutes = Math.max(...validEndTimes);
    const nowMinutes = now.getHours() * 60 + now.getMinutes();

    if (nowMinutes < maxEndMinutes) {
      const endHHMM = Utilities.formatString('%02d:%02d', Math.floor(maxEndMinutes / 60), maxEndMinutes % 60);
      Logger.log(`最終退勤時刻（${endHHMM}）前のため待機します。`);
      return;
    }

    const message = buildHolidayLeaveMessage_internal(dutyList);
    const isSuccess = sendMessageToChatworkAPI_internal(CHATWORK_POST_ROOM_ID, message);

    if (isSuccess) {
      props.setProperty(LEAVE_SENT_DATE_PROP_KEY, todayKey);
      Logger.log(`退勤連絡を送信し、本日(${todayKey})の送信済みフラグを記録しました。`);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * 【テスト確認用】Chatworkへ送信せず、ログに退勤連絡文面のプレビューを出力します
 */
function test_previewHolidayLeaveNotification() {
  const now = new Date();
  const isHol = isHoliday_common(now);
  const dutyList = getTodayShiftsForLeave_internal(now);

  Logger.log(`本日休日判定: ${isHol}`);
  if (dutyList.length === 0) {
    Logger.log('本日の勤務者: 0名（終日不在のため送信対象外）');
    return;
  }

  const maxEnd = Math.max(...dutyList.filter(d => d.endMinutes !== null).map(d => d.endMinutes));
  const endHHMM = Utilities.formatString('%02d:%02d', Math.floor(maxEnd / 60), maxEnd % 60);
  Logger.log(`判定された最終退勤時刻: ${endHHMM}`);
  Logger.log('▼ 送信メッセージプレビュー ▼\n' + buildHolidayLeaveMessage_internal(dutyList));
}

function buildHolidayLeaveMessage_internal(dutyList) {
  const memberLines = dutyList
    .map(item => `${item.name}：${item.formattedTime}`)
    .join('\n');

  return `[info][title]ドクターサポート部退勤連絡[/title]拠点各位の皆様\n`
    + `＜本日の勤務者＞\n`
    + `${memberLines}\n`
    + `ドクターサポート部の勤務者退勤時刻となります。\n`
    + `以降の緊急連絡についてはギブランさんメンションをお願いします。\n`
    + `業務連絡はメンション不要で構いません。\n`
    + `引き続き何卒よろしくお願い申し上げます。[/info]`;
}

function getTodayShiftsForLeave_internal(targetDate) {
  const yearShort = targetDate.getFullYear() % 100;
  const month = targetDate.getMonth() + 1;
  const day = targetDate.getDate();
  const yymm = `${yearShort}${('0' + month).slice(-2)}`;

  const ss = SpreadsheetApp.openById(LEAVE_MASTER_SPREADSHEET_ID);
  const allSheets = ss.getSheets();
  let sheet = allSheets.find(s => s.getName() === `確定_${yymm}`);
  if (!sheet) {
    sheet = allSheets.find(s => !LEAVE_EXCLUDED_SHEETS.includes(s.getName()) && s.getName().includes(yymm));
  }
  if (!sheet) return [];

  const maxCol = Math.min(sheet.getLastColumn(), 39);
  const headerVals = sheet.getRange(10, 1, 2, maxCol).getDisplayValues();
  let nameCol = -1;
  let dateCol = -1;

  for (let i = 0; i < headerVals[0].length; i++) {
    if (headerVals[0][i].trim() === '氏名' || headerVals[1][i].trim() === '氏名') {
      nameCol = i;
    }
  }
  for (let i = 8; i < maxCol; i++) {
    const parts = headerVals[0][i].trim().split('/');
    if (parts.length === 2 && parseInt(parts[0], 10) === month && parseInt(parts[1], 10) === day) {
      dateCol = i;
      break;
    }
  }
  if (nameCol === -1 || dateCol === -1) return [];

  const lastRow = sheet.getLastRow();
  let dataStartRow = 13;
  if (lastRow >= 1) {
    const colA = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();
    for (let i = 0; i < colA.length; i++) {
      if (colA[i][0].trim() === '除外') {
        dataStartRow = i + 2;
        break;
      }
    }
  }
  if (lastRow < dataStartRow) return [];

  const values = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, maxCol).getDisplayValues();
  const list = [];

  for (const row of values) {
    if (row[0].toString().toUpperCase() === 'TRUE') continue;
    const name = row[nameCol]?.trim() || '';
    if (name === '時間' || name === '合計' || /^\d+$/.test(name)) break;
    if (!name) continue;

    const shift = row[dateCol]?.trim() || '';
    if (shift && !LEAVE_NON_WORK_STRINGS.includes(shift) && /\d/.test(shift)) {
      const parsed = parseShiftTime_common(shift);
      list.push({
        name: name,
        startMinutes: parsed ? parsed.startMinutes : 9999,
        endMinutes: parsed ? parsed.endMinutes : null,
        formattedTime: parsed ? parsed.formattedTime : shift
      });
    }
  }

  list.sort((a, b) => {
    if (a.startMinutes !== b.startMinutes) return a.startMinutes - b.startMinutes;
    return (a.endMinutes || 0) - (b.endMinutes || 0);
  });

  return list;
}