// === 【休日退勤連絡】設定項目 ===
const LEAVE_MASTER_SPREADSHEET_ID = '1yj4dG9BzNnUU85C-CdYCwPXougm4XJv3cC2JxwMacK0';
const LEAVE_CHATWORK_ROOM_ID = '49055690';
const LEAVE_CHATWORK_TOKEN = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');
const LEAVE_SENT_DATE_PROP_KEY = 'LAST_HOLIDAY_LEAVE_SENT_DATE';

const LEAVE_EXCLUDED_SHEETS = ["祝日", "土日祝業務内容", "参考1", "シフト構成パターン（小谷編集中）", "メンバーTODO", "携帯番号"];
const LEAVE_NON_WORK_STRINGS = ["休", "有給", "0", ""];

/**
 * 【本番トリガー用（15分おき実行推奨）】
 * 土日祝・年末年始において、本日のDS部勤務者全員の退勤時刻を過ぎたら1日1回Chatworkへ通知します。
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

  // 3. 休日ガード（平日はキャッシュ判定のみで即終了）
  if (!isHoliday_common(now)) {
    return;
  }

  // 4. 排他ロック（重複実行防止）
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    // ロック取得後に再確認
    if (props.getProperty(LEAVE_SENT_DATE_PROP_KEY) === todayKey) return;

    if (!LEAVE_CHATWORK_TOKEN || !LEAVE_CHATWORK_ROOM_ID) {
      Logger.log('エラー: CHATWORK_API_TOKEN または ルームID が未設定です。');
      return;
    }

    const dutyList = getTodayShiftsForLeave_internal(now);

    // 休日にもともと勤務者0人（終日不在）の日はスキップ
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

    // 全員の退勤時刻を過ぎたらメッセージ生成＆送信
    const message = buildHolidayLeaveMessage_internal(dutyList);
    const isSuccess = postHolidayLeaveToChatwork_internal(LEAVE_CHATWORK_ROOM_ID, message);

    if (isSuccess) {
      props.setProperty(LEAVE_SENT_DATE_PROP_KEY, todayKey);
      Logger.log(`退勤連絡を送信し、本日(${todayKey})の送信済みフラグを記録しました。`);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * 【テスト確認用（Chatworkへは送信せずログにプレビュー出力）】
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

// --- 以下はヘルパー関数 ---

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

  // 開始時間順（同じなら終了時間順）でソート
  list.sort((a, b) => {
    if (a.startMinutes !== b.startMinutes) return a.startMinutes - b.startMinutes;
    return (a.endMinutes || 0) - (b.endMinutes || 0);
  });

  return list;
}

function postHolidayLeaveToChatwork_internal(roomId, message) {
  const apiUrl = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;
  const options = {
    method: 'post',
    headers: { 'X-ChatWorkToken': LEAVE_CHATWORK_TOKEN },
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
    Logger.log(`Chatwork通信例外: ${e.toString()}`);
    return false;
  }
}