// === 【統合版】明日のシフト＆不在アラート通知（Chatwork用） ===
const CW_INTEGRATED_MASTER_ID = '1yj4dG9BzNnUU85C-CdYCwPXougm4XJv3cC2JxwMacK0'; 
const CW_INTEGRATED_OUTPUT_ID = '1bKbpLbiDZ7V_pDgnUcRApiSA9jEIQ05Q6Msd1AIi3IU'; 
const CW_INTEGRATED_MENTION_SHEET = '追加メンション'; 

// Chatwork連携設定
const CW_INTEGRATED_ROOM_ID = '325431839'; 
const CW_INTEGRATED_TOKEN = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');

const CW_INTEGRATED_EXCLUDED = ["祝日", "土日祝業務内容", "参考1", "シフト構成パターン（小谷編集中）", "メンバーTODO", "携帯番号"];
const CW_INTEGRATED_NON_WORK = ["休", "有給", "0", ""];

/**
 * [トリガー実行用・本番] 
 * 明日のシフトを取得し、「時間順の出勤者一覧」と「不在時間アラート」を結合してChatworkに通知します。
 */
function sendTomorrowShiftAndAbsence_Chatwork() {
  executeShiftNotification_Integrated(1); // 1日後（明日）を指定して実行
}

/**
 * 【テスト実行用】
 * メッセージ通りの「今日（9月9日）」のデータで抽出・生成し、Chatwork(325431839)に直接投稿します。
 */
function test_sendTodayShiftAndAbsence_Chatwork() {
  Logger.log("テスト投稿を開始します（対象：今日の日付）");
  executeShiftNotification_Integrated(0); // 0日後（今日）を指定して実行
}

// --- メイン処理 ---
function executeShiftNotification_Integrated(addDays) {
  const ssOutput = SpreadsheetApp.openById(CW_INTEGRATED_OUTPUT_ID);
  const mentionSheet = ssOutput.getSheetByName(CW_INTEGRATED_MENTION_SHEET);
  
  if (!mentionSheet) {
    Logger.log("エラー:「追加メンション」シートが見つかりません。");
    return;
  }
  
  const bizHoursStr = mentionSheet.getRange('B1').getDisplayValue().trim() || "07:00-19:00";
  const cwMentionStr = mentionSheet.getRange('B2').getDisplayValue().trim();

  // 日付の設定（本番は+1日、テストは+0日）
  const targetDate = new Date();
  targetDate.setDate(targetDate.getDate() + addDays);
  
  const dateStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'M月d日');
  const dayOfWeekStr = ['日', '月', '火', '水', '木', '金', '土'][targetDate.getDay()];
  
  const rawShiftData = getShiftData_Integrated(targetDate); 

  const shiftBlock = generateShiftListBlock(rawShiftData);
  const absenceBlock = generateAbsenceBlock(rawShiftData, bizHoursStr);

  let finalMessage = cwMentionStr ? `${cwMentionStr}\n` : "";
  
  // 文面のタイトルを「明日」に固定（テストで今日の日付を出す場合でもご提示のテキストに合わせるため）
  finalMessage += `[info][title]明日（${dateStr} ${dayOfWeekStr}）のDS部 出勤体制[/title]`;
  finalMessage += `${shiftBlock}\n`;
  
  if (absenceBlock) {
    finalMessage += `[hr]⚠️ 【不在時間帯のお知らせ】\n営業時間（${bizHoursStr}）において、以下の時間帯で対応者が不在です。\n\n${absenceBlock}`;
  }

  const closingText = "明日もよろしくお願いします。\nお休みなどの連絡は社員はギブランさん、清野さん。アルバイトは島津さんにメンションをしてください。";
  finalMessage += `[hr]${closingText}[/info]`;

  Logger.log("▼ 送信内容プレビュー ▼\n" + finalMessage);
  
  sendMessageToChatwork_Integrated(CW_INTEGRATED_ROOM_ID, finalMessage);
}


// --- メッセージ生成ロジック ---
function generateShiftListBlock(shiftData) {
  if (shiftData.length === 0) return "明日の勤務者は不在です。";

  const listWithValues = shiftData.map(item => {
    const times = parseTimeRange_Integrated(item.time);
    return {
      name: item.name,
      timeText: item.time,
      startVal: times ? times.start : 999,
      endVal: times ? times.end : 0,
      formattedTime: times ? `${formatDecimalTime_Integrated(times.start)}-${formatDecimalTime_Integrated(times.end)}` : item.time
    };
  });

  listWithValues.sort((a, b) => {
    if (a.startVal !== b.startVal) return a.startVal - b.startVal;
    return a.endVal - b.endVal;
  });

  return listWithValues.map(item => `${item.name}さん：${item.formattedTime}`).join("\n");
}

function generateAbsenceBlock(shiftData, bizHoursStr) {
  const bizTimes = parseTimeRange_Integrated(bizHoursStr);
  if (!bizTimes || shiftData.length === 0) return "";

  const workedRanges = [];
  shiftData.forEach(shift => {
    const times = parseTimeRange_Integrated(shift.time);
    if (times) workedRanges.push(times);
  });

  if (workedRanges.length === 0) return "";

  workedRanges.sort((a, b) => a.start - b.start);
  const mergedRanges = [];
  workedRanges.forEach(curr => {
    if (mergedRanges.length === 0) {
      mergedRanges.push(curr); return;
    }
    const prev = mergedRanges[mergedRanges.length - 1];
    if (curr.start <= prev.end) prev.end = Math.max(prev.end, curr.end);
    else mergedRanges.push(curr);
  });

  const absences = [];
  let currentCheckPoint = bizTimes.start;
  mergedRanges.forEach(range => {
    if (currentCheckPoint < range.start) {
      absences.push({ start: currentCheckPoint, end: Math.min(range.start, bizTimes.end) });
    }
    currentCheckPoint = Math.max(currentCheckPoint, range.end);
  });
  if (currentCheckPoint < bizTimes.end) {
    absences.push({ start: currentCheckPoint, end: bizTimes.end });
  }

  if (absences.length === 0) return ""; 

  let blockStr = "";
  absences.forEach(abs => {
    const startStr = formatDecimalTime_Integrated(abs.start);
    const endStr = formatDecimalTime_Integrated(abs.end);
    let nextNames = [];
    let nextTimeStr = "";
    
    if (abs.end < bizTimes.end) {
      nextTimeStr = endStr;
      shiftData.forEach(shift => {
        const times = parseTimeRange_Integrated(shift.time);
        if (times && times.start === abs.end) nextNames.push(shift.name);
      });
    }

    blockStr += `・${startStr}-${endStr}\n`;
    if (nextNames.length > 0) {
      blockStr += ` 次点出勤者：${nextNames.join('、')}さん（${nextTimeStr}）\n\n`;
    } else {
      blockStr += ` 次点出勤者：なし\n\n`;
    }
  });

  return blockStr;
}

// --- 時間・計算ヘルパー関数 ---
function parseTimeRange_Integrated(str) {
  if (!str) return null;
  const parts = str.split('-');
  if (parts.length < 1) return null;
  const start = parseFloat(parts[0]);
  const end = parts.length > 1 ? parseFloat(parts[1]) : start;
  if (isNaN(start)) return null;
  return { start: start, end: isNaN(end) ? start : end };
}

function formatDecimalTime_Integrated(val) {
  const h = Math.floor(val);
  const m = Math.round((val - h) * 60);
  return Utilities.formatString("%02d:%02d", h, m);
}

// --- ★修正: データ取得・送信ヘルパー関数 ---
function getShiftData_Integrated(targetDate) {
  const yearShort = targetDate.getFullYear() % 100;
  const month = targetDate.getMonth() + 1;
  const day = targetDate.getDate();
  const yymm = `${yearShort}${('0' + month).slice(-2)}`;

  const ss = SpreadsheetApp.openById(CW_INTEGRATED_MASTER_ID);
  
  const allSheets = ss.getSheets();
  let sheet = allSheets.find(s => s.getName() === `確定_${yymm}`);
  if (!sheet) {
    sheet = allSheets.find(s => !CW_INTEGRATED_EXCLUDED.includes(s.getName()) && s.getName().includes(yymm));
  }
  if (!sheet) return [];

  const headerVals = sheet.getRange(10, 1, 2, sheet.getMaxColumns()).getDisplayValues();
  let nameCol = -1, dateCol = -1;
  for (let i = 0; i < headerVals[0].length; i++) {
    if (headerVals[0][i].trim() === "氏名" || headerVals[1][i].trim() === "氏名") nameCol = i;
  }
  for (let i = 8; i < 39; i++) { 
    if (i >= headerVals[0].length) break;
    const parts = headerVals[0][i].trim().split('/');
    if (parts.length === 2 && parseInt(parts[0]) === month && parseInt(parts[1]) === day) { 
      dateCol = i; break; 
    }
  }

  if (nameCol === -1 || dateCol === -1) return [];

  const lastRow = sheet.getLastRow();
  
  // ▼▼ 動的取得：A列から「除外」を探す ▼▼
  const colAVals = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();
  let dataStartRow = 13;
  for (let i = 0; i < colAVals.length; i++) {
    if (colAVals[i][0].trim() === "除外") {
      dataStartRow = i + 2; 
      break;
    }
  }
  // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

  if (lastRow < dataStartRow) return [];
  const values = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, sheet.getMaxColumns()).getDisplayValues();
  
  const results = [];
  for (const row of values) {
    if (row[0].toString().toUpperCase() === 'TRUE') continue;
    
    const name = row[nameCol]?.trim() || "";
    
    // ▼▼ 終了判定：氏名に「時間」や数字が来たらストップ ▼▼
    if (name === "時間" || name === "合計" || /^\d+$/.test(name)) {
      break; 
    }
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    if (!name) continue;

    const shift = row[dateCol]?.trim() || "";
    
    if (shift && !CW_INTEGRATED_NON_WORK.includes(shift) && /\d/.test(shift)) {
      results.push({ name: name, time: shift });
    }
  }
  return results;
}

function sendMessageToChatwork_Integrated(roomId, message) {
  if (!CW_INTEGRATED_TOKEN) {
    Logger.log("エラー: Chatwork APIトークンが設定されていません。");
    return;
  }
  
  const apiUrl = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;
  const options = {
    method: 'post',
    headers: { 'X-ChatWorkToken': CW_INTEGRATED_TOKEN },
    payload: { body: message },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    if (response.getResponseCode() === 200) {
      Logger.log("Chatwork送信成功");
    } else {
      Logger.log("Chatwork送信エラー: " + response.getContentText());
    }
  } catch(e) {
    Logger.log("Chatwork通信エラー: " + e.toString());
  }
}