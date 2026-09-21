// === 設定項目 ===
const SOURCE_SPREADSHEET_ID = '1yj4dG9BzNnUU85C-CdYCwPXougm4XJv3cC2JxwMacK0'; 
const DESTINATION_SPREADSHEET_ID_FOR_SHIFT = '1bKbpLbiDZ7V_pDgnUcRApiSA9jEIQ05Q6Msd1AIi3IU'; 
const DEST_SHEET_NAME_FOR_SHIFT = 'DS部';
const EXCLUDED_SHEET_NAMES = ["祝日", "土日祝業務内容", "参考1", "シフト構成パターン（小谷編集中）", "メンバーTODO", "携帯番号"];
const NON_WORK_STRINGS = ["休", "有給"];

/**
 * シフト表から今日の勤務者を抽出し、「DS部」シートを更新します。
 */
function updateTodaysWorkList() {
  const today = new Date();
  const currentYearShort = today.getFullYear() % 100;
  const currentMonth = today.getMonth() + 1;
  const currentDay = today.getDate();
  const currentYYMM = `${currentYearShort}${('0' + currentMonth).slice(-2)}`;
  const dateStringToWrite = Utilities.formatDate(today, Session.getTimeZone(), 'M/d');
  
  const targetSheet = findTargetSheet_internal(currentYYMM);

  if (!targetSheet) {
    Logger.log(`対象月(${currentYYMM})のシフトシートが見つかりませんでした。`);
    return;
  }
  Logger.log(`シート「${targetSheet.getName()}」を処理します。`);
    
  const workingNames = extractWorkingNames_internal(targetSheet, currentDay, currentMonth);
  writeToDestinationSheet_internal(workingNames, dateStringToWrite);
  Logger.log(`「${DEST_SHEET_NAME_FOR_SHIFT}」シートへの書き込みが完了しました。`);
}

// --- 以下はヘルパー関数です ---

function findTargetSheet_internal(currentYYMM) {
  const sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  const allSheets = sourceSpreadsheet.getSheets();
  const targetSheetNameForCurrentMonth = `確定_${currentYYMM}`;
  
  let foundSheet = allSheets.find(sheet => sheet.getName() === targetSheetNameForCurrentMonth);
  if (foundSheet) return foundSheet;

  return allSheets.find(sheet => {
    const sheetName = sheet.getName();
    return !EXCLUDED_SHEET_NAMES.includes(sheetName) && sheetName.includes(currentYYMM);
  });
}

/**
 * [追加] A列から「除外」を探し、その次の行番号を返す関数（動的スタート）
 */
function findDataStartRow_internal(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 13; 
  
  const colA = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();
  for (let i = 0; i < colA.length; i++) {
    if (colA[i][0].trim() === "除外") {
      return i + 2; // 「除外」の次の行
    }
  }
  return 13; // 見つからない場合のデフォルト
}

/**
 * [修正版] シフトシートから勤務者を抽出します。
 * ★開始行を動的取得に変更
 * ★集計行（氏名が「時間」や数字）に到達したら読み込みを終了する処理を追加
 */
function extractWorkingNames_internal(sheet, currentDay, currentMonth) {
  const names = [];
  const headerRow = 10;
  
  // ▼ 変更箇所: 開始行を動的に取得
  const dataStartRow = findDataStartRow_internal(sheet);
  
  const DATE_START_COLUMN = 9;  // I列
  const DATE_END_COLUMN = 39;   // AM列

  const headerValues = sheet.getRange(headerRow, 1, 2, sheet.getMaxColumns()).getDisplayValues();
  const row10 = headerValues[0];
  const row11 = headerValues[1];

  let nameColumnIndex = -1;
  let todayColumnIndex = -1;

  for (let i = 0; i < row10.length; i++) {
    const cell10 = row10[i].toString().trim();
    const cell11 = row11[i].toString().trim();
    if (cell10 === "氏名" || cell11 === "氏名") {
      nameColumnIndex = i;
    }
  }
  
  for (let i = DATE_START_COLUMN - 1; i < DATE_END_COLUMN; i++) {
    if (i >= row10.length) break; 
    
    const cell10 = row10[i].toString().trim();
    const parts = cell10.split('/');
    if (parts.length === 2 && parseInt(parts[0], 10) === currentMonth && parseInt(parts[1], 10) === currentDay) {
      todayColumnIndex = i;
      break; 
    }
  }

  if (nameColumnIndex === -1 || todayColumnIndex === -1) {
    if (nameColumnIndex === -1) Logger.log(`エラー: 10/11行目に「氏名」が見つかりませんでした。`);
    if (todayColumnIndex === -1) Logger.log(`エラー: 今日(${currentMonth}/${currentDay})の日付が見つかりませんでした。`);
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return [];

  const values = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, sheet.getMaxColumns()).getDisplayValues();

  for (const rowData of values) {
    if (rowData[0].toString().toUpperCase() === 'TRUE') {
      continue;
    }

    const name = rowData[nameColumnIndex]?.trim() || "";
    
    // ▼ 変更箇所: 集計行（「時間」「合計」や数字のみのセル）に到達したらループを強制終了
    if (name === "時間" || name === "合計" || /^\d+$/.test(name)) {
      break; 
    }

    // 途中の空行はスキップ
    if (!name) continue;

    const shift = rowData[todayColumnIndex]?.trim() || "";
    
    if (shift && shift !== '0' && !NON_WORK_STRINGS.includes(shift) && /\d/.test(shift)) {
      names.push(name);
    }
  }
  return names;
}

function writeToDestinationSheet_internal(names, dateString) {
  const destSpreadsheet = SpreadsheetApp.openById(DESTINATION_SPREADSHEET_ID_FOR_SHIFT);
  const destSheet = destSpreadsheet.getSheetByName(DEST_SHEET_NAME_FOR_SHIFT);
  if (!destSheet) {
    Logger.log(`書き込み先シート「${DEST_SHEET_NAME_FOR_SHIFT}」が見つかりません。`);
    return;
  }
  
  if (destSheet.getLastRow() >= 2) {
    destSheet.getRange(2, 1, destSheet.getLastRow() - 1, destSheet.getMaxColumns()).clearContent();
  }
  
  destSheet.getRange("A1").setValue("日付");
  destSheet.getRange("B1").setValue("DS部勤務者");

  if (names && names.length > 0) {
    const valuesToWrite = names.map(name => [dateString, name]);
    destSheet.getRange(2, 1, valuesToWrite.length, 2).setValues(valuesToWrite);
    Logger.log(`${names.length}名の勤務者を書き込みました: ${names.join(', ')}`);
  } else {
    destSheet.getRange("A2").setValue(dateString);
    destSheet.getRange("B2").setValue("不在");
    Logger.log(`本日の勤務者はいませんでした。「不在」と書き込みました。`);
  }
}