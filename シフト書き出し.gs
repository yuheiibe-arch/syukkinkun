// === 設定項目 ===
const SOURCE_SPREADSHEET_ID = '1yj4dG9BzNnUU85C-CdYCwPXougm4XJv3cC2JxwMacK0'; 
const DESTINATION_SPREADSHEET_ID_FOR_SHIFT = '1bKbpLbiDZ7V_pDgnUcRApiSA9jEIQ05Q6Msd1AIi3IU'; 
const DEST_SHEET_NAME_FOR_SHIFT = 'DS部';
const EXCLUDED_SHEET_NAMES = ["祝日", "土日祝業務内容", "参考1", "シフト構成パターン（小谷編集中）", "メンバーTODO", "携帯番号"];
const NON_WORK_STRINGS = ["休", "有給"];

/**
 * シフト表から今日の勤務者と勤務時間を抽出し、「DS部」シートを更新します。
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
    
  const workingStaffList = extractWorkingNames_internal(targetSheet, currentDay, currentMonth);
  writeToDestinationSheet_internal(workingStaffList, dateStringToWrite);
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
 * A列から「除外」を探し、その次の行番号を返す関数（動的スタート）
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
 * シフトシートから勤務者と勤務時間を抽出し、開始時間順で返します。
 */
function extractWorkingNames_internal(sheet, currentDay, currentMonth) {
  const staffList = [];
  const headerRow = 10;
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
    
    // 集計行（「時間」「合計」や数字のみのセル）に到達したらループを強制終了
    if (name === "時間" || name === "合計" || /^\d+$/.test(name)) {
      break; 
    }

    if (!name) continue;

    const shift = rowData[todayColumnIndex]?.trim() || "";
    
    if (shift && shift !== '0' && !NON_WORK_STRINGS.includes(shift) && /\d/.test(shift)) {
      const parsed = parseShiftTime_common(shift);
      staffList.push({
        name: name,
        time: parsed ? parsed.formattedTime : shift,
        startMinutes: parsed ? parsed.startMinutes : 9999,
        endMinutes: parsed ? parsed.endMinutes : 0
      });
    }
  }

  // 開始時間順（同じなら終了時間順）でソート
  staffList.sort((a, b) => {
    if (a.startMinutes !== b.startMinutes) return a.startMinutes - b.startMinutes;
    return a.endMinutes - b.endMinutes;
  });

  return staffList;
}

function writeToDestinationSheet_internal(staffList, dateString) {
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
  destSheet.getRange("C1").setValue("勤務時間");

  if (staffList && staffList.length > 0) {
    const valuesToWrite = staffList.map(item => [dateString, item.name, item.time]);
    destSheet.getRange(2, 1, valuesToWrite.length, 3).setValues(valuesToWrite);
    Logger.log(`${staffList.length}名の勤務者を書き込みました: ${staffList.map(s => `${s.name}(${s.time})`).join(', ')}`);
  } else {
    destSheet.getRange("A2:C2").setValues([[dateString, "不在", ""]]);
    Logger.log(`本日の勤務者はいませんでした。「不在」と書き込みました。`);
  }
}