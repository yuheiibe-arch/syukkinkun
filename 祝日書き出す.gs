function writeWeekendAndHolidays() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName('休日当番拠点長');
  const holidaySheet = ss.getSheetByName('祝日');

  const startRow = 14;
  const startDate = new Date('2026-01-12');
  const endDate = new Date('2027-03-31');

  // 祝日取得（yyyy/MM/dd で正規化）
  const holidayValues = holidaySheet.getRange('A1:A')
    .getValues()
    .flat()
    .filter(v => v)
    .map(d => Utilities.formatDate(new Date(d), 'Asia/Tokyo', 'yyyy/MM/dd'));

  const results = [];
  const dayNames = ['日','月','火','水','木','金','土'];

  for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    const day = d.getDay();
    const dateStr = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
    const isHoliday = holidayValues.includes(dateStr);
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

  // 既存データ削除 → 書き込み
  targetSheet.getRange(startRow, 1, targetSheet.getMaxRows(), 1).clearContent();
  if (results.length > 0) {
    targetSheet.getRange(startRow, 1, results.length, 1).setValues(results);
  }
}
