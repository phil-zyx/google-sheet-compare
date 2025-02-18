/**
 * 检查指定ID是否与其他ID冲突
 */
function checkSingleIdConflict({ value, sheet: sheetName, row, column, columnName }) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 遍历所有表格查找相同ID
  for (const sheet of ss.getSheets()) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnIndex = headers.findIndex(header => header.toString() === columnName);
    if (columnIndex === -1 || sheet.getLastRow() <= 1) continue;
    
    // 获取列数据并检查冲突
    const data = sheet.getRange(2, columnIndex + 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      const [id] = data[i];
      if (!id?.toString().trim()) continue;
      
      const isCurrentCell = sheet.getName() === sheetName && 
                          i + 2 === row && 
                          columnIndex + 1 === column;
      
      if (id.toString() === value.toString() && !isCurrentCell) {
        // 找到第一个冲突就返回
        return [{
          sheet: sheet.getName(),
          row: i + 2,
          column: columnIndex + 1
        }];
      }
    }
  }
  
  return [];
}

/**
 * 检查ID冲突并标记
 */
function checkIdConflicts(editedCell) {
  clearIdConflictHighlights();
  
  if (!editedCell) return;
  
  const { sheet, range } = editedCell;
  const value = range.getValue();
  if (!value) {
    range.setBackground(null).setNote(null);
    return;
  }
  
  const headerValue = sheet.getRange(1, range.getColumn()).getValue();
  const conflicts = checkSingleIdConflict({
    value,
    sheet: sheet.getName(),
    row: range.getRow(),
    column: range.getColumn(),
    columnName: headerValue
  });
  
  if (conflicts.length > 0) {
    const conflictLocations = conflicts.map(loc => `${loc.sheet} 第${loc.row}行`).join('\n');
    const userNote = `${headerValue} ID冲突:\n在以下位置重复:\n${conflictLocations}`;
    
    range
      .setBackground(ID_CHECKER_CONFIG.COLORS.CONFLICT)
      .setNote(NoteManager.addSystemNote(
        userNote,
        NOTE_CONSTANTS.TYPES.CONFLICT,
        JSON.stringify({ locations: conflicts.map(({sheet, row}) => ({sheet, row})) })
      ));
    
    SpreadsheetApp.getActiveSpreadsheet().toast('发现ID冲突，已用红色标记。', '警告', 5);
  }
}

/**
 * 清除所有ID冲突标记
 */
function clearIdConflictHighlights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  ss.getSheets().forEach(sheet => {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    headers.forEach((header, index) => {
      if (!header?.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)) return;
      
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, index + 1, lastRow - 1, 1)
          .setBackground(null)
          .clearNote();
      }
    });
  });
  
  ss.toast('已清除所有ID冲突标记。', '提示', 3);
}
