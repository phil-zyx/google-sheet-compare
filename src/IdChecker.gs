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
  
  if (!editedCell) return;
  
  const { sheet, range } = editedCell;
  const value = range.getValue();
  if (!value) {
    range.setBackground(null);
    NoteManager.removeMarkFromCell(range, NOTE_CONSTANTS.TYPES.CONFLICT);
    return;
  }
  
  // 先移除历史标记
  NoteManager.removeMarkFromCell(range, NOTE_CONSTANTS.TYPES.CONFLICT);

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
    const userNote = `在以下位置重复:\n${conflictLocations}`;
    
    range.setBackground(ID_CHECKER_CONFIG.COLORS.CONFLICT);
    range.setNote(NoteManager.addSystemNote(
      null,
      NOTE_CONSTANTS.TYPES.CONFLICT,  // 标记类型
      userNote
    ));
    
    SpreadsheetApp.getActiveSpreadsheet().toast('发现ID冲突，已用红色标记。', '警告', 3);
  }
}
