// ID检查器的配置
const ID_CHECKER_CONFIG = {
  COLORS: {
    CONFLICT: '#ff0000',  // 冲突标记颜色 - 红色
  },
  ID_COLUMN_SUFFIX: '_INT_id',   // ID列的后缀
};

/**
 * 检查指定ID是否与其他ID冲突
 * @param {Object} params - 检查参数
 * @param {string} params.value - 要检查的ID值
 * @param {string} params.sheet - 工作表名称
 * @param {number} params.row - 行号
 * @param {number} params.column - 列号
 * @param {string} params.columnName - 列名
 */
function checkSingleIdConflict(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const conflicts = [];
  
  // 遍历所有表格查找相同ID
  sheets.forEach(sheet => {
    const headerRow = 1;
    const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 找到匹配的列索引
    const columnIndex = headers.findIndex(header => header.toString() === params.columnName);
    if (columnIndex === -1) return; // 如果这个表格没有匹配的列，直接跳过
    
    // 只获取需要的列的数据
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return; // 如果只有表头行，直接跳过
    
    const columnData = sheet.getRange(2, columnIndex + 1, lastRow - 1, 1).getValues();
    
    // 检查这一列的值
    columnData.forEach((row, rowIndex) => {
      const id = row[0];
      if (!id || id === '') return; // 跳过空ID
      
      // 如果找到相同的ID，但不是当前编辑的单元格
      if (id.toString() === params.value.toString() && 
          !(sheet.getName() === params.sheet && 
            rowIndex + 2 === params.row && 
            columnIndex + 1 === params.column)) {
        conflicts.push({
          sheet: sheet.getName(),
          row: rowIndex + 2,
          column: columnIndex + 1
        });
      }
    });
  });
  
  return conflicts;
}

/**
 * 检查当前表格中所有以_INT_id结尾的列的ID冲突
 * @param {Object} [editedCell] - 如果提供，则只检查这个单元格的ID
 */
function checkIdConflicts(editedCell) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 如果提供了编辑的单元格信息，只检查这个特定的ID
  if (editedCell) {
    const { sheet, range } = editedCell;
    const headerRange = sheet.getRange(1, range.getColumn());
    const headerValue = headerRange.getValue();
    const value = range.getValue();
    
    if (!value || value === '') {
      // 如果是删除操作，清除样式和注释
      range.setBackground(null);
      range.setNote(null);
      return;
    }
    
    // 先清除所有冲突标记，确保其他页签上没有残留的标记
    // clearIdConflictHighlights();
    
    // 清除当前单元格的冲突标记
    range.setBackground(null);
    
    // 检查这个特定的ID
    const conflicts = checkSingleIdConflict({
      value: value,
      sheet: sheet.getName(),
      row: range.getRow(),
      column: range.getColumn(),
      columnName: headerValue
    });
    
    if (conflicts.length > 0) {
      // 只标记当前编辑的单元格
      range.setBackground(ID_CHECKER_CONFIG.COLORS.CONFLICT);
      
      // 添加注释说明冲突位置
      const comment = `${headerValue} ID冲突:\n在以下位置重复:\n${
        conflicts
          .map(loc => `${loc.sheet} 第${loc.row}行`)
          .join('\n')
      }`;
      range.setNote(comment);
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `发现ID冲突，已用红色标记。`,
        '警告',
        5
      );
    }
    return;
  }
  
  // 如果没有提供编辑的单元格信息，清除所有标记
  clearIdConflictHighlights();
}

/**
 * 清除所有ID冲突标记
 */
function clearIdConflictHighlights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    const headerRow = 1;
    const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 找到所有以_INT_id结尾的列
    headers.forEach((header, index) => {
      if (!header || !header.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)) {
        return;
      }
      
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) { // 确保有数据行
        const idColumn = sheet.getRange(2, index + 1, lastRow - 1, 1);
        idColumn.setBackground(null);
        idColumn.clearNote();
      }
    });
  });
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '已清除所有ID冲突标记。',
    '提示',
    3
  );
}

/**
 * 当编辑表格时触发的函数
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e 编辑事件对象
 */
function onEdit(e) {
  // 获取编辑的范围
  const range = e.range;
  const sheet = range.getSheet();
  const column = range.getColumn();
  
  // 获取编辑列的表头
  const headerRange = sheet.getRange(1, column);
  const headerValue = headerRange.getValue();
  
  // 检查是否编辑的是 ID 列
  if (headerValue && headerValue.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)) {
    // 设置一个短暂的延迟，确保值已经更新
    Utilities.sleep(100);
    checkIdConflicts({ sheet, range });
  }
}

/**
 * 安装编辑触发器
 */
function installEditTrigger() {
  // 删除现有的触发器以避免重复
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 创建新的触发器
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
    
  SpreadsheetApp.getActive().toast('ID检查触发器已安装', '提示', 3);
}
