// ID检查器的配置
const ID_CHECKER_CONFIG = {
  COLORS: {
    CONFLICT: '#ff0000',  // 冲突标记颜色 - 红色
  },
  ID_COLUMN_SUFFIX: '_INT_id',   // ID列的后缀
};

/**
 * 检查当前表格中所有以_INT_id结尾的列的ID冲突
 */
function checkIdConflicts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const conflictMap = new Map(); // 用于存储ID及其位置
  const conflicts = new Set(); // 用于存储冲突的ID
  
  // 第一步：收集所有ID及其位置
  sheets.forEach(sheet => {
    const headerRow = 1;
    const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 找出所有以_INT_id结尾的列
    headers.forEach((header, columnIndex) => {
      if (!header || !header.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)) {
        return;
      }
      
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      
      // 从第二行开始遍历（跳过表头）
      for (let row = 1; row < values.length; row++) {
        const id = values[row][columnIndex];
        if (!id) continue; // 跳过空ID
        
        const location = {
          sheet: sheet.getName(),
          row: row + 1,
          column: columnIndex + 1,
          columnName: header.toString()
        };
        
        const key = `${header}_${id}`; // 使用列名和ID组合作为键，这样不同类型的ID就不会互相影响
        
        if (conflictMap.has(key)) {
          conflicts.add(key);
          conflictMap.get(key).push(location);
        } else {
          conflictMap.set(key, [location]);
        }
      }
    });
  });
  
  // 第二步：清除所有现有的冲突标记
  clearIdConflictHighlights();
  
  // 第三步：标记所有冲突
  conflicts.forEach(key => {
    const locations = conflictMap.get(key);
    const [columnName, id] = key.split('_').slice(0, -1).join('_'); // 提取列名（去掉最后的_id）
    
    locations.forEach(loc => {
      const sheet = ss.getSheetByName(loc.sheet);
      const cell = sheet.getRange(loc.row, loc.column);
      cell.setBackground(ID_CHECKER_CONFIG.COLORS.CONFLICT);
      
      // 添加注释说明冲突位置
      const comment = `${columnName} ID冲突: ${id}\n在以下位置重复:\n${
        locations
          .filter(l => !(l.sheet === loc.sheet && l.row === loc.row))
          .map(l => `${l.sheet} 第${l.row}行`)
          .join('\n')
      }`;
      cell.setNote(comment);
    });
  });
  
  // 如果有冲突，显示提示
  if (conflicts.size > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `发现 ${conflicts.size} 个ID冲突，已用红色标记。`,
      '警告',
      5
    );
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '没有发现ID冲突。',
      '提示',
      3
    );
  }
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
    checkIdConflicts();
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
