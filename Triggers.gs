/**
 * 设置编辑触发器
 * 如果已存在则先删除再创建新的触发器
 */
function createEditTrigger() {
  try {
    // 删除现有的编辑触发器
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onEdit') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // 创建新的编辑触发器
    const ss = SpreadsheetApp.getActive();
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
      
    SpreadsheetApp.getActive().toast('编辑触发器已安装', '提示', 3);
  } catch (error) {
    Logger.log('Error creating edit trigger: ' + error);
  }
}

/**
 * 打开文档时的触发器
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('配置表工具')
    .addItem('新建页签', 'createNewSheetTab')
    .addItem('比较差异', 'showCompareDialog')
    .addItem('合并表格', 'showMergeDialog')
    .addItem('清除所有标记', 'clearAllHighlights')
    .addToUi();
}

/**
 * 当编辑表格时的触发器
 * @param {Object} e 编辑事件对象
 */
function onEdit(e) {
  try {
    // 1. 处理原值记录
    const range = e.range;
    const sheet = range.getSheet();
    const oldValue = e.oldValue || '';
    const newValue = range.getValue();
    
    // 如果值没有变化，不做处理
    if (oldValue === newValue) {
      return;
    }
    
    // 获取当前单元格的注释
    let note = range.getNote();
    const basePrefix = "原值: ";
    
    // 如果没有原值注释，或者注释不是以原值开头，直接设置新的原值
    if (!note || !note.startsWith(basePrefix)) {
      range.setNote(basePrefix + oldValue);
    } else {
      // 如果已经有其他注释（比如冲突信息），保留这些信息
      const lines = note.split('\n');
      const updatedLines = lines.map(line => {
        if (line.startsWith(basePrefix)) {
          return basePrefix + oldValue;
        }
        return line;
      });
      range.setNote(updatedLines.join('\n'));
    }

    // 2. 处理ID检查
    const column = range.getColumn();
    const headerRange = sheet.getRange(1, column);
    const headerValue = headerRange.getValue();
    
    // 检查是否编辑的是 ID 列
    if (headerValue && headerValue.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)) {
      // 设置一个短暂的延迟，确保值已经更新
      Utilities.sleep(100);
      // 只检查当前编辑的单元格
      checkIdConflicts({
        sheet: sheet,
        range: range
      });
    }
  } catch (error) {
    console.error('onEdit触发器出错:', error);
  }
}