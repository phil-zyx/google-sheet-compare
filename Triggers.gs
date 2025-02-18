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
    .addItem('清除系统注释', 'clearAllSystemNotes')
    .addToUi();
}

/**
 * 当编辑表格时的触发器
 * @param {Object} e 编辑事件对象
 */
function onEdit(e) {
  try {
    // 1. 处理基准值记录
    const range = e.range;
    const sheet = range.getSheet();
    const oldValue = e.oldValue;
    const newValue = range.getValue();
    
    // 获取当前单元格的注释
    let note = range.getNote();
    
    // 只在有oldValue且当前没有基准值记录时记录
    if (oldValue !== undefined && 
        !NoteManager.getSystemNote(note, NOTE_CONSTANTS.TYPES.BASE_VALUE)) {
      // 添加基准值到系统注释，保留用户原有注释
      const newNote = NoteManager.addSystemNote(
        note,
        NOTE_CONSTANTS.TYPES.BASE_VALUE,
        oldValue.toString()
      );
      range.setNote(newNote);
    }

    // 2. 处理ID检查 - 无论是否有oldValue都需要检查
    const column = range.getColumn();
    const headerRange = sheet.getRange(1, column);
    const headerValue = headerRange.getValue();
    
    // 检查是否编辑的是 ID 列
    if (headerValue && headerValue.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)) {
      // 设置一个短暂的延迟，确保值已经更新
      Utilities.sleep(100);
      // 只检查 ID 列的单元格
      const idRange = sheet.getRange(range.getRow(), column, range.getNumRows(), 1);
      checkIdConflicts({
        sheet: sheet,
        range: idRange
      });
    }
  } catch (error) {
    console.error('onEdit触发器出错:', error);
  }
}

/**
 * 清除所有系统注释
 */
function clearAllSystemNotes() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getDataRange();
    const notes = range.getNotes();
    let hasChanges = false;
    
    // 处理每个单元格的注释
    const cleanedNotes = notes.map(row => 
      row.map(note => {
        if (!note) return note; // 如果没有注释，直接返回
        
        const cleaned = NoteManager.removeAllSystemNotes(note);
        if (cleaned !== note) {
          hasChanges = true;
        }
        return cleaned || ''; // 确保返回空字符串而不是null或undefined
      })
    );
    
    // 只有在有变更时才更新注释
    if (hasChanges) {
      range.setNotes(cleanedNotes);
      return {
        success: true,
        message: "已清除所有系统注释"
      };
    }
    
    return {
      success: true,
      message: "没有找到需要清除的系统注释"
    };
  } catch (error) {
    console.error('清除系统注释失败:', error);
    return {
      success: false,
      message: "清除系统注释失败: " + error.toString()
    };
  }
}