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
    .addItem('清除所有标记', 'clearAllMarks')
    .addToUi();
}

/**
 * 当编辑表格时的触发器
 * @param {Object} e 编辑事件对象
 */
function onEdit(e) {
  try {
    // 检查表格是否包含ID列
    const sheet = e.range.getSheet();
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const hasIdColumn = headerRow.some(header => 
      header && header.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)
    );
    
    // 如果没有ID列，直接返回
    if (!hasIdColumn) return;

    // 1. 处理基准值记录和颜色标记
    const range = e.range;
    const oldValue = e.oldValue;
    const newValue = range.getValue();
    
    // 获取当前单元格的背景色和注释
    const currentBg = range.getBackground();
    let note = range.getNote();
    
    // 如果当前单元格已经是新增状态（绿色），则不做任何改变
    if (currentBg === SHEET_CONSTANTS.COLORS.ADDED) {
      return;
    }

    // 检查是否已经有修改记录（通过背景色判断）
    const isAlreadyModified = currentBg === SHEET_CONSTANTS.COLORS.MODIFIED;

    if (oldValue !== undefined) {  // 是修改操作
      // 设置为修改颜色（浅蓝色）
      range.setBackground(SHEET_CONSTANTS.COLORS.MODIFIED);
      
      // 只在首次修改时记录基准值
      if (!isAlreadyModified) {
        // 保持原始值的格式
        let baseValue = oldValue;
        // 如果是整数，确保以整数形式存储
        if (Number.isInteger(Number(oldValue))) {
          baseValue = parseInt(oldValue, 10);
        }
        
        // 添加基准值到系统注释，保留用户原有注释
        const newNote = NoteManager.addSystemNote(
          note,
          NOTE_CONSTANTS.TYPES.BASE_VALUE,
          baseValue.toString()
        );
        range.setNote(newNote);
      }
    } else if (newValue && newValue.toString().trim() !== '') {
      // 如果是新增值，设置为新增颜色（淡绿色）
      range.setBackground(SHEET_CONSTANTS.COLORS.ADDED);
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
 * 清除所有标记和系统注释
 * @param {boolean} showConfirm 是否显示确认对话框
 * @returns {Object} 操作结果
 */
function clearAllMarks(showConfirm = true) {
  try {
    if (showConfirm) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        '确认清除',
        '是否要清除当前表格中所有的比较标记和系统注释？',
        ui.ButtonSet.YES_NO
      );

      if (response !== ui.Button.YES) {
        return { success: false, message: "操作已取消" };
      }
    }

    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getDataRange();
    const [backgrounds, notes, values] = [
      range.getBackgrounds(),
      range.getNotes(),
      range.getValues()
    ];

    const newBackgrounds = [];
    const newNotes = [];
    const newValues = [];
    const rowsToKeep = [];

    // 检查每一行，标记需要保留的行
    for (let i = 0; i < backgrounds.length; i++) {
      let isDeletedRow = true;
      let hasHighlight = false;

      // 检查这一行是否是比较时新增的行
      for (let j = 0; j < backgrounds[i].length; j++) {
        const currentBg = backgrounds[i][j];
        const currentNote = notes[i][j];

        if (currentBg === COMPARE_CONSTANTS.COLORS.REMOVED) {
          hasHighlight = true;
        }

        if (currentNote && currentNote.includes("此行在基准表中不存在")) {
          hasHighlight = true;
        }

        // 如果这一行有任何非高亮的单元格，说明不是新增的行
        if (currentBg !== COMPARE_CONSTANTS.COLORS.REMOVED && 
            currentBg !== COMPARE_CONSTANTS.COLORS.MODIFIED && 
            currentBg !== COMPARE_CONSTANTS.COLORS.ADDED && 
            currentBg !== COMPARE_CONSTANTS.COLORS.HEADER_MODIFIED) {
          isDeletedRow = false;
        }
      }

      // 如果这一行不是新增的行，或者是第一行（表头），就保留它
      if (!isDeletedRow || !hasHighlight || i === 0) {
        rowsToKeep.push(i);

        const backgroundRow = [];
        const noteRow = [];

        for (let j = 0; j < backgrounds[i].length; j++) {
          const currentBg = backgrounds[i][j];
          let currentNote = notes[i][j];

          // 清除所有比较标记的背景色
          if (currentBg === COMPARE_CONSTANTS.COLORS.MODIFIED || 
              currentBg === COMPARE_CONSTANTS.COLORS.ADDED || 
              currentBg === COMPARE_CONSTANTS.COLORS.REMOVED || 
              currentBg === COMPARE_CONSTANTS.COLORS.HEADER_MODIFIED ||
              currentBg === SHEET_CONSTANTS.COLORS.MODIFIED ||
              currentBg === SHEET_CONSTANTS.COLORS.ADDED ||
              currentBg === SHEET_CONSTANTS.COLORS.REMOVED ||
              currentBg === SHEET_CONSTANTS.COLORS.HEADER_MODIFIED ||
              currentBg === SHEET_CONSTANTS.COLORS.CONFLICT ||
              currentBg === MERGE_CONSTANTS.COLORS.NEW ||
              currentBg === MERGE_CONSTANTS.COLORS.CONFLICT ||
              currentBg === MERGE_CONSTANTS.COLORS.UPDATED ||
              currentBg === MERGE_CONSTANTS.COLORS.RESOLVED) {
            backgroundRow.push(null);
          } else {
            backgroundRow.push(currentBg);
          }

          // 清除系统注释
          if (currentNote) {
            currentNote = NoteManager.removeAllSystemNotes(currentNote);
            noteRow.push(currentNote || '');
          } else {
            noteRow.push('');
          }
        }

        newBackgrounds.push(backgroundRow);
        newNotes.push(noteRow);
        newValues.push(values[i]);
      }
    }

    // 更新表格
    if (rowsToKeep.length < backgrounds.length) {
      // 如果有行被删除，更新表格并删除多余的行
      const newRange = sheet.getRange(1, 1, newBackgrounds.length, backgrounds[0].length);
      newRange.setBackgrounds(newBackgrounds);
      newRange.setNotes(newNotes);
      newRange.setValues(newValues);

      if (backgrounds.length > newBackgrounds.length) {
        sheet.deleteRows(newBackgrounds.length + 1, backgrounds.length - newBackgrounds.length);
      }
    } else {
      // 如果没有行被删除，只更新背景色和注释
      range.setBackgrounds(newBackgrounds);
      range.setNotes(newNotes);
    }

    return {
      success: true,
      message: "已清除所有标记和系统注释"
    };
  } catch (error) {
    console.error('清除标记和注释失败:', error);
    return {
      success: false,
      message: "清除失败: " + error.toString()
    };
  }
}