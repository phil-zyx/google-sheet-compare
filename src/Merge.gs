/**
 * 执行合并操作
 * @param {Object} config 合并配置
 * @param {Sheet} targetSheet 目标表格，如果不指定则使用config中的targetSheet
 * @returns {Object} 合并结果
 */
function mergeSheets(config, targetSheet) {
  if (!config || !config.sourceSheet || (!config.targetSheet && !targetSheet)) {
    return {
      success: false,
      message: "配置参数无效"
    };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(config.sourceSheet);
  var targetSheetName = targetSheet ? targetSheet.getName() : config.targetSheet;
  var targetSheet = targetSheet ? targetSheet : ss.getSheetByName(targetSheetName);
  
  if (!sourceSheet || !targetSheet) {
    return {
      success: false,
      message: "未找到指定的表格，请检查表格名称"
    };
  }

  try {
    // 获取三个表的数据：源表、目标表和基准表
    var sourceRange = sourceSheet.getDataRange();
    var targetRange = targetSheet.getDataRange();
    
    var sourceData = sourceRange.getValues();
    var targetData = targetRange.getValues();
    
    // 获取源表的注释（用于获取基准值）
    var sourceNotes = sourceRange.getNotes();
    var sourceBaseData = sourceNotes.map(row => 
      row.map(note => {
        var baseValue = NoteManager.getSystemNote(note, NOTE_CONSTANTS.TYPES.BASE_VALUE);
        return baseValue;
      })
    );
    
    // 获取表头
    var sourceHeaders = sourceData[0];
    var targetHeaders = targetData[0];
    
    // 找到ID列
    var sourceIdColIndex = -1;
    var targetIdColIndex = -1;
    
    sourceHeaders.forEach((header, index) => {
      if (header.toString().endsWith(MERGE_CONSTANTS.ID_SUFFIX)) {
        sourceIdColIndex = index;
      }
    });
    
    targetHeaders.forEach((header, index) => {
      if (header.toString().endsWith(MERGE_CONSTANTS.ID_SUFFIX)) {
        targetIdColIndex = index;
      }
    });
    
    if (sourceIdColIndex === -1 || targetIdColIndex === -1) {
      return {
        success: false,
        message: `未找到ID列（以${MERGE_CONSTANTS.ID_SUFFIX}结尾的列）`
      };
    }

    // 创建表头映射
    var headerMap = {};
    sourceHeaders.forEach((header, index) => {
      headerMap[header] = {sourceIndex: index, targetIndex: -1};
    });
    
    // 检查新增列
    var newColumns = [];
    sourceHeaders.forEach((header, index) => {
      if (!targetHeaders.includes(header) && !header.toString().endsWith(MERGE_CONSTANTS.ID_SUFFIX)) {
        newColumns.push({
          header: header,
          sourceIndex: index
        });
      }
    });

    // 如果有新增列，在目标表和预览表中添加这些列
    if (newColumns.length > 0) {
      // 在目标表最后添加新列
      targetHeaders = targetHeaders.concat(newColumns.map(col => col.header));
      targetSheet.getRange(1, targetHeaders.length - newColumns.length + 1, 1, newColumns.length)
        .setValues([newColumns.map(col => col.header)])
        .setBackground(MERGE_CONSTANTS.COLORS.NEW);
      
      // 更新headerMap
      newColumns.forEach((col, idx) => {
        headerMap[col.header].targetIndex = targetHeaders.length - newColumns.length + idx;
      });
      
      // 为新列添加空值
      var emptyColumns = Array(newColumns.length).fill('');
      for (var i = 1; i < targetData.length; i++) {
        targetSheet.getRange(i + 1, targetHeaders.length - newColumns.length + 1, 1, newColumns.length)
          .setValues([emptyColumns]);
      }
      
      // 更新targetData以包含新列
      targetData = targetSheet.getDataRange().getValues();
    }

    targetHeaders.forEach((header, index) => {
      if (headerMap[header]) {
        headerMap[header].targetIndex = index;
      }
    });

    // 将目标表数据转换为以ID为键的Map
    var targetDataMap = new Map();
    for (var i = 1; i < targetData.length; i++) {
      var id = targetData[i][targetIdColIndex];
      if (id) {
        targetDataMap.set(id.toString(), {
          rowIndex: i,
          data: targetData[i],
        });
      }
    }

    // 记录需要处理的变更
    var changes = {
      newRows: [],
      updates: [], // 新增：记录可以直接更新的行
      conflicts: []
    };

    // 处理源表数据
    for (var i = 1; i < sourceData.length; i++) {
      var sourceRow = sourceData[i];
      var id = sourceRow[sourceIdColIndex];
      
      if (!id) continue; // 跳过空ID行
      
      id = id.toString();
      var targetRow = targetDataMap.get(id);
      
      if (!targetRow) {
        // 新行，直接添加到新行列表
        changes.newRows.push(sourceRow);
      } else {
        // 检查修改情况
        var hasConflict = false;
        var conflictColumns = [];
        var updateColumns = [];
        
        for (var header in headerMap) {
          var sourceIndex = headerMap[header].sourceIndex;
          var targetIndex = headerMap[header].targetIndex;
          
          if (targetIndex === -1) continue; // 跳过目标表中不存在的列
          
          // 目标表中的当前值
          var currentValue = targetRow.data[targetIndex];
          // 当前表的当前值
          var sourceValue = sourceRow[sourceIndex];
          // 当前表中标记的 base
          var sourceBaseValue = sourceBaseData[i][sourceIndex];
          
          // 标准化值
          function normalizeValue(value) {
            if (value === null || value === undefined) return '';
            
            // 如果是数字或者可以转换为数字
            const num = Number(value);
            if (!isNaN(num)) {
              // 对于整数，返回整数字符串
              if (Number.isInteger(num)) {
                return String(num);
              }
              // 对于小数，统一格式化（去除末尾的0）
              return String(parseFloat(num.toFixed(10)));
            }
            
            // 非数字类型，转换为字符串
            return String(value);
          }
          
          // 检查源表和目标表是否都进行了修改
          var sourceModified = sourceBaseValue && normalizeValue(sourceValue) !== normalizeValue(sourceBaseValue);
          if (sourceModified && normalizeValue(sourceBaseValue) !== normalizeValue(currentValue) && normalizeValue(sourceValue) !== normalizeValue(currentValue) ) {
            hasConflict = true;
            conflictColumns.push({
              header: header,
              sourceValue: sourceValue,
              targetValue: currentValue,
              sourceBaseValue: sourceBaseValue,
            });
          } else if (sourceModified) {
            updateColumns.push({
              header: header,
              sourceValue: sourceValue,
              baseValue: sourceValue // 更新基准值为新的源值
            });
          }
        }
        
        if (hasConflict) {
          changes.conflicts.push({
            id: id,
            sourceRowIndex: i,
            targetRowIndex: targetRow.rowIndex,
            columns: conflictColumns
          });
        } else if (updateColumns.length > 0) {
          changes.updates.push({
            id: id,
            sourceRowIndex: i,
            targetRowIndex: targetRow.rowIndex,
            columns: updateColumns
          });
        }
      }
    }

    // 处理变更
    // 1. 添加新行
    if (changes.newRows.length > 0) {
      var lastRow = targetSheet.getLastRow();
      var newRowsRange = targetSheet.getRange(lastRow + 1, 1, changes.newRows.length, sourceHeaders.length);
      newRowsRange.setValues(changes.newRows);
      newRowsRange.setBackground(MERGE_CONSTANTS.COLORS.NEW);
      
      // 为新行添加基准值注释
      var newRowsNotes = changes.newRows.map(row => 
        row.map(value => NoteManager.addSystemNote('', NOTE_CONSTANTS.TYPES.BASE_VALUE, value.toString()))
      );
      newRowsRange.setNotes(newRowsNotes);
    }

    // 2. 处理可以直接更新的行
    changes.updates.forEach(update => {
      update.columns.forEach(col => {
        var targetIndex = headerMap[col.header].targetIndex;
        var range = targetSheet.getRange(update.targetRowIndex + 1, targetIndex + 1);
        range.setValue(col.sourceValue);
        range.setBackground(MERGE_CONSTANTS.COLORS.UPDATED);
        
        // 更新基准值注释
        const currentNote = range.getNote();
        const newNote = NoteManager.addSystemNote(
          NoteManager.removeSystemNote(currentNote, NOTE_CONSTANTS.TYPES.BASE_VALUE),
          NOTE_CONSTANTS.TYPES.BASE_VALUE,
          col.sourceValue.toString()
        );
        range.setNote(newNote);
      });
    });

    // 3. 标记冲突
    changes.conflicts.forEach(conflict => {
      conflict.columns.forEach(col => {
        var targetIndex = headerMap[col.header].targetIndex;
        var range = targetSheet.getRange(conflict.targetRowIndex + 1, targetIndex + 1);
        range.setBackground(MERGE_CONSTANTS.COLORS.CONFLICT);
        
        const conflictInfo = `${config.targetSheet}: ${col.targetValue}\n${config.sourceSheet}: ${col.sourceValue}`;
        const currentNote = range.getNote();
        const newNote = NoteManager.addSystemNote(
          NoteManager.removeSystemNote(currentNote, NOTE_CONSTANTS.TYPES.CONFLICT),
          NOTE_CONSTANTS.TYPES.CONFLICT,
          conflictInfo
        );
        range.setNote(newNote);
      });
    });

    // 在合并成功后记录日志
    LogManager.addLog(
      LOG_CONSTANTS.TYPES.MERGE,
      config.sourceSheet,
      "生成合并预览成功",
      `目标表格：${targetSheetName}\n` +
      `新增行数：${changes.newRows.length}\n` +
      `更新行数：${changes.updates.length}\n` +
      `冲突行数：${changes.conflicts.length}`
    );

    return {
      success: true,
      message: `合并完成\n新增行数: ${changes.newRows.length}\n更新行数: ${changes.updates.length}\n冲突行数: ${changes.conflicts.length}`,
      changes: changes
    };
  } catch (error) {
    return {
      success: false,
      message: "合并过程中出错: " + error.toString()
    };
  }
}

/**
 * 从注释中提取基准值
 */
function extractBaseValue(note) {
  return NoteManager.getSystemNote(note, NOTE_CONSTANTS.TYPES.BASE_VALUE);
}

/**
 * 确认合并预览表到目标表
 * @param {string} sourceSheetName 源表格名称
 * @param {string} targetSheetName 目标表格名称
 * @param {string} previewSheetName 预览表格名称
 * @returns {Object} 合并结果
 */
function confirmMergeFromPreview(sourceSheetName, targetSheetName, previewSheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName(sourceSheetName);
    var targetSheet = ss.getSheetByName(targetSheetName);
    var previewSheet = ss.getSheetByName(previewSheetName);
    
    if (!sourceSheet || !targetSheet || !previewSheet) {
      return {
        success: false,
        message: "未找到指定的表格，请检查表格名称"
      };
    }

    // 获取预览表的数据和注释
    var previewRange = previewSheet.getDataRange();
    var previewNotes = previewRange.getNotes();
    
    // 检查是否存在未解决的冲突
    var hasUnresolvedConflicts = false;
    for (var i = 0; i < previewNotes.length; i++) {
      for (var j = 0; j < previewNotes[i].length; j++) {
        var note = previewNotes[i][j];
        var conflictInfo = NoteManager.getSystemNote(note, NOTE_CONSTANTS.TYPES.CONFLICT);
        if (conflictInfo) {
          hasUnresolvedConflicts = true;
          break;
        }
      }
      if (hasUnresolvedConflicts) break;
    }
    
    if (hasUnresolvedConflicts) {
      return {
        success: false,
        message: "存在未解决的冲突，请先解决所有冲突后再确认合并"
      };
    }

    // 获取预览表的其他数据
    var previewData = previewRange.getValues();
    var previewBackgrounds = previewRange.getBackgrounds();

    // 只复制有变化的单元格（新增、更新或解决的冲突）
    for (var i = 0; i < previewData.length; i++) {
      for (var j = 0; j < previewData[i].length; j++) {
        var background = previewBackgrounds[i][j];
        // 检查单元格是否有变化（根据背景色判断）
        if (background === MERGE_CONSTANTS.COLORS.NEW || 
            background === MERGE_CONSTANTS.COLORS.UPDATED || 
            background === MERGE_CONSTANTS.COLORS.RESOLVED) {
          var targetCell = targetSheet.getRange(i + 1, j + 1);
          targetCell.setValue(previewData[i][j]);
          targetCell.setNote(previewNotes[i][j]);
          targetCell.setBackground(MERGE_CONSTANTS.COLORS.MERGED);  // 或者完全不设置背景色
        }
      }
    }

    // 标记源表为已合并
    var headerRow = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    var statusColIndex = -1;
    
    // 在源表格名称后添加"(已合并)"标记
    const newSourceSheetName = sourceSheetName.endsWith('(已合并)') 
      ? sourceSheetName 
      : `${sourceSheetName}(已合并)`;
    sourceSheet.setName(newSourceSheetName);
    
    // 4. 删除预览表
    ss.deleteSheet(previewSheet);
    
    // 5. 清除预览状态
    const cache = CacheService.getScriptCache();
    cache.remove('merge_preview_state');
    
    // 6. 激活目标页签
    targetSheet.activate();
    
    // 记录确认合并成功的日志
    LogManager.addLog(
      LOG_CONSTANTS.TYPES.MERGE,
      sourceSheetName,
      "确认合并成功",
      `目标表格：${targetSheetName}`
    );
    
    return {
      success: true,
      message: "合并完成！只更新了变更的单元格。源表已标记为已合并状态。"
    };
    
  } catch (error) {
    console.error('确认合并失败:', error);
    return {
      success: false,
      message: "确认合并过程中出错: " + error.toString()
    };
  }
}

/**
 * 获取预览状态
 * @param {string} sourceSheet 源表格名称
 * @param {string} targetSheet 目标表格名称
 * @returns {Object|null} 预览状态对象，如果没有则返回null
 */
function getPreviewState(sourceSheet, targetSheet) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const previewSheetName = `${sourceSheet} -> ${targetSheet} 合并预览`;
    const previewSheet = ss.getSheetByName(previewSheetName);
    
    if (!previewSheet) {
      return null;
    }
    
    return {
      previewSheetName: previewSheetName,
      sourceSheet: sourceSheet,
      targetSheet: targetSheet
    };
  } catch (e) {
    console.error('获取预览状态失败:', e);
    return null;
  }
}

/**
 * 显示确认对话框
 */
function showConfirmDialog() {
  try {
    // 获取当前活动页签
    const activeSheet = SpreadsheetApp.getActiveSheet();
    const sheetName = activeSheet.getName();
    console.log('当前页签名称:', sheetName);
    
    // 从页签名称中解析源表和目标表
    // 预览页签的命名格式为: "sourceSheet -> targetSheet 合并预览"
    const match = sheetName.match(/^(.*?)\s*->\s*(.*?)\s*合并预览$/);
    console.log('页签名称匹配结果:', match);
    
    if (!match) {
      showAlert('请在合并预览页签中使用此功能');
      return;
    }
    
    const [_, sourceSheet, targetSheet] = match;
    console.log('解析出的源表和目标表:', { sourceSheet, targetSheet });
    
    // 检查预览状态
    const previewState = getPreviewState(sourceSheet.trim(), targetSheet.trim());
    console.log('获取到的预览状态:', previewState);
    
    if (!previewState) {
      showAlert('未找到有效的预览状态，请重新执行合并预览');
      return;
    }
    
    console.log('准备显示对话框');
    // 显示合并对话框
    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutputFromFile('MergeDialog')
      .setWidth(600)
      .setHeight(600)
      .setTitle('合并表格');
    
    ui.showModalDialog(html, '合并表格');
    console.log('对话框已显示');
  } catch (error) {
    console.error('显示确认对话框失败:', error);
    showAlert('显示确认对话框失败: ' + error.toString());
  }
}

/**
 * 预览合并结果
 * @param {Object} config 合并配置
 * @returns {Object} 预览结果
 */
function previewMerge(config) {
  if (!config || !config.sourceSheet || !config.targetSheet) {
    return {
      success: false,
      message: "配置参数无效"
    };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 检查是否已存在预览状态
  const existingPreview = getPreviewState(config.sourceSheet, config.targetSheet);
  if (existingPreview) {
    return {
      success: false,
      message: `已存在 "${config.sourceSheet} -> ${config.targetSheet}" 的预览，请先完成或取消现有预览`
    };
  }

  Logger.log('解析预合并参数: %s', config);
  var sourceSheet = ss.getSheetByName(config.sourceSheet);
  var targetSheet = ss.getSheetByName(config.targetSheet);
  
  if (!sourceSheet || !targetSheet) {
    return {
      success: false,
      message: "未找到指定的表格，请检查表格名称"
    };
  }

  try {
    // 创建预览表格
    var previewSheet = createPreviewSheet(config.targetSheet, config.sourceSheet);
    
    // 复制目标表格的数据到预览表格，但不包括背景色和系统注释
    var targetRange = targetSheet.getDataRange();
    var targetData = targetRange.getValues();
    
    // 只复制数据，不复制格式和注释
    previewSheet.getRange(1, 1, targetData.length, targetData[0].length)
      .setValues(targetData);

    // 执行合并预览
    var result = mergeSheets(config, previewSheet);
    
    if (result.success) {
      return {
        success: true,
        previewSheetName: previewSheet.getName(),
        sourceSheet: config.sourceSheet,
        targetSheet: config.targetSheet,
        changes: result.changes,
        message: `预览已生成，请在"${previewSheet.getName()}"表格中查看\n${result.message}`
      };
    } else {
      // 如果预览失败，删除预览表格
      deletePreviewSheet(previewSheet.getName());
      return result;
    }
  } catch (error) {
    console.error('预览失败:', error);
    return {
      success: false,
      message: "预览生成失败：" + error.toString()
    };
  }
}

/**
 * 删除预览表格
 * @param {string} previewSheetName 预览表格名称
 * @returns {boolean} 是否成功删除
 */
function deletePreviewSheet(previewSheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(previewSheetName);
    if (sheet) {
      ss.deleteSheet(sheet);
      return true;
    }
    return false;
  } catch (error) {
    console.error('删除预览表格失败:', error);
    return false;
  }
}

/**
 * 解决合并冲突
 * @param {Object} config 解决配置
 * @returns {Object} 操作结果
 */
function resolveConflict(config) {
  if (!config || !config.row || !config.header || config.value === undefined) {
    return {
      success: false,
      message: "参数无效"
    };
  }

  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 找到对应的列
    var colIndex = headers.findIndex(h => h === config.header);
    if (colIndex === -1) {
      return {
        success: false,
        message: "未找到指定列: " + config.header
      };
    }

    // 更新单元格值
    var cell = sheet.getRange(config.row + 1, colIndex + 1);
    cell.setValue(config.value);
    cell.setBackground(MERGE_CONSTANTS.COLORS.RESOLVED);
    
    // 更新注释：清除冲突信息，更新基准值
    const currentNote = cell.getNote();
    let newNote = NoteManager.removeSystemNote(currentNote, NOTE_CONSTANTS.TYPES.CONFLICT);
    newNote = NoteManager.addSystemNote(
      newNote,
      NOTE_CONSTANTS.TYPES.BASE_VALUE,
      config.value.toString()
    );
    cell.setNote(newNote);

    return {
      success: true,
      message: "已更新单元格值"
    };
  } catch (error) {
    return {
      success: false,
      message: "更新失败: " + error.toString()
    };
  }
}

/**
 * 显示提示信息
 * @param {string} message 提示信息
 */
function showAlert(message) {
  SpreadsheetApp.getUi().alert(message);
}

/**
 * 显示对话框
 * @param {string} [dialogType='merge'] 对话框类型
 */
function showDialog(dialogType = 'merge') {
  // 创建新的对话框实例
  var html = HtmlService.createHtmlOutputFromFile('MergeDialog')
    .setWidth(600)
    .setHeight(600)
    .setTitle('合并表格');
  
  // 使用showModalDialog而不是showDialog以确保对话框总是在前面
  SpreadsheetApp.getUi().showModalDialog(html, '合并表格');
}

/**
 * 创建预览表格
 * @param {string} targetSheetName 目标表格名称
 * @param {string} sourceSheetName 源表格名称
 * @returns {Sheet} 预览表格
 */
function createPreviewSheet(targetSheetName, sourceSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var previewName = `${sourceSheetName} -> ${targetSheetName} 合并预览`;
  var existingSheet = ss.getSheetByName(previewName);
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }
  return ss.insertSheet(previewName);
}

/**
 * 显示合并对话框
 */
function showMergeDialog() {
  // 检查当前表格是否已经标记为已合并
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var currentSheetName = currentSheet.getName();
  
  if (currentSheetName.endsWith('(已合并)')) {
    SpreadsheetApp.getUi().alert(
      '无法合并',
      '当前表格已经标记为已合并状态，不能重复发起合并。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var html = HtmlService.createHtmlOutputFromFile('MergeDialog')
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '表格合并工具');
}
