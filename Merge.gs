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
    
    // 获取注释
    var targetNotes = targetRange.getNotes();
    
    // 获取基准数据（从系统注释中）
    var baseData = targetNotes.map(row => 
      row.map(note => NoteManager.getSystemNote(note, NOTE_CONSTANTS.TYPES.BASE_VALUE))
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
          baseData: baseData[i] // 包含原值信息的注释
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
          
          var currentValue = targetRow.data[targetIndex];
          var sourceValue = sourceRow[sourceIndex];
          
          // 从注释中获取原值
          var note = targetRow.baseData[targetIndex];
          var baseValue = note ? extractBaseValue(note) : currentValue;
          
          if (currentValue !== sourceValue) {
            if (currentValue === baseValue) {
              // 目标值未被修改，可以直接使用源表的修改
              updateColumns.push({
                header: header,
                sourceValue: sourceValue,
                baseValue: baseValue
              });
            } else if (sourceValue !== baseValue) {
              // 源表和目标表都做了修改，且修改不一致
              hasConflict = true;
              conflictColumns.push({
                header: header,
                sourceValue: sourceValue,
                targetValue: currentValue,
                baseValue: baseValue
              });
            }
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
        
        // 添加冲突信息注释
        const conflictInfo = `当前值: ${col.targetValue}\n源表值: ${col.sourceValue}\n基值: ${col.baseValue}`;
        const currentNote = range.getNote();
        const newNote = NoteManager.addSystemNote(
          NoteManager.removeSystemNote(currentNote, NOTE_CONSTANTS.TYPES.CONFLICT),
          NOTE_CONSTANTS.TYPES.CONFLICT,
          conflictInfo
        );
        range.setNote(newNote);
      });
    });

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
 * 添加基准值到注释
 */
function addBaseValue(value) {
  return NoteManager.addSystemNote('', NOTE_CONSTANTS.TYPES.BASE_VALUE, value.toString());
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

    // 获取预览表和目标表的数据
    var previewRange = previewSheet.getDataRange();
    var previewData = previewRange.getValues();
    var previewBackgrounds = previewRange.getBackgrounds();
    var previewNotes = previewRange.getNotes();

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
          // 可以选择设置一个统一的、较淡的背景色来标识已合并的单元格
          targetCell.setBackground('#f5f5f5');  // 或者完全不设置背景色
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
 * 保存预览状态
 * @param {Object} previewData 预览数据
 */
function savePreviewState(previewData) {
  const cache = CacheService.getScriptCache();
  const state = {
    ...previewData,
    timestamp: new Date().getTime()
  };
  
  // 使用源表和目标表的组合作为唯一标识
  const cacheKey = generatePreviewStateKey(previewData.sourceSheet, previewData.targetSheet);
  console.log('保存预览状态:', state, '缓存键:', cacheKey);
  cache.put(cacheKey, JSON.stringify(state), 3600); // 1小时过期
}

/**
 * 获取预览状态
 * @param {string} sourceSheet 源表格名称
 * @param {string} targetSheet 目标表格名称
 * @returns {Object|null} 预览状态对象，如果没有则返回null
 */
function getPreviewState(sourceSheet, targetSheet) {
  const cache = CacheService.getScriptCache();
  const cacheKey = generatePreviewStateKey(sourceSheet, targetSheet);
  const state = cache.get(cacheKey);
  console.log('获取预览状态缓存:', state, '缓存键:', cacheKey);
  
  if (!state) {
    console.log('没有预览状态缓存');
    return null;
  }
  
  try {
    const parsedState = JSON.parse(state);
    // 检查预览表格是否还存在
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const previewSheet = ss.getSheetByName(parsedState.previewSheetName);
    
    if (!previewSheet) {
      // 如果预览表格不存在，清除状态
      cache.remove(cacheKey);
      return null;
    }
    
    console.log('返回预览状态:', parsedState);
    return parsedState;
  } catch (e) {
    console.error('解析预览状态失败:', e);
    return null;
  }
}

/**
 * 生成预览状态的缓存键
 * @param {string} sourceSheet 源表格名称
 * @param {string} targetSheet 目标表格名称
 * @returns {string} 缓存键
 */
function generatePreviewStateKey(sourceSheet, targetSheet) {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return `merge_preview_${spreadsheetId}_${sourceSheet}_${targetSheet}`;
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
    
    // 复制目标表格的所有数据和格式到预览表格
    var targetRange = targetSheet.getDataRange();
    var targetData = targetRange.getValues();
    var targetFormats = targetRange.getBackgrounds();
    var targetNotes = targetRange.getNotes();
    
    previewSheet.getRange(1, 1, targetData.length, targetData[0].length)
      .setValues(targetData)
      .setBackgrounds(targetFormats)
      .setNotes(targetNotes);

    // 执行合并预览
    var result = mergeSheets(config, previewSheet);
    
    if (result.success) {
      const previewData = {
        success: true,
        previewSheetName: previewSheet.getName(),
        sourceSheet: config.sourceSheet,
        targetSheet: config.targetSheet,
        changes: result.changes
      };
      savePreviewState(previewData);
      return {
        ...previewData,
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
