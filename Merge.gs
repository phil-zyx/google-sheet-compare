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
    // 获取三个表的数据：源表、目标表和基准表（原值）
    var sourceRange = sourceSheet.getDataRange();
    var targetRange = targetSheet.getDataRange();
    
    var sourceData = sourceRange.getValues();
    var targetData = targetRange.getValues();
    
    // 获取基准数据（从注释中）
    var baseData = targetRange.getNotes();
    
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
      
      // 为新行添加原值注释
      var newRowsNotes = changes.newRows.map(row => 
        row.map(value => addBaseValue(value))
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
        range.setNote(addBaseValue(col.sourceValue));
      });
    });

    // 3. 标记冲突
    changes.conflicts.forEach(conflict => {
      conflict.columns.forEach(col => {
        var targetIndex = headerMap[col.header].targetIndex;
        var range = targetSheet.getRange(conflict.targetRowIndex + 1, targetIndex + 1);
        range.setBackground(MERGE_CONSTANTS.COLORS.CONFLICT);
        range.setNote(`${MERGE_CONSTANTS.CONFLICT_PREFIX}当前值: ${col.targetValue}\n源表值: ${col.sourceValue}\n原值: ${col.baseValue}`);
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

// 从注释中提取原值
function extractBaseValue(note) {
  if (!note) return null;
  const basePrefix = "原值: ";
  const lines = note.split('\n');
  for (const line of lines) {
    if (line.startsWith(basePrefix)) {
      return line.substring(basePrefix.length);
    }
  }
  return null;
}

// 添加原值到注释
function addBaseValue(value) {
  return `原值: ${value}`;
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

    // 1. 清空目标表
    targetSheet.clear();
    
    // 2. 复制预览表的所有内容到目标表
    var previewRange = previewSheet.getDataRange();
    var previewData = previewRange.getValues();
    var previewFormats = previewRange.getBackgrounds();
    var previewNotes = previewRange.getNotes();
    
    targetSheet.getRange(1, 1, previewData.length, previewData[0].length)
      .setValues(previewData)
      .setBackgrounds(previewFormats)
      .setNotes(previewNotes);
    
    // 3. 标记源表为已合并
    var headerRow = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    var statusColIndex = -1;
    
    // 查找或创建状态列
    headerRow.forEach((header, index) => {
      if (header === '合并状态') {
        statusColIndex = index;
      }
    });
    
    if (statusColIndex === -1) {
      // 如果没有状态列，添加一个
      statusColIndex = headerRow.length;
      sourceSheet.getRange(1, statusColIndex + 1).setValue('合并状态');
    }
    
    // 标记所有数据行为已合并
    var lastRow = sourceSheet.getLastRow();
    if (lastRow > 1) {
      var statusRange = sourceSheet.getRange(2, statusColIndex + 1, lastRow - 1, 1);
      var statusValues = new Array(lastRow - 1).fill(['已合并']);
      statusRange.setValues(statusValues)
                .setBackground('#e8f5e9')  // 浅绿色背景
                .setFontColor('#2e7d32');  // 深绿色文字
    }
    
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
      message: "合并完成！源表已标记为已合并状态。"
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
  console.log('保存预览状态:', state);
  cache.put('merge_preview_state', JSON.stringify(state), 3600); // 1小时过期
}

/**
 * 获取预览状态
 * @returns {Object|null} 预览状态对象，如果没有则返回null
 */
function getPreviewState() {
  const cache = CacheService.getScriptCache();
  const state = cache.get('merge_preview_state');
  console.log('获取预览状态缓存:', state);
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
      cache.remove('merge_preview_state');
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
 * 显示确认对话框
 */
function showConfirmDialog() {
  // 显示合并对话框，对话框会自动检查预览状态
  showDialog('merge');
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
    var previewSheet = createPreviewSheet(config.targetSheet);
    
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
    
    // 清除冲突标记
    cell.setBackground(MERGE_CONSTANTS.COLORS.RESOLVED);
    cell.clearNote();
    
    // 检查该行是否还有其他冲突
    var rowRange = sheet.getRange(config.row + 1, 1, 1, headers.length);
    var backgrounds = rowRange.getBackgrounds()[0];
    var hasMoreConflicts = backgrounds.some(bg => bg === MERGE_CONSTANTS.COLORS.CONFLICT);
    
    // 如果没有更多冲突，将整行标记为已解决
    if (!hasMoreConflicts) {
      rowRange.setBackground(MERGE_CONSTANTS.COLORS.RESOLVED);
    }

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
 * @returns {Sheet} 预览表格
 */
function createPreviewSheet(targetSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var previewName = `预览_${targetSheetName}_${new Date().getTime()}`;
  var existingSheet = ss.getSheetByName(previewName);
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }
  return ss.insertSheet(previewName);
}
