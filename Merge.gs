// 合并相关的常量定义
const MERGE_CONSTANTS = {
  COLORS: {
    AUTO_MERGED: "#e3f2fd",    // 自动合并 - 浅蓝色
    CONFLICT: "#fff59d",       // 冲突 - 浅黄色
    RESOLVED: "#c8e6c9"        // 已解决 - 浅绿色
  },
  CONFLICT_PREFIX: "💡 冲突: "
};

/**
 * 执行合并操作
 * @param {Object} config 合并配置
 * @returns {Object} 合并结果
 */
function mergeSheets(config) {
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
    // 获取数据范围
    var sourceRange = sourceSheet.getDataRange();
    var targetRange = targetSheet.getDataRange();
    var sourceData = sourceRange.getValues();
    var targetData = targetRange.getValues();
    
    // 执行合并分析
    var mergeResult = analyzeMergeChanges(sourceData, targetData);
    
    // 标记合并结果
    highlightMergeResults(targetSheet, mergeResult);
    
    return {
      success: true,
      message: `合并分析完成。发现 ${mergeResult.autoMerged.length} 个可自动合并项，${mergeResult.conflicts.length} 个冲突需要解决。`,
      result: mergeResult
    };
  } catch (error) {
    return {
      success: false,
      message: "合并过程发生错误: " + error.toString()
    };
  }
}

/**
 * 分析合并变更
 * @param {Array} sourceData 源数据
 * @param {Array} targetData 目标数据
 * @returns {Object} 合并分析结果
 */
function analyzeMergeChanges(sourceData, targetData) {
  var result = {
    autoMerged: [],
    conflicts: [],
    final: []
  };
  
  // 获取最大行列数
  var maxRows = Math.max(sourceData.length, targetData.length);
  var maxCols = Math.max(
    sourceData[0] ? sourceData[0].length : 0,
    targetData[0] ? targetData[0].length : 0
  );
  
  // 初始化final数组
  for (var i = 0; i < maxRows; i++) {
    result.final[i] = [];
    for (var j = 0; j < maxCols; j++) {
      var sourceValue = sourceData[i] && sourceData[i][j] !== undefined ? sourceData[i][j] : "";
      var targetValue = targetData[i] && targetData[i][j] !== undefined ? targetData[i][j] : "";
      
      if (sourceValue === targetValue) {
        // 值相同，直接使用
        result.final[i][j] = sourceValue;
      } else if (sourceValue === "" && targetValue !== "") {
        // 目标有值而源为空，保留目标值
        result.final[i][j] = targetValue;
        result.autoMerged.push({row: i, col: j, value: targetValue});
      } else if (sourceValue !== "" && targetValue === "") {
        // 源有值而目标为空，使用源值
        result.final[i][j] = sourceValue;
        result.autoMerged.push({row: i, col: j, value: sourceValue});
      } else {
        // 冲突情况
        result.final[i][j] = targetValue;
        result.conflicts.push({
          row: i,
          col: j,
          sourceValue: sourceValue,
          targetValue: targetValue
        });
      }
    }
  }
  
  return result;
}

/**
 * 高亮显示合并结果
 * @param {Sheet} sheet 目标表格
 * @param {Object} mergeResult 合并结果
 */
function highlightMergeResults(sheet, mergeResult) {
  // 高亮自动合并的单元格
  mergeResult.autoMerged.forEach(function(item) {
    var cell = sheet.getRange(item.row + 1, item.col + 1);
    cell.setBackground(MERGE_CONSTANTS.COLORS.AUTO_MERGED);
    cell.setValue(item.value);
  });
  
  // 高亮冲突的单元格
  mergeResult.conflicts.forEach(function(item) {
    var cell = sheet.getRange(item.row + 1, item.col + 1);
    cell.setBackground(MERGE_CONSTANTS.COLORS.CONFLICT);
    cell.setNote(MERGE_CONSTANTS.CONFLICT_PREFIX + 
                `源值: ${item.sourceValue}\n` +
                `目标值: ${item.targetValue}`);
  });
}

/**
 * 解决指定的冲突
 * @param {Object} config 解决配置
 * @returns {Object} 操作结果
 */
function resolveConflict(config) {
  if (!config || !config.sheet || !config.row || !config.col || !config.value) {
    return {
      success: false,
      message: "解决冲突的参数无效"
    };
  }
  
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheet);
    if (!sheet) {
      return {
        success: false,
        message: "未找到指定的表格"
      };
    }
    
    var cell = sheet.getRange(config.row, config.col);
    cell.setValue(config.value);
    cell.setBackground(MERGE_CONSTANTS.COLORS.RESOLVED);
    cell.clearNote();
    
    return {
      success: true,
      message: "冲突已解决"
    };
  } catch (error) {
    return {
      success: false,
      message: "解决冲突时发生错误: " + error.toString()
    };
  }
}
