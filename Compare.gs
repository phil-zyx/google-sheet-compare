/**
 * 显示合并对话框
 */
function showMergeDialog() {
  var html = HtmlService.createHtmlOutputFromFile('MergeDialog')
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '表格合并工具');
}

// 显示配置对话框
function showCompareDialog() {
  // 每次打开对话框时清除缓存，确保获取最新数据
  var cache = CacheService.getScriptCache();
  cache.remove('sheet_info');
  
  var html = HtmlService.createHtmlOutputFromFile('CompareDialog')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '表格比较配置');
}

// 执行比较操作 - 优化数据处理
function compareSheets(config) {
  // 应该添加输入验证
  if (!config || !config.sheet1 || !config.sheet2) {
    return {
      success: false,
      message: "配置参数无效"
    };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName(config.sheet1);
  var sheet2 = ss.getSheetByName(config.sheet2);
  
  if (!sheet1 || !sheet2) {
    return {
      success: false,
      message: "未找到指定的表格，请检查表格名称"
    };
  }

  try {
    // 批量获取数据以减少API调用
    var range1 = sheet1.getDataRange();
    var range2 = sheet2.getDataRange();
    
    var [data1, backgrounds1, notes1] = [
      range1.getValues(),
      range1.getBackgrounds(),
      range1.getNotes()
    ];
    
    var data2 = range2.getValues();
    
    // 获取表头数据
    var headers1 = sheet1.getRange(1, 1, 1, range1.getNumColumns()).getValues()[0];
    var headers2 = sheet2.getRange(1, 1, 1, range2.getNumColumns()).getValues()[0];
    
    // 使用对象来存储表头映射，提高查找效率
    var headerMap = {};
    var unmatchedHeaders1 = [];
    var unmatchedHeaders2 = [];
    
    // 记录表头1中的列索引
    headers1.forEach((header, index) => {
      headerMap[header] = {sheet1Index: index, sheet2Index: -1};
    });
    
    // 查找表头2中对应的列索引
    headers2.forEach((header, index) => {
      if (headerMap[header]) {
        headerMap[header].sheet2Index = index;
      } else {
        unmatchedHeaders2.push({header: header, index: index});
      }
    });
    
    // 找出表头1中未匹配的列
    headers1.forEach((header, index) => {
      if (headerMap[header].sheet2Index === -1) {
        unmatchedHeaders1.push({header: header, index: index});
      }
    });

    // 如果有表头差异，显示详细信息
    if (unmatchedHeaders1.length > 0 || unmatchedHeaders2.length > 0) {
      var warningMessage = "检测到表头差异：\n";
      if (unmatchedHeaders1.length > 0) {
        warningMessage += "\n当前表格独有的列：\n" + 
          unmatchedHeaders1.map(h => `- ${h.header}`).join("\n");
      }
      if (unmatchedHeaders2.length > 0) {
        warningMessage += "\n对比表格独有的列：\n" + 
          unmatchedHeaders2.map(h => `- ${h.header}`).join("\n");
      }
      warningMessage += "\n\n系统将只比较两个表格中都存在的列，是否继续？";
      
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert('表头差异提醒', warningMessage, ui.ButtonSet.YES_NO);
      
      if (response !== ui.Button.YES) {
        return {
          success: false,
          message: "操作已取消"
        };
      }
    }

    // 检查是否存在之前的比较标记
    var backgrounds = backgrounds1;
    var notes = notes1;
    var hasExistingMarks = false;

    // 检查是否存在比较标记
    for (var i = 0; i < backgrounds.length && !hasExistingMarks; i++) {
      for (var j = 0; j < backgrounds[i].length && !hasExistingMarks; j++) {
        if (backgrounds[i][j] === COMPARE_CONSTANTS.COLORS.MODIFIED || backgrounds[i][j] === COMPARE_CONSTANTS.COLORS.ADDED ||
            (notes[i][j] && (notes[i][j].includes("当前值:") || 
             notes[i][j].includes("对比值:") || 
             notes[i][j].includes("在对比表中未找到对应数据") ||
             notes[i][j].includes("最近比较时间:")))) {
          hasExistingMarks = true;
        }
      }
    }

    // 如果存在标记，提示用户
    if (hasExistingMarks) {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert(
        '发现已有比较标记',
        '当前表格中存在之前的比较标记，是否清除后继续？',
        ui.ButtonSet.YES_NO
      );

      if (response !== ui.Button.YES) {
        return {
          success: false,
          message: "操作已取消"
        };
      }

      // 清除现有标记
      clearAllHighlights(false); // 传入 false 表示不显示确认对话框
    }

    // 获取当前背景色
    var currentBackgrounds = sheet1.getDataRange().getBackgrounds();
    
    var differences = {
      total: 0,
      modified: 0,
      missing: 0,
      added: 0,
      removed: 0,  // 新增：记录删除的行数
      headerDiff: unmatchedHeaders1.length + unmatchedHeaders2.length
    };

    // 准备批量更新数组
    var backgroundColors = [];
    var notes = [];
    var rows = data1.length;
    var cols = headers1.length;

    // 处理表头行
    var headerBackgrounds = [];
    var headerNotes = [];
    for (var j = 0; j < cols; j++) {
      if (headerMap[headers1[j]].sheet2Index === -1) {
        headerBackgrounds.push(COMPARE_CONSTANTS.COLORS.HEADER_MODIFIED);
        headerNotes.push("此列在对比表格中不存在");
      } else {
        headerBackgrounds.push(currentBackgrounds[0][j]);
        headerNotes.push(null);
      }
    }
    backgroundColors.push(headerBackgrounds);
    notes.push(headerNotes);

    // 记录已处理的行（用于后续检查删除的行）
    var processedRows = new Set();

    // 处理数据行
    for(var i = 1; i < rows; i++) {
      var backgroundRow = [];
      var noteRow = [];
      
      // 记录这一行的内容用于后续比较
      var rowContent = '';
      for(var j = 0; j < cols; j++) {
        var sheet2Col = headerMap[headers1[j]].sheet2Index;
        if (sheet2Col !== -1) {
          rowContent += data1[i][j] + '|';
        }
      }
      processedRows.add(rowContent);

      for(var j = 0; j < cols; j++) {
        var header = headers1[j];
        var sheet2Col = headerMap[header].sheet2Index;
        
        if (sheet2Col === -1) {
          // 此列在表格2中不存在
          backgroundRow.push(COMPARE_CONSTANTS.COLORS.ADDED);
          noteRow.push("此列在对比表格中不存在");
          differences.added++;
          differences.total++;
        } else if (i < data2.length) {
          if (data1[i][j] !== data2[i][sheet2Col]) {
            backgroundRow.push(COMPARE_CONSTANTS.COLORS.MODIFIED);
            noteRow.push("当前值: " + data1[i][j] + "\n对比值: " + data2[i][sheet2Col]);
            differences.modified++;
            differences.total++;
          } else {
            backgroundRow.push(currentBackgrounds[i][j]);
            noteRow.push(null);
          }
        } else {
          backgroundRow.push(COMPARE_CONSTANTS.COLORS.ADDED);
          noteRow.push("在对比表格中未找到对应数据");
          differences.added++;
          differences.total++;
        }
      }
      backgroundColors.push(backgroundRow);
      notes.push(noteRow);
    }

    // 检查B表中存在而A表中不存在的行
    for(var i = 1; i < data2.length; i++) {
      var rowContent = '';
      for(var j = 0; j < data2[i].length; j++) {
        rowContent += data2[i][j] + '|';
      }
      
      if (!processedRows.has(rowContent)) {
        differences.removed++;
        differences.total++;
        
        // 将删除的行添加到结果中
        var removedRow = [];
        var removedNote = [];
        var dataRow = new Array(cols).fill('');
        
        for(var j = 0; j < cols; j++) {
          removedRow.push(COMPARE_CONSTANTS.COLORS.REMOVED);  // 使用删除专用的颜色
          var sheet2Col = headerMap[headers1[j]].sheet2Index;
          if (sheet2Col !== -1) {
            var value = data2[i][sheet2Col];
            dataRow[j] = value;
            removedNote.push("此行在基准表中不存在\n对比表中的值: " + value);
          } else {
            removedNote.push("此行在基准表中不存在");
          }
        }
        
        backgroundColors.push(removedRow);
        notes.push(removedNote);
        data1.push(dataRow);
        rows++;
      }
    }

    // 批量更新以减少API调用
    var updateRange = sheet1.getRange(1, 1, rows, cols);
    updateRange.setBackgrounds(backgroundColors);
    updateRange.setNotes(notes);
    updateRange.setValues(data1);  // 更新包括新添加的行

    // 更新比较信息
    var infoNote = "最近比较时间: " + new Date().toLocaleString() + "\n" +
                   "对比表格: " + config.sheet2 + "\n" +
                   "差异总数: " + differences.total + "\n" +
                   "└─ 值不同: " + differences.modified + "\n" +
                   "└─ 新增项: " + differences.added + "\n" +
                   "└─ 删除项: " + differences.removed + "\n" +
                   "└─ 表头差异: " + differences.headerDiff;
    
    sheet1.getRange(1, 1).setNote(infoNote);

    return {
      success: true,
      message: `比较完成！\n发现 ${differences.total} 处差异\n${differences.modified} 处值不同\n${differences.added} 处新增\n${differences.removed} 处删除\n${differences.headerDiff} 处表头差异`,
      differences: differences
    };
  } catch (error) {
    return {
      success: false,
      message: "发生错误: " + error.toString()
    };
  }
}

// 清除所有高亮和注释
function clearAllHighlights(showConfirm = true) {
  var ui = SpreadsheetApp.getUi();
  
  if (showConfirm) {
    var response = ui.alert(
      '确认清除',
      '是否要清除当前表格中所有的比较标记？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }
  }
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var backgrounds = range.getBackgrounds();
  var notes = range.getNotes();
  var values = range.getValues();
  
  var newBackgrounds = [];
  var newNotes = [];
  var newValues = [];
  var rowsToKeep = [];
  
  // 检查每一行，标记需要保留的行
  for (var i = 0; i < backgrounds.length; i++) {
    var isDeletedRow = true;
    var hasHighlight = false;
    
    // 检查这一行是否是比较时新增的行（通过检查背景色和注释）
    for (var j = 0; j < backgrounds[i].length; j++) {
      var currentBg = backgrounds[i][j];
      var currentNote = notes[i][j];
      
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
      
      var backgroundRow = [];
      var noteRow = [];
      
      for (var j = 0; j < backgrounds[i].length; j++) {
        var currentBg = backgrounds[i][j];
        var currentNote = notes[i][j];
        
        // 清除所有比较标记的背景色
        if (currentBg === COMPARE_CONSTANTS.COLORS.MODIFIED || 
            currentBg === COMPARE_CONSTANTS.COLORS.ADDED || 
            currentBg === COMPARE_CONSTANTS.COLORS.REMOVED || 
            currentBg === COMPARE_CONSTANTS.COLORS.HEADER_MODIFIED) {
          backgroundRow.push(null);
        } else {
          backgroundRow.push(currentBg);
        }
        
        // 清除所有比较相关的注释
        if (currentNote && (
            currentNote.includes("当前值:") || 
            currentNote.includes("对比值:") || 
            currentNote.includes("在对比表格中未找到对应数据") ||
            currentNote.includes("最近比较时间:") ||
            currentNote.includes("此列在对比表格中不存在") ||
            currentNote.includes("此行在基准表中不存在"))) {
          noteRow.push("");
        } else {
          noteRow.push(currentNote);
        }
      }
      
      newBackgrounds.push(backgroundRow);
      newNotes.push(noteRow);
      newValues.push(values[i]);
    }
  }
  
  // 如果有行被删除，更新表格
  if (rowsToKeep.length < backgrounds.length) {
    var newRange = sheet.getRange(1, 1, newBackgrounds.length, backgrounds[0].length);
    newRange.setBackgrounds(newBackgrounds);
    newRange.setNotes(newNotes);
    newRange.setValues(newValues);
    
    // 删除多余的行
    if (backgrounds.length > newBackgrounds.length) {
      sheet.deleteRows(newBackgrounds.length + 1, backgrounds.length - newBackgrounds.length);
    }
  } else {
    // 如果没有行被删除，只更新背景色和注释
    range.setBackgrounds(newBackgrounds);
    range.setNotes(newNotes);
  }
}

/**
 * 基于当前表格创建新的页签
 */
function createNewSheetTab() {
  console.log('开始创建新页签');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();
  console.log('当前页签:', currentSheet.getName());
  
  // 弹出对话框让用户输入新页签名称
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '新建页签',
    '请输入新页签名称：\n(将基于当前页签 "' + currentSheet.getName() + '" 创建)',
    ui.ButtonSet.OK_CANCEL
  );

  // 处理用户输入
  if (response.getSelectedButton() == ui.Button.OK) {
    var newSheetName = response.getResponseText().trim();
    
    // 验证输入的名称
    if (newSheetName === '') {
      ui.alert('错误', '页签名称不能为空', ui.ButtonSet.OK);
      return;
    }
    
    // 检查是否已存在同名页签
    if (ss.getSheetByName(newSheetName)) {
      ui.alert('错误', '已存在同名页签："' + newSheetName + '"', ui.ButtonSet.OK);
      return;
    }
    
    try {
      // 复制当前页签
      var newSheet = currentSheet.copyTo(ss);
      newSheet.setName(newSheetName);
      
      // 将新页签移动到当前页签后面
      var sheets = ss.getSheets();
      var currentIndex = sheets.findIndex(function(sheet) {
        return sheet.getName() === currentSheet.getName();
      });
      ss.setActiveSheet(newSheet);
      console.log('切换到新页签:', newSheet.getName());
      ss.moveActiveSheet(currentIndex + 2);
      console.log('移动页签完成');
      
      // 创建用户友好的创建信息
      var creationInfo = {
        '创建者': Session.getActiveUser().getEmail(),
        '创建时间': new Date().toLocaleString(),
        '来源页签': currentSheet.getName()
      };
      
      var a1Cell = newSheet.getRange("A1");
      var currentNote = a1Cell.getNote() || '';
      var newNote = NoteManager.addSystemNote(
        currentNote,
        NOTE_CONSTANTS.TYPES.SHEET_CREATION,
        Object.entries(creationInfo)
          .map(([key, value]) => `${key}：${value}`)
          .join('\n')
      );
      a1Cell.setNote(newNote);
      console.log('添加创建记录到A1单元格');
      
      // 清除缓存以确保getSheetInfo()返回最新数据
      var cache = CacheService.getScriptCache();
      cache.remove('sheet_info');
      console.log('清除sheet_info缓存');
      
      ui.alert('成功', '已创建新页签："' + newSheetName + '"', ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('错误', '创建页签失败：' + error.toString(), ui.ButtonSet.OK);
    }
  }
}