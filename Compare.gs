// 获取所有表格名称
function getSheetNames() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  return sheets.map(function(sheet) {
    return sheet.getName();
  });
}

// 获取表格信息
function getSheetInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet().getName();
  var sheets = ss.getSheets().map(function(sheet) {
    return sheet.getName();
  });
  
  return {
    sheets: sheets,
    activeSheet: activeSheet
  };
}

// 创建自定义菜单
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('表格比较工具')
    .addItem('比较差异', 'showCompareDialog')
    .addItem('清除所有标记', 'clearAllHighlights')
    .addToUi();
}

// 显示配置对话框
function showCompareDialog() {
  var html = HtmlService.createHtmlOutputFromFile('CompareDialog')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '表格比较配置');
}

// 执行比较操作
function compareSheets(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName(config.sheet1);
  var sheet2 = ss.getSheetByName(config.sheet2);
  
  const COLORS = {
    MODIFIED: "#ffcdd2",  // 红色 - 修改
    ADDED: "#c8e6c9",     // 绿色 - 新增
    HEADER_MODIFIED: "#fff9c4", // 黄色 - 表头变更
  };
  
  if (!sheet1 || !sheet2) {
    return {
      success: false,
      message: "未找到指定的表格，请检查表格名称"
    };
  }

  try {
    // 获取两个表格的数据范围
    var range1 = sheet1.getDataRange();
    var range2 = sheet2.getDataRange();
    
    // 获取表头数据
    var headers1 = sheet1.getRange(1, 1, 1, range1.getNumColumns()).getValues()[0];
    var headers2 = sheet2.getRange(1, 1, 1, range2.getNumColumns()).getValues()[0];
    
    // 创建表头映射关系
    var headerMap = {};
    var unmatchedHeaders1 = [];
    var unmatchedHeaders2 = [];
    
    // 记录表头1中的列索引
    headers1.forEach((header, index) => {
      headerMap[header] = {
        sheet1Index: index,
        sheet2Index: -1
      };
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
    var backgrounds = range1.getBackgrounds();
    var notes = range1.getNotes();
    var hasExistingMarks = false;

    // 检查是否存在比较标记
    for (var i = 0; i < backgrounds.length && !hasExistingMarks; i++) {
      for (var j = 0; j < backgrounds[i].length && !hasExistingMarks; j++) {
        if (backgrounds[i][j] === COLORS.MODIFIED || backgrounds[i][j] === COLORS.ADDED ||
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

    // 获取数据
    var data1 = sheet1.getDataRange().getValues();
    var data2 = sheet2.getDataRange().getValues();
    
    // 获取当前背景色
    var currentBackgrounds = sheet1.getDataRange().getBackgrounds();
    
    var differences = {
      total: 0,
      modified: 0,
      missing: 0,
      added: 0,
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
        headerBackgrounds.push(COLORS.HEADER_MODIFIED);
        headerNotes.push("此列在对比表格中不存在");
      } else {
        headerBackgrounds.push(currentBackgrounds[0][j]);
        headerNotes.push(null);
      }
    }
    backgroundColors.push(headerBackgrounds);
    notes.push(headerNotes);

    // 处理数据行
    for(var i = 1; i < rows; i++) {
      var backgroundRow = [];
      var noteRow = [];
      for(var j = 0; j < cols; j++) {
        var header = headers1[j];
        var sheet2Col = headerMap[header].sheet2Index;
        
        if (sheet2Col === -1) {
          // 此列在表格2中不存在
          backgroundRow.push(COLORS.ADDED);
          noteRow.push("此列在对比表格中不存在");
          differences.added++;
          differences.total++;
        } else if (i < data2.length) {
          if (data1[i][j] !== data2[i][sheet2Col]) {
            backgroundRow.push(COLORS.MODIFIED);
            noteRow.push("当前值: " + data1[i][j] + "\n对比值: " + data2[i][sheet2Col]);
            differences.modified++;
            differences.total++;
          } else {
            backgroundRow.push(currentBackgrounds[i][j]);
            noteRow.push(null);
          }
        } else {
          backgroundRow.push(COLORS.ADDED);
          noteRow.push("在对比表格中未找到对应数据");
          differences.added++;
          differences.total++;
        }
      }
      backgroundColors.push(backgroundRow);
      notes.push(noteRow);
    }

    // 批量更新单元格背景色和注释
    var range = sheet1.getRange(1, 1, rows, cols);
    range.setBackgrounds(backgroundColors);
    range.setNotes(notes);

    // 更新比较信息
    var infoNote = "最近比较时间: " + new Date().toLocaleString() + "\n" +
                   "对比表格: " + config.sheet2 + "\n" +
                   "差异总数: " + differences.total + "\n" +
                   "└─ 值不同: " + differences.modified + "\n" +
                   "└─ 新增项: " + differences.added + "\n" +
                   "└─ 表头差异: " + differences.headerDiff;
    
    sheet1.getRange(1, 1).setNote(infoNote);

    return {
      success: true,
      message: `比较完成！\n发现 ${differences.total} 处差异\n${differences.modified} 处值不同\n${differences.added} 处新增\n${differences.headerDiff} 处表头差异`,
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
  var rows = range.getNumRows();
  var cols = range.getNumColumns();
  
  // 创建新的背景色和注释数组
  var newBackgrounds = [];
  var newNotes = [];
  
  // 遍历所有单元格
  for (var i = 0; i < rows; i++) {
    var backgroundRow = [];
    var noteRow = [];
    for (var j = 0; j < cols; j++) {
      var currentBg = backgrounds[i][j];
      var currentNote = notes[i][j];
      
      // 清除比较功能产生的所有标记颜色（包括表头差异的黄色标记）
      if (currentBg === "#ffcdd2" || currentBg === "#c8e6c9" || currentBg === "#fff9c4") {
        backgroundRow.push(null);
      } else {
        backgroundRow.push(currentBg);
      }
      
      // 清除包含比较相关文字的注释
      if (currentNote && (currentNote.includes("当前值:") || 
          currentNote.includes("对比值:") || 
          currentNote.includes("在对比表格中未找到对应数据") ||
          currentNote.includes("最近比较时间:") ||
          currentNote.includes("此列在对比表格中不存在"))) {  // 添加表头差异的注释判断
        noteRow.push("");
      } else {
        noteRow.push(currentNote);
      }
    }
    newBackgrounds.push(backgroundRow);
    newNotes.push(noteRow);
  }
  
  // 批量更新
  range.setBackgrounds(newBackgrounds);
  range.setNotes(newNotes);
  
  if (showConfirm) {
    ui.alert('完成', '比较标记已清除', ui.ButtonSet.OK);
  }
}

// 获取表格信息
function getSheetInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet().getName();
  var sheets = ss.getSheets().map(function(sheet) {
    return sheet.getName();
  });
  
  return {
    sheets: sheets,
    activeSheet: activeSheet
  };
}