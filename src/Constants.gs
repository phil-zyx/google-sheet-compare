// --------------------- 常量 --------------------------------

// 表格相关常量
const SHEET_CONSTANTS = {
  COLORS: {
    MODIFIED: "#b3e5fc",  // 修改 - 浅蓝色
    ADDED: "#dcedc8",    // 新增 - 淡绿色
    HEADER_MODIFIED: "#fff9c4",  // 表头修改 - 浅黄色
    CONFLICT: "#f8bbd0"  // 冲突 - 粉色
  }
};

// 合并相关常量
const MERGE_CONSTANTS = {
  ID_SUFFIX: '_INT_id',
  CONFLICT_PREFIX: '冲突: ',
  PREVIEW_SUFFIX: '_预览',
  COLORS: {
    NEW: '#dcedc8',      // 浅绿色 - 新行
    CONFLICT: '#ffdce0', // 浅红色 - 冲突
    UPDATED: '#b3e5fc',  // 浅蓝色 - 已更新
    RESOLVED: "#e8f5e9", // 已解决 - 更浅的绿色
    MERGED: "#dfcd4d"    // 合并入表 - 橘色
  }
};

// 比较相关常量
const COMPARE_CONSTANTS = {
  COLORS: {
    MODIFIED: "#ffcdd2",  // 修改 - 浅红色
    ADDED: "#dcedc8",    // 新增 - 浅绿色
    REMOVED: "#ffdce0",  // 删除 - Git风格浅红色
    HEADER_MODIFIED: "#fff9c4"  // 表头修改 - 浅黄色
  }
};

// ID检查器相关常量
const ID_CHECKER_CONFIG = {
  COLORS: {
    CONFLICT: '#ff0000',  // 冲突标记颜色 - 红色
  },
  ID_COLUMN_SUFFIX: '_INT_id',   // ID列的后缀
};

// 注释相关常量
const NOTE_CONSTANTS = {
  // 系统注释使用键值对格式
  SYSTEM_NOTE_START: '===== 系统信息开始 =====\n',
  SYSTEM_NOTE_END: '\n===== 系统信息结束 =====',
  
  TYPES: {
    BASE_VALUE: 'BASE',     // 基准值
    CONFLICT: 'CONFLICT',   // 冲突信息
    MERGE_INFO: 'MERGE',    // 合并信息
    VERSION: 'VERSION',     // 版本信息
    SHEET_CREATION: 'CREATION'  // 页签创建信息
  },

  // 添加分隔符常量
  KEY_VALUE_SEPARATOR: ': ',  // 键值分隔符
  LINE_SEPARATOR: '\n'        // 行分隔符
};

// 日志相关常量
const LOG_CONSTANTS = {
  SHEET_NAME: "配置表工具操作日志表",
  HEADERS: [
    "时间",
    "操作类型",
    "操作人",
    "操作表名",
    "操作内容",
    "详细信息"
  ],
  TYPES: {
    MERGE: "合并操作",
    COMPARE: "比较操作",
    CONFLICT_RESOLVE: "冲突解决",
    SHEET_CREATE: "创建表格",
    SHEET_UPDATE: "更新表格",
    SHEET_DELETE: "删除表格"
  },
  // 日志保留配置
  RETENTION: {
    MAX_ROWS: 10000,        // 最大保留行数
    CLEANUP_THRESHOLD: 0.9,  // 清理阈值（当达到最大行数的90%时触发清理）
    CLEANUP_TARGET: 0.7,     // 清理目标（清理后保留最大行数的70%）
    MIN_DAYS: 30            // 最小保留天数（无论行数多少，30天内的日志都保留）
  }
};

// --------------------- 常量 --------------------------------

// --------------------- NoteManager ------------------------ 

/**
 * 注释管理工具类
 */
class NoteManager {
  /**
   * 提取系统注释，自动合并多个系统注释块
   */
  static extractSystemNotes(note) {
    if (!note) return {};
    
    const systemNotes = {};
    let currentPosition = 0;
    let hasMultipleBlocks = false;
    
    // 查找所有系统注释块并合并
    while (true) {
      const start = note.indexOf(NOTE_CONSTANTS.SYSTEM_NOTE_START, currentPosition);
      if (start === -1) break;
      
      const end = note.indexOf(NOTE_CONSTANTS.SYSTEM_NOTE_END, start);
      if (end === -1) break;
      
      // 如果不是第一个块，标记存在多个块
      if (currentPosition > 0) {
        hasMultipleBlocks = true;
      }
      
      const notesSection = note.substring(
        start + NOTE_CONSTANTS.SYSTEM_NOTE_START.length,
        end
      );

      const noteRegex = /^(.+?):\s*\n([\s\S]*?)(?=\n\w+:|$)/gm;
      let match;
      
      while ((match = noteRegex.exec(notesSection)) !== null) {
        const [, key, value] = match;
        systemNotes[key.trim()] = value.trim();
      }
      
      currentPosition = end + NOTE_CONSTANTS.SYSTEM_NOTE_END.length;
    }
    
    // 如果发现多个块，自动清理并重写注释
    if (hasMultipleBlocks) {
      const cleanNote = this.removeAllSystemNotes(note);
      const systemPart = this.formatSystemNotes(systemNotes);
      const newNote = this.appendSystemNote(cleanNote, systemPart);
      
      // 如果是在单元格上下文中，尝试更新单元格注释
      try {
        const cell = SpreadsheetApp.getActiveRange();
        if (cell) {
          cell.setNote(newNote);
        }
      } catch (e) {
        // 忽略错误，因为可能不在单元格上下文中
      }
    }
    
    return systemNotes;
  }

  /**
   * 添加系统注释
   */
  static addSystemNote(originalNote, type, content) {
    // 添加参数验证
    if (!type || content === undefined) {
      throw new Error('Type and content are required');
    }
    
    // 直接使用 extractSystemNotes 进行合并处理
    const systemNotes = this.extractSystemNotes(originalNote);
    systemNotes[type] = content;
    
    const systemPart = this.formatSystemNotes(systemNotes);
    return this.appendSystemNote(this.removeAllSystemNotes(originalNote), systemPart);
  }
  
  /**
   * 获取系统注释内容
   * @param {string} note 完整注释
   * @param {string} type 注释类型
   * @returns {string|null} 系统注释内容
   */
  static getSystemNote(note, type) {
    const systemNotes = this.extractSystemNotes(note);
    return systemNotes[type] || null;
  }
  
  /**
   * 移除指定类型的系统注释
   * @param {string} note 完整注释
   * @param {string} type 注释类型
   * @returns {string} 清理后的注释
   */
  static removeSystemNote(note, type) {
    const systemNotes = this.extractSystemNotes(note);
    delete systemNotes[type];
    
    // 如果没有剩余的系统注释，返回清理后的原始注释
    if (Object.keys(systemNotes).length === 0) {
      return this.removeAllSystemNotes(note);
    }
    
    const systemPart = this.formatSystemNotes(systemNotes);
    return this.appendSystemNote(this.removeAllSystemNotes(note), systemPart);
  }
  
  /**
   * 格式化系统注释
   * @private
   * @param {Object} systemNotes 系统注释对象
   * @returns {string} 格式化后的系统注释
   */
  static formatSystemNotes(systemNotes) {
    if (Object.keys(systemNotes).length === 0) return '';
    
    const formattedNotes = Object.entries(systemNotes)
      .map(([key, value]) => `${key}${NOTE_CONSTANTS.KEY_VALUE_SEPARATOR}\n${value}`)
      .join(NOTE_CONSTANTS.LINE_SEPARATOR);
    
    return `${NOTE_CONSTANTS.SYSTEM_NOTE_START}${formattedNotes}${NOTE_CONSTANTS.SYSTEM_NOTE_END}`;
  }
  
  /**
   * 移除所有系统注释
   * @private
   * @param {string} note 完整注释
   * @returns {string} 移除系统注释后的原始注释
   */
  static removeAllSystemNotes(note) {
    if (!note) return '';
    
    const start = note.indexOf(NOTE_CONSTANTS.SYSTEM_NOTE_START);
    if (start === -1) return note;
    
    const end = note.indexOf(NOTE_CONSTANTS.SYSTEM_NOTE_END);
    if (end === -1) return note;
    
    return note.substring(0, start) + note.substring(end + NOTE_CONSTANTS.SYSTEM_NOTE_END.length);
  }
  
  /**
   * 在原始注释后追加系统注释
   * @private
   * @param {string} originalNote 原始注释
   * @param {string} systemNote 系统注释部分
   * @returns {string} 组合后的完整注释
   */
  static appendSystemNote(originalNote, systemNote) {
    if (!systemNote) return originalNote || '';
    if (!originalNote) return systemNote;
    
    return `${originalNote.trim()}\n${systemNote}`;
  }

  /**
   * 移除单元格中指定类型的标记
   * @param {Range} cell 目标单元格
   * @param {string} type 要移除的标记类型
   * @returns {boolean} 是否成功移除标记
   */
  static removeMarkFromCell(cell, type) {
    if (!cell) return false;
    
    const note = cell.getNote();
    if (!note) return false;
    
    const newNote = this.removeSystemNote(note, type);
    
    // 如果注释内容没有变化，说明没有找到对应类型的标记
    if (newNote === note) return false;
    
    // 如果新注释为空，则完全清除注释
    if (newNote.trim() === '') {
      cell.clearNote();
    } else {
      cell.setNote(newNote);
    }
    
    return true;
  }
} 

// --------------------- NoteManager ------------------------ 

// --------------------- LogManager ------------------------ 

/**
 * 日志管理工具类
 */
class LogManager {
  /**
   * 初始化日志表
   * @private
   * @returns {Sheet} 日志表对象
   */
  static _initLogSheet() {
    const ss = SpreadsheetApp.getActive();
    let sheet = ss.getSheetByName(LOG_CONSTANTS.SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(LOG_CONSTANTS.SHEET_NAME);
      sheet.getRange(1, 1, 1, LOG_CONSTANTS.HEADERS.length)
        .setValues([LOG_CONSTANTS.HEADERS])
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    return sheet;
  }

  /**
   * 添加日志记录
   * @param {string} type 操作类型（使用 LOG_CONSTANTS.TYPES 中的值）
   * @param {string} sheetName 操作的表名
   * @param {string} action 操作内容
   * @param {string} [details=''] 详细信息（可选）
   */
  static addLog(type, sheetName, action, details = '') {
    const sheet = this._initLogSheet();
    const user = Session.getActiveUser().getEmail();
    const timestamp = new Date().toLocaleString("zh-CN");
    
    const logRow = [
      timestamp,
      type,
      user,
      sheetName,
      action,
      details
    ];
    
    // 在第二行插入新日志（保持表头在第一行）
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, logRow.length).setValues([logRow]);

    // 检查是否需要清理日志
    this._checkAndCleanupLogs(sheet);
  }

  /**
   * 检查并清理日志
   * @private
   * @param {Sheet} sheet 日志表对象
   */
  static _checkAndCleanupLogs(sheet) {
    const currentRows = sheet.getLastRow();
    const threshold = LOG_CONSTANTS.RETENTION.MAX_ROWS * LOG_CONSTANTS.RETENTION.CLEANUP_THRESHOLD;
    
    // 如果当前行数超过阈值，触发清理
    if (currentRows > threshold) {
      const targetRows = Math.floor(LOG_CONSTANTS.RETENTION.MAX_ROWS * LOG_CONSTANTS.RETENTION.CLEANUP_TARGET);
      const data = sheet.getDataRange().getValues();
      
      // 确保保留表头
      if (data.length <= 1) return;
      
      // 计算最小保留日期
      const minDate = new Date();
      minDate.setDate(minDate.getDate() - LOG_CONSTANTS.RETENTION.MIN_DAYS);
      
      // 从后往前查找需要保留的最后一行
      let deleteFromRow = data.length;
      let foundDeleteRow = false;
      
      for (let i = data.length - 1; i > 1; i--) {
        const logDate = new Date(data[i][0]);
        
        // 如果找到了一行需要删除的数据（在最小保留日期之前，且超出目标行数）
        if (logDate < minDate && i > targetRows) {
          deleteFromRow = i;
          foundDeleteRow = true;
          break;
        }
      }
      
      // 如果需要删除行
      if (deleteFromRow < data.length) {
        const rowsToDelete = data.length - deleteFromRow;
        sheet.deleteRows(deleteFromRow + 1, rowsToDelete);
        
        // 记录清理操作（插入到第二行）
        const newLog = [
          new Date().toLocaleString("zh-CN"),
          "系统维护",
          "系统",
          LOG_CONSTANTS.SHEET_NAME,
          "日志清理",
          `清理了 ${rowsToDelete} 条历史日志记录`
        ];
        sheet.insertRowAfter(1);
        sheet.getRange(2, 1, 1, newLog.length).setValues([newLog]);
      }
    }
  }

  /**
   * 手动触发日志清理
   * @param {number} [days=30] 保留天数
   */
  static manualCleanup(days = LOG_CONSTANTS.RETENTION.MIN_DAYS) {
    const sheet = this._initLogSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return; // 只有表头或空表，直接返回
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    
    // 从后往前查找需要删除的行
    let deleteFromRow = data.length;
    for (let i = data.length - 1; i > 1; i--) {
      const logDate = new Date(data[i][0]);
      if (logDate < cutoffDate) {
        deleteFromRow = i;
        break;
      }
    }
    
    if (deleteFromRow < data.length) {
      const rowsToDelete = data.length - deleteFromRow;
      sheet.deleteRows(deleteFromRow + 1, rowsToDelete);
      
      // 记录清理操作（插入到第二行）
      const newLog = [
        new Date().toLocaleString("zh-CN"),
        "系统维护",
        "系统",
        LOG_CONSTANTS.SHEET_NAME,
        "手动日志清理",
        `清理了 ${rowsToDelete} 条${days}天前的历史日志记录`
      ];
      sheet.insertRowAfter(1);
      sheet.getRange(2, 1, 1, newLog.length).setValues([newLog]);
    }
  }
}

// --------------------- LogManager ------------------------ 