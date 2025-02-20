// --------------------- 常量 --------------------------------

// 表格相关常量
const SHEET_CONSTANTS = {
  COLORS: {
    MODIFIED: "#b3e5fc",  // 修改 - 浅蓝色
    ADDED: "#dcedc8",    // 新增 - 淡绿色
    REMOVED: "#ffdce0",  // 删除 - Git风格浅红色
    HEADER_MODIFIED: "#fff9c4",  // 表头修改 - 浅黄色
    CONFLICT: "#f8bbd0"  // 冲突 - 粉色
  },
  CACHE_DURATION: 600
};

// 合并相关常量
const MERGE_CONSTANTS = {
  ID_SUFFIX: '_INT_id',
  CONFLICT_PREFIX: '冲突: ',
  PREVIEW_SUFFIX: '_预览',
  COLORS: {
    NEW: '#b7e1cd',      // 浅绿色 - 新行
    CONFLICT: '#ffdce0', // 浅红色 - 冲突
    UPDATED: '#c9daf8',  // 浅蓝色 - 已更新
    RESOLVED: "#e8f5e9"  // 已解决 - 更浅的绿色
  },
  CACHE_KEYS: {
    MERGE_STATE: "merge_state",
    CURRENT_CONFLICTS: "current_conflicts"
  },
  RESOLUTION_TYPES: {
    SOURCE: "source",
    TARGET: "target",
    CUSTOM: "custom",
    AUTO: "auto"
  }
};

// 比较相关常量
const COMPARE_CONSTANTS = {
  COLORS: {
    MODIFIED: "#ffcdd2",  // 修改 - 浅红色
    ADDED: "#c8e6c9",    // 新增 - 浅绿色
    REMOVED: "#ffdce0",  // 删除 - Git风格浅红色
    HEADER_MODIFIED: "#fff9c4"  // 表头修改 - 浅黄色
  },
  CACHE_DURATION: 600
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

// --------------------- 常量 --------------------------------

// --------------------- NoteManager ------------------------ 

/**
 * 注释管理工具类
 */
class NoteManager {
  /**
   * 添加系统注释
   * @param {string} originalNote 原始注释
   * @param {string} type 注释类型
   * @param {string} content 注释内容
   * @returns {string} 新的注释内容
   */
  static addSystemNote(originalNote, type, content) {
    const systemNotes = this.extractSystemNotes(originalNote);
    systemNotes[type] = content;
    
    const systemPart = this.formatSystemNotes(systemNotes);
    return this.appendSystemNote(originalNote, systemPart);
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
   * 提取系统注释
   * @private
   * @param {string} note 完整注释
   * @returns {Object} 系统注释对象
   */
  static extractSystemNotes(note) {
    if (!note) return {};
    
    const start = note.indexOf(NOTE_CONSTANTS.SYSTEM_NOTE_START);
    const end = note.indexOf(NOTE_CONSTANTS.SYSTEM_NOTE_END);
    
    if (start === -1 || end === -1) return {};
    
    const notesSection = note.substring(
      start + NOTE_CONSTANTS.SYSTEM_NOTE_START.length,
      end
    );

    const systemNotes = {};
    notesSection.split(NOTE_CONSTANTS.LINE_SEPARATOR).forEach(line => {
      const [key, ...valueParts] = line.split(NOTE_CONSTANTS.KEY_VALUE_SEPARATOR);
      if (key && valueParts.length > 0) {
        systemNotes[key.trim()] = valueParts.join(NOTE_CONSTANTS.KEY_VALUE_SEPARATOR).trim();
      }
    });
    
    return systemNotes;
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