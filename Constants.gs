// 表格相关常量
const SHEET_CONSTANTS = {
  COLORS: {
    MODIFIED: "#ffcdd2",  // 修改 - 浅红色
    ADDED: "#c8e6c9",    // 新增 - 浅绿色
    REMOVED: "#ffdce0",  // 删除 - Git风格浅红色
    HEADER_MODIFIED: "#fff9c4",  // 表头修改 - 浅黄色
    CONFLICT: "#f8bbd0"  // 冲突 - 粉色
  },
  CACHE_DURATION: 600
};

// 版本控制相关常量
const VERSION_CONSTANTS = {
  NOTES: {
    BASE_INFO_PREFIX: "BASE_INFO:",
    ORIGINAL_VALUE_PREFIX: "ORIGINAL_VALUE:",
    CHANGE_HISTORY_PREFIX: "HISTORY:"
  },
  MAX_HISTORY_ENTRIES: 10
};

// 合并相关常量
const MERGE_CONSTANTS = {
  ID_SUFFIX: '_INT_id',
  CONFLICT_PREFIX: '冲突: ',
  PREVIEW_SUFFIX: '_预览',
  COLORS: {
    NEW: '#b7e1cd',      // 浅绿色 - 新行
    CONFLICT: '#fce8b2', // 浅黄色 - 冲突
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
  },
  NOTES: {
    ORIGINAL_VALUE_PREFIX: "原值:",
    BASE_INFO_KEY: "baseInfo"
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