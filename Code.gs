/**
 * Base64 Encode Input
 * @param {any | Array<any[]>} input - Input cell, or range of cells
 * @param {boolean} [OPT_webSafe=true] - If should use websafe variant of base64
 * @param {boolean} [OPT_plainText=false] - If should treat input as plaintext instead of UTF-8
 */
function base64Encode(input, OPT_webSafe, OPT_plainText) {
  if (!input) return input;
  const charSet = OPT_plainText ? Utilities.Charset.US_ASCII : Utilities.Charset.UTF_8;
  const useWebSafe = OPT_webSafe !== false;
  const encoder = useWebSafe ? Utilities.base64EncodeWebSafe : Utilities.base64Encode;
  
  if (Array.isArray(input)) {
    return input.map(t => base64Encode(t, OPT_webSafe, OPT_plainText));
  }
  
  return encoder(input, charSet);
}

/**
 * Base64 Decode Input
 * @param {any | Array<any[]>} input - Input cell, or range of cells
 * @param {boolean} [OPT_webSafe=true] - If should use websafe variant of base64
 * @param {boolean} [OPT_plainText=false] - If should treat input as plaintext instead of UTF-8
 */
function base64Decode(input, OPT_webSafe, OPT_plainText) {
  if (!input) return input;
  const charSet = OPT_plainText ? Utilities.Charset.US_ASCII : Utilities.Charset.UTF_8;
  const useWebSafe = OPT_webSafe !== false;
  const decoder = useWebSafe ? Utilities.base64DecodeWebSafe : Utilities.base64Decode;
  
  if (Array.isArray(input)) {
    return input.map(t => base64Decode(t, OPT_webSafe, OPT_plainText));
  }
  
  return decoder(input, charSet);
}

/**
 * 获取当前文档中所有表格的信息（带缓存）
 * @returns {Object} 包含所有表格名称和当前表格的对象
 */
function getSheetInfo() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'sheet_info';
    const cached = cache.get(cacheKey);
    
    if (cached != null) {
      return JSON.parse(cached);
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    const result = {
      sheets: ss.getSheets().map(sheet => sheet.getName()),
      activeSheet: activeSheet.getName()
    };
    
    cache.put(cacheKey, JSON.stringify(result), 600);
    return result;
    
  } catch (error) {
    throw new Error("获取表格信息失败: " + error.toString());
  }
}

/**
 * 获取当前页签名称
 * @returns {string} 当前页签名称
 */
function getCurrentSheetName() {
  return SpreadsheetApp.getActiveSheet().getName();
}

/**
 * 获取当前页签A1单元格的注释
 * @returns {string} A1单元格的注释内容
 */
function getCurrentSheetA1Note() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const a1Note = sheet.getRange("A1").getNote();
    return a1Note || '';
  } catch (error) {
    console.error('获取A1单元格注释失败:', error);
    return '';
  }
}