/**
 * SheetService - Унифицированный сервис для работы с листами
 * 
 * Централизует функции получения и создания листов для устранения дублирования
 */

/**
 * Получает лист по имени
 * @param {string} sheetName - Имя листа
 * @returns {Sheet|null} Лист или null если не найден
 */
function getSheet_(sheetName) {
  const ss = SpreadsheetApp.getActive()
  return ss.getSheetByName(sheetName)
}

/**
 * Получает или создает лист по имени
 * @param {string} sheetName - Имя листа
 * @returns {Sheet} Лист (созданный или существующий)
 */
function getOrCreateSheet_(sheetName) {
  const ss = SpreadsheetApp.getActive()
  let sheet = ss.getSheetByName(sheetName)
  if (!sheet) {
    sheet = ss.insertSheet(sheetName)
  }
  return sheet
}

/**
 * Получает лист Invest
 * @returns {Sheet|null} Лист Invest или null
 */
function getInvestSheet_() {
  return getSheet_(SHEET_NAMES.INVEST)
}

/**
 * Получает или создает лист Invest
 * @returns {Sheet} Лист Invest
 */
function getOrCreateInvestSheet_() {
  return getOrCreateSheet_(SHEET_NAMES.INVEST)
}

/**
 * Получает лист Sales
 * @returns {Sheet|null} Лист Sales или null
 */
function getSalesSheet_() {
  return getSheet_(SHEET_NAMES.SALES)
}

/**
 * Получает или создает лист Sales
 * @returns {Sheet} Лист Sales
 */
function getOrCreateSalesSheet_() {
  return getOrCreateSheet_(SHEET_NAMES.SALES)
}

/**
 * Получает лист History
 * @returns {Sheet|null} Лист History или null
 */
function getHistorySheet_() {
  return getSheet_(SHEET_NAMES.HISTORY)
}

/**
 * Получает или создает лист History
 * @returns {Sheet} Лист History
 */
function getOrCreateHistorySheet_() {
  return getOrCreateSheet_(SHEET_NAMES.HISTORY)
}

/**
 * Универсальная функция для создания и настройки листа лога
 * @param {string} sheetName - Имя листа
 * @param {Array<string>} headers - Массив заголовков колонок
 * @param {Array<number>} columnWidths - Массив ширин колонок (соответствует порядку заголовков)
 * @returns {Sheet} Созданный или существующий лист
 */
function createLogSheet_(sheetName, headers, columnWidths) {
  const ss = SpreadsheetApp.getActive()
  let sheet = ss.getSheetByName(sheetName)
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName)
    const numCols = headers.length
    
    // Устанавливаем заголовки
    sheet.getRange(HEADER_ROW, 1, 1, numCols).setValues([headers])
    
    // Форматируем заголовок
    formatHeaderRange_(sheet.getRange(HEADER_ROW, 1, 1, numCols))
    
    // Замораживаем строку заголовка
    sheet.setFrozenRows(HEADER_ROW)
    
    // Устанавливаем ширины колонок
    columnWidths.forEach((width, index) => {
      sheet.setColumnWidth(index + 1, width)
    })
  }
  
  return sheet
}

