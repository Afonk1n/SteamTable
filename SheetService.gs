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
 * Получает лист PortfolioStats
 * @returns {Sheet|null} Лист PortfolioStats или null
 */
function getPortfolioStatsSheet_() {
  return getSheet_(SHEET_NAMES.PORTFOLIO_STATS)
}

/**
 * Получает или создает лист PortfolioStats
 * @returns {Sheet} Лист PortfolioStats
 */
function getOrCreatePortfolioStatsSheet_() {
  return getOrCreateSheet_(SHEET_NAMES.PORTFOLIO_STATS)
}

/**
 * Получает лист AutoLog
 * @returns {Sheet|null} Лист AutoLog или null
 */
function getAutoLogSheet_() {
  return getSheet_(SHEET_NAMES.AUTO_LOG)
}

/**
 * Получает или создает лист AutoLog
 * @returns {Sheet} Лист AutoLog
 */
function getOrCreateAutoLogSheet_() {
  const headers = ['Дата/Время', 'Лист', 'Действие', 'Результат']
  const columnWidths = [150, 100, 200, 100]
  return createLogSheet_(SHEET_NAMES.AUTO_LOG, headers, columnWidths)
}

/**
 * Получает лист Log
 * @returns {Sheet|null} Лист Log или null
 */
function getLogSheet_() {
  return getSheet_(SHEET_NAMES.LOG)
}

/**
 * Получает или создает лист Log
 * @returns {Sheet} Лист Log
 */
function getOrCreateLogSheet_() {
  const headers = ['Дата/Время', 'Операция', 'Предмет', 'Количество', 'Цена', 'Сумма', 'Источник']
  const columnWidths = [150, 100, 200, 100, 100, 100, 100]
  return createLogSheet_(SHEET_NAMES.LOG, headers, columnWidths)
}

/**
 * Получает лист HeroStats
 * @returns {Sheet|null} Лист HeroStats или null
 */
function getHeroStatsSheet_() {
  return getSheet_(SHEET_NAMES.HERO_STATS)
}

/**
 * Получает или создает лист HeroStats
 * @returns {Sheet} Лист HeroStats
 */
function getOrCreateHeroStatsSheet_() {
  return getOrCreateSheet_(SHEET_NAMES.HERO_STATS)
}

/**
 * Получает лист HeroMapping
 * @returns {Sheet|null} Лист HeroMapping или null
 */
function getHeroMappingSheet_() {
  return getSheet_(SHEET_NAMES.HERO_MAPPING)
}

/**
 * Получает или создает лист HeroMapping
 * @returns {Sheet} Лист HeroMapping
 */
function getOrCreateHeroMappingSheet_() {
  return getOrCreateSheet_(SHEET_NAMES.HERO_MAPPING)
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

