/**
 * FormattingService - Унифицированный сервис для условного форматирования
 * 
 * Централизует все правила условного форматирования для устранения дублирования
 * и обеспечения консистентности между листами Invest, Sales и History.
 */

/**
 * Создает правила форматирования для трендов
 * @param {Range} trendRange - Диапазон колонки трендов
 * @returns {Array<ConditionalFormatRule>} Массив правил форматирования
 */
function createTrendFormattingRules_(trendRange) {
  return [
    // Форматирование для объединенного формата "🟩 Растет 12 дн." - проверяем начало строки
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith('🟩')
      .setBackground(COLORS.TREND_UP)
      .setRanges([trendRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith('🟥')
      .setBackground(COLORS.TREND_DOWN)
      .setRanges([trendRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith('🟨')
      .setBackground(COLORS.TREND_SIDEWAYS)
      .setRanges([trendRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith('🟪')
      .setBackground(COLORS.TREND_UNKNOWN)
      .setRanges([trendRange])
      .build(),
    // Обратная совместимость: если формат старый (только эмодзи)
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('🟩')
      .setBackground(COLORS.TREND_UP)
      .setRanges([trendRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('🟥')
      .setBackground(COLORS.TREND_DOWN)
      .setRanges([trendRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('🟨')
      .setBackground(COLORS.TREND_SIDEWAYS)
      .setRanges([trendRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('🟪')
      .setBackground(COLORS.TREND_UNKNOWN)
      .setRanges([trendRange])
      .build()
  ]
}

/**
 * Создает правила форматирования для фаз цикла
 * @param {Range} phaseRange - Диапазон колонки фаз
 * @returns {Array<ConditionalFormatRule>} Массив правил форматирования
 */
function createPhaseFormattingRules_(phaseRange) {
  return [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('ДНО')
      .setBackground('#c8e6c9')
      .setRanges([phaseRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('РОСТ')
      .setBackground('#dcedc8')
      .setRanges([phaseRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('ПИК')
      .setBackground('#ffcdd2')
      .setRanges([phaseRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('КОРРЕКЦИЯ')
      .setBackground('#fff9c4')
      .setRanges([phaseRange])
      .build()
  ]
}

/**
 * Создает правила форматирования для рекомендаций
 * @param {Range} recommendationRange - Диапазон колонки рекомендаций
 * @returns {Array<ConditionalFormatRule>} Массив правил форматирования
 */
function createRecommendationFormattingRules_(recommendationRange) {
  return [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('КУПИТЬ')
      .setBackground('#c8e6c9')
      .setFontColor('#1b5e20')
      .setBold(true)
      .setRanges([recommendationRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('ПРОДАТЬ')
      .setBackground('#ffcdd2')
      .setFontColor('#b71c1c')
      .setBold(true)
      .setRanges([recommendationRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('ДЕРЖАТЬ')
      .setBackground('#bbdefb')
      .setFontColor('#0d47a1')
      .setRanges([recommendationRange])
      .build()
  ]
}

/**
 * Создает правила форматирования для прибыли/убытка
 * @param {Range} profitRange - Диапазон колонок прибыли
 * @returns {Array<ConditionalFormatRule>} Массив правил форматирования
 */
function createProfitFormattingRules_(profitRange) {
  return [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(COLORS.LOSS)
      .setRanges([profitRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(COLORS.PROFIT)
      .setRanges([profitRange])
      .build()
  ]
}

/**
 * Создает правила форматирования для падения цены (Sales)
 * @param {Range} dropRange - Диапазон колонки падения цены
 * @returns {Array<ConditionalFormatRule>} Массив правил форматирования
 */
function createPriceDropFormattingRules_(dropRange) {
  return [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(COLORS.PROFIT)
      .setRanges([dropRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(COLORS.LOSS)
      .setRanges([dropRange])
      .build()
  ]
}

/**
 * Применяет все правила условного форматирования для аналитики
 * @param {Sheet} sheet - Лист для форматирования
 * @param {Object} config - Конфигурация колонок (trendCol, phaseCol, recommendationCol)
 * @param {number} lastRow - Последняя строка с данными
 */
function applyAnalyticsFormatting_(sheet, config, lastRow) {
  if (lastRow <= 1) {
    sheet.setConditionalFormatRules([])
    return
  }

  const rules = []
  const dataRowCount = lastRow - 1

  if (config.trendCol) {
    const trendRange = sheet.getRange(DATA_START_ROW, config.trendCol, dataRowCount, 1)
    rules.push(...createTrendFormattingRules_(trendRange))
  }

  if (config.phaseCol) {
    const phaseRange = sheet.getRange(DATA_START_ROW, config.phaseCol, dataRowCount, 1)
    rules.push(...createPhaseFormattingRules_(phaseRange))
  }

  if (config.recommendationCol) {
    const recommendationRange = sheet.getRange(DATA_START_ROW, config.recommendationCol, dataRowCount, 1)
    rules.push(...createRecommendationFormattingRules_(recommendationRange))
  }

  if (config.profitRange) {
    rules.push(...createProfitFormattingRules_(config.profitRange))
  }

  if (config.dropRange) {
    rules.push(...createPriceDropFormattingRules_(config.dropRange))
  }

  sheet.setConditionalFormatRules(rules)
}

/**
 * Универсальная функция для базового форматирования таблицы Invest/Sales
 * @param {Sheet} sheet - Лист для форматирования
 * @param {Array<string>} headers - Массив заголовков
 * @param {Object} columns - Объект с колонками (INVEST_COLUMNS или SALES_COLUMNS)
 * @param {Function} getSheetFn - Функция для получения листа (для валидации)
 * @param {string} sheetName - Имя листа для ошибок
 * @returns {number} lastRow - Последняя строка или 0 если ошибка
 */
function formatTableBase_(sheet, headers, columns, getSheetFn, sheetName) {
  if (!sheet || !headers || !Array.isArray(headers) || headers.length === 0) {
    console.error(`${sheetName}: некорректные параметры для форматирования`)
    if (headers && !Array.isArray(headers)) {
      SpreadsheetApp.getUi().alert(`Ошибка: заголовки ${sheetName} не определены`)
    }
    return 0
  }
  
  const lastRow = sheet.getLastRow()
  
  // Устанавливаем заголовки
  sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers])
  
  // Проверяем и исправляем заголовок "Потенциал" на "Потенциал (P85)" если нужно
  const potentialColIndex = getColumnIndex(columns.POTENTIAL)
  if (potentialColIndex > 0) {
    const currentPotentialHeader = sheet.getRange(HEADER_ROW, potentialColIndex).getValue()
    if (currentPotentialHeader && currentPotentialHeader !== 'Потенциал (P85)') {
      sheet.getRange(HEADER_ROW, potentialColIndex).setValue('Потенциал (P85)')
    }
  }
  
  // Форматируем заголовок
  formatHeaderRange_(sheet.getRange(HEADER_ROW, 1, 1, headers.length))
  
  // Устанавливаем высоты строк
  sheet.setRowHeight(HEADER_ROW, HEADER_HEIGHT)
  if (lastRow > 1) {
    sheet.setRowHeights(DATA_START_ROW, lastRow - 1, ROW_HEIGHT)
  }
  
  // Замораживаем строку заголовка
  sheet.setFrozenRows(HEADER_ROW)
  
  return lastRow
}

/**
 * Универсальная функция для форматирования новой строки в Invest/Sales
 * @param {Sheet} sheet - Лист для форматирования
 * @param {number} row - Номер строки
 * @param {Object} config - Конфигурация (COLUMNS, STEAM_APPID для Sales)
 * @param {Object} numberFormatConfig - Конфигурация форматов чисел {columnKey: format}
 * @param {boolean} addImageAndLink - Добавлять ли изображение и ссылку (для Sales)
 */
function formatNewRowUniversal_(sheet, row, config, numberFormatConfig, addImageAndLink = false) {
  if (row <= HEADER_ROW) return
  
  const name = sheet.getRange(`B${row}`).getValue()
  if (!name) return
  
  // Базовое форматирование строки
  const numCols = getColumnIndex(config.COLUMNS.RECOMMENDATION)
  sheet.getRange(row, 1, 1, numCols).setVerticalAlignment('middle').setHorizontalAlignment('center')
  sheet.getRange(`B${row}`).setHorizontalAlignment('left')
  
  // Форматы чисел
  for (const [columnKey, format] of Object.entries(numberFormatConfig)) {
    const colIndex = getColumnIndex(config.COLUMNS[columnKey])
    if (colIndex > 0) {
      sheet.getRange(row, colIndex).setNumberFormat(format)
    }
  }
  
  // Добавляем изображение и ссылку (только для Sales)
  if (addImageAndLink && config.STEAM_APPID) {
    setImageAndLink_(sheet, row, config.STEAM_APPID, name, config.COLUMNS)
  }
  
  // Устанавливаем высоту строки
  sheet.setRowHeight(row, ROW_HEIGHT)
}

