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

