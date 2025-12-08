/**
 * PortfolioStats - Трекинг динамики портфеля
 * 
 * Отслеживает изменения портфеля инвестиций:
 * - Сохраняет историю общих показателей портфеля (1 раз в сутки после дневного сбора)
 * - Позволяет анализировать динамику портфеля за периоды
 */

// Заголовки колонок для таблицы истории
const PORTFOLIO_STATS_HEADERS = [
  'Дата/Время',
  'Вложения',
  'Текущая стоимость',
  'Прибыль/убыток',
  'Прибыль/убыток (%)',
  'Средняя прибыльность'
]

/**
 * Получает или создает лист PortfolioStats
 */
function getOrCreatePortfolioStatsSheet_() {
  return getOrCreateSheet_(SHEET_NAMES.PORTFOLIO_STATS)
}

/**
 * Рассчитывает текущие метрики портфеля (5 базовых метрик)
 * @returns {Object} {totalInvestment, totalCurrentValue, totalProfit, totalProfitPercent, avgProfitability}
 */
function portfolioStats_calculateMetrics_() {
  const investSheet = getInvestSheet_()
  
  if (!investSheet) {
    console.log('PortfolioStats: лист Invest не найден')
    return null
  }
  
  const lastRow = investSheet.getLastRow()
  if (lastRow <= 1) {
    return null // Нет данных
  }
  
  // Читаем данные из Invest batch-запросом
  const count = lastRow - 1
  const quantities = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.QUANTITY), count, 1).getValues()
  const totalInvestments = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.TOTAL_INVESTMENT), count, 1).getValues()
  const currentValues = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.CURRENT_VALUE_AFTER_FEE), count, 1).getValues()
  const profits = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.PROFIT), count, 1).getValues()
  const profitPercents = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.PROFIT_AFTER_FEE), count, 1).getValues()
  
  let totalInvestment = 0
  let totalCurrentValue = 0
  let totalProfit = 0
  let totalPositions = 0
  let sumProfitPercent = 0
  
  for (let i = 0; i < count; i++) {
    const quantity = Number(quantities[i][0]) || 0
    if (quantity <= 0) continue // Пропускаем позиции с нулевым количеством
    
    const investment = Number(totalInvestments[i][0]) || 0
    const currentValue = Number(currentValues[i][0]) || 0
    const profit = Number(profits[i][0]) || 0
    const profitPercent = Number(profitPercents[i][0]) || 0
    
    totalInvestment += investment
    totalCurrentValue += currentValue
    totalProfit += profit
    totalPositions++
    sumProfitPercent += profitPercent
  }
  
  // Рассчитываем общий процент прибыли
  const totalProfitPercent = totalInvestment > 0 
    ? ((totalCurrentValue / totalInvestment) - 1) 
    : 0
  
  // Средняя прибыльность позиций
  const avgProfitability = totalPositions > 0 
    ? sumProfitPercent / totalPositions 
    : 0
  
  return {
    totalInvestment,
    totalCurrentValue,
    totalProfit,
    totalProfitPercent,
    avgProfitability
  }
}

/**
 * Сохраняет текущие метрики портфеля в историю
 * Вставляет новую строку сверху (после заголовка)
 */
function portfolioStats_saveHistory_() {
  const sheet = getOrCreatePortfolioStatsSheet_()
  const metrics = portfolioStats_calculateMetrics_()
  
  if (!metrics) {
    console.log('PortfolioStats: нет данных для сохранения в историю')
    return
  }
  
  // Убеждаемся что заголовки есть
  if (sheet.getLastRow() < HEADER_ROW) {
    portfolioStats_formatTable()
  }
  
  // Вставляем новую строку сразу после заголовка (строка 2)
  const insertRow = HEADER_ROW + 1
  sheet.insertRowAfter(HEADER_ROW)
  
  // Сбрасываем форматирование заголовка для новой строки
  const newRowRange = sheet.getRange(insertRow, 1, 1, PORTFOLIO_STATS_HEADERS.length)
  newRowRange.setBackground(null) // Сбрасываем фон
  newRowRange.setFontWeight('normal') // Сбрасываем жирный шрифт
  
  const now = new Date()
  sheet.getRange(insertRow, 1, 1, PORTFOLIO_STATS_HEADERS.length).setValues([[
    now,
    metrics.totalInvestment,
    metrics.totalCurrentValue,
    metrics.totalProfit,
    metrics.totalProfitPercent,
    metrics.avgProfitability
  ]])
  
  // Форматирование даты/времени
  sheet.getRange(insertRow, 1).setNumberFormat('dd.MM.yyyy HH:mm')
  
  // Форматирование чисел
  sheet.getRange(insertRow, 2, 1, 2).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Вложения и стоимость
  sheet.getRange(insertRow, 4).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Прибыль/убыток
  sheet.getRange(insertRow, 5).setNumberFormat(NUMBER_FORMATS.PERCENT) // Прибыль/убыток %
  sheet.getRange(insertRow, 6).setNumberFormat(NUMBER_FORMATS.PERCENT) // Средняя прибыльность
  
  // Выравнивание
  sheet.getRange(insertRow, 1, 1, PORTFOLIO_STATS_HEADERS.length).setVerticalAlignment('middle')
  sheet.getRange(insertRow, 1, 1, PORTFOLIO_STATS_HEADERS.length).setHorizontalAlignment('center')
  
  // Цветовая индикация для прибыли/убытка
  if (metrics.totalProfit > 0) {
    sheet.getRange(insertRow, 4).setBackground(COLORS.PROFIT)
  } else if (metrics.totalProfit < 0) {
    sheet.getRange(insertRow, 4).setBackground(COLORS.LOSS)
  }
  
  console.log(`PortfolioStats: метрики сохранены в историю`)
}

/**
 * Форматирует лист PortfolioStats
 * Создает заголовки и применяет стандартное форматирование
 */
function portfolioStats_formatTable() {
  const sheet = getOrCreatePortfolioStatsSheet_()
  
  // Устанавливаем заголовки
  sheet.getRange(HEADER_ROW, 1, 1, PORTFOLIO_STATS_HEADERS.length).setValues([PORTFOLIO_STATS_HEADERS])
  
  // Форматирование заголовков
  formatHeaderRange_(sheet.getRange(HEADER_ROW, 1, 1, PORTFOLIO_STATS_HEADERS.length))
  
  // Высота заголовка
  sheet.setRowHeight(HEADER_ROW, HEADER_HEIGHT)
  
  // Замораживаем заголовок
  sheet.setFrozenRows(HEADER_ROW)
  
  // Ширины колонок
  sheet.setColumnWidth(1, 150) // Дата/Время
  sheet.setColumnWidth(2, 120) // Вложения
  sheet.setColumnWidth(3, 140) // Текущая стоимость
  sheet.setColumnWidth(4, 140) // Прибыль/убыток
  sheet.setColumnWidth(5, 130) // Прибыль/убыток %
  sheet.setColumnWidth(6, 150) // Средняя прибыльность
  
  // Форматируем существующие данные (если есть)
  const lastRow = sheet.getLastRow()
  if (lastRow > HEADER_ROW) {
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, PORTFOLIO_STATS_HEADERS.length)
    
    // Форматирование даты/времени
    sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, 1).setNumberFormat('dd.MM.yyyy HH:mm')
    
    // Форматирование чисел
    sheet.getRange(DATA_START_ROW, 2, lastRow - HEADER_ROW, 2).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Вложения и стоимость
    sheet.getRange(DATA_START_ROW, 4, lastRow - HEADER_ROW, 1).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Прибыль/убыток
    sheet.getRange(DATA_START_ROW, 5, lastRow - HEADER_ROW, 1).setNumberFormat(NUMBER_FORMATS.PERCENT) // Прибыль/убыток %
    sheet.getRange(DATA_START_ROW, 6, lastRow - HEADER_ROW, 1).setNumberFormat(NUMBER_FORMATS.PERCENT) // Средняя прибыльность
    
    // Выравнивание
    dataRange.setVerticalAlignment('middle')
    dataRange.setHorizontalAlignment('center')
    
    // Высота строк данных
    sheet.setRowHeights(DATA_START_ROW, lastRow - HEADER_ROW, ROW_HEIGHT)
    
    // Цветовая индикация для прибыли/убытка
    const profitRange = sheet.getRange(DATA_START_ROW, 4, lastRow - HEADER_ROW, 1)
    const profitValues = profitRange.getValues()
    for (let i = 0; i < profitValues.length; i++) {
      const profit = Number(profitValues[i][0]) || 0
      if (profit > 0) {
        sheet.getRange(DATA_START_ROW + i, 4).setBackground(COLORS.PROFIT)
      } else if (profit < 0) {
        sheet.getRange(DATA_START_ROW + i, 4).setBackground(COLORS.LOSS)
      }
    }
  }
  
  console.log('PortfolioStats: лист отформатирован')
}

/**
 * Основная функция обновления аналитики портфеля
 * Теперь используется только для форматирования листа (если вызывается вручную)
 * Автоматическое сохранение истории происходит через portfolioStats_saveHistory_()
 */
function portfolioStats_update() {
  portfolioStats_formatTable()
}

/**
 * Ручное сохранение текущих метрик портфеля в историю (для тестирования)
 * Можно вызывать из меню для проверки работы сохранения истории
 */
function portfolioStats_saveHistoryManual() {
  try {
    portfolioStats_saveHistory_()
    SpreadsheetApp.getUi().alert('✅ Метрики портфеля сохранены в историю')
  } catch (e) {
    console.error('PortfolioStats: ошибка при ручном сохранении истории:', e)
    SpreadsheetApp.getUi().alert('❌ Ошибка при сохранении истории: ' + e.message)
  }
}
