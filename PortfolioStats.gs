/**
 * PortfolioStats - Аналитика портфеля инвестиций
 * 
 * Создает и обновляет лист со статистикой по портфелю:
 * - Общие показатели (вложения, стоимость, прибыль)
 * - Топ-5 прибыльных/убыточных позиций
 * - Распределение по фазам и трендам
 * - Диаграммы визуализации
 */

const PORTFOLIO_STATS_SHEET_NAME = SHEET_NAMES.PORTFOLIO_STATS

/**
 * Получает или создает лист PortfolioStats
 */
function getOrCreatePortfolioStatsSheet_() {
  return getOrCreateSheet_(PORTFOLIO_STATS_SHEET_NAME)
}

/**
 * Основная функция обновления аналитики портфеля
 */
function portfolioStats_update() {
  const sheet = getOrCreatePortfolioStatsSheet_()
  const investSheet = getInvestSheet_()
  
  if (!investSheet) {
    console.log('PortfolioStats: лист Invest не найден')
    return
  }
  
  const lastRow = investSheet.getLastRow()
  if (lastRow <= 1) {
    // Нет данных в Invest
    portfolioStats_clearSheet_(sheet)
    return
  }
  
  // Читаем все данные из Invest одним batch-запросом
  const count = lastRow - 1
  const names = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.NAME), count, 1).getValues()
  const quantities = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.QUANTITY), count, 1).getValues()
  const totalInvestments = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.TOTAL_INVESTMENT), count, 1).getValues()
  const currentValues = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.CURRENT_VALUE_AFTER_FEE), count, 1).getValues()
  const profits = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.PROFIT), count, 1).getValues()
  const profitPercents = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.PROFIT_AFTER_FEE), count, 1).getValues()
  const phases = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.PHASE), count, 1).getValues()
  const trends = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.TREND), count, 1).getValues()
  
  // Рассчитываем метрики
  const stats = portfolioStats_calculateMetrics_(
    names, quantities, totalInvestments, currentValues, profits, profitPercents, phases, trends
  )
  
  // Форматируем и заполняем лист
  portfolioStats_formatAndFill_(sheet, stats)
  
  console.log(`PortfolioStats: аналитика обновлена. Позиций: ${stats.totalPositions}`)
}

/**
 * Рассчитывает все метрики портфеля
 */
function portfolioStats_calculateMetrics_(names, quantities, totalInvestments, currentValues, profits, profitPercents, phases, trends) {
  let totalInvestment = 0
  let totalCurrentValue = 0
  let totalProfit = 0
  let totalPositions = 0
  let profitableCount = 0
  let unprofitableCount = 0
  let neutralCount = 0
  
  const positions = []
  const phaseDistribution = new Map()
  const trendDistribution = new Map()
  
  for (let i = 0; i < names.length; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue
    
    const qty = Number(quantities[i][0]) || 0
    if (qty <= 0) continue // Пропускаем позиции с нулевым количеством
    
    const inv = Number(totalInvestments[i][0]) || 0
    const currVal = Number(currentValues[i][0]) || 0
    const profit = Number(profits[i][0]) || 0
    const profitPercent = Number(profitPercents[i][0]) || 0
    const phase = String(phases[i][0] || '').trim()
    const trend = String(trends[i][0] || '').trim()
    
    totalInvestment += inv
    totalCurrentValue += currVal
    totalProfit += profit
    totalPositions++
    
    // Классификация по прибыльности
    if (profitPercent > 0.01) {
      profitableCount++
    } else if (profitPercent < -0.01) {
      unprofitableCount++
    } else {
      neutralCount++
    }
    
    // Сохраняем позицию для топ-5
    positions.push({
      name,
      profit,
      profitPercent,
      investment: inv,
      currentValue: currVal
    })
    
    // Распределение по фазам
    if (phase) {
      phaseDistribution.set(phase, (phaseDistribution.get(phase) || 0) + 1)
    }
    
    // Распределение по трендам (извлекаем эмодзи из объединенного формата)
    if (trend) {
      const trendMatch = trend.match(/^([🟥🟩🟨🟪])/)
      const trendEmoji = trendMatch ? trendMatch[1] : '🟪'
      trendDistribution.set(trendEmoji, (trendDistribution.get(trendEmoji) || 0) + 1)
    }
  }
  
  // Сортируем позиции по прибыльности
  positions.sort((a, b) => b.profit - a.profit)
  const top5Profitable = positions.slice(0, 5)
  const top5Unprofitable = positions.slice(-5).reverse()
  
  // Средняя прибыльность
  const avgProfitability = totalPositions > 0 
    ? positions.reduce((sum, p) => sum + p.profitPercent, 0) / totalPositions 
    : 0
  
  // Общая прибыль в процентах
  const totalProfitPercent = totalInvestment > 0 
    ? ((totalCurrentValue - totalInvestment) / totalInvestment) 
    : 0
  
  return {
    totalInvestment,
    totalCurrentValue,
    totalProfit,
    totalProfitPercent,
    avgProfitability,
    totalPositions,
    profitableCount,
    unprofitableCount,
    neutralCount,
    top5Profitable,
    top5Unprofitable,
    phaseDistribution,
    trendDistribution
  }
}

/**
 * Форматирует и заполняет лист данными
 */
function portfolioStats_formatAndFill_(sheet, stats) {
  // Очищаем лист
  sheet.clear()
  
  // Заголовок
  sheet.getRange(1, 1).setValue('СТАТИСТИКА ПОРТФЕЛЯ')
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold')
  sheet.setRowHeight(1, 40)
  
  // Общие показатели
  let row = 3
  sheet.getRange(row, 1).setValue('Общие показатели').setFontWeight('bold').setFontSize(12)
  row++
  
  sheet.getRange(row, 1).setValue('Общая сумма вложений:')
  sheet.getRange(row, 2).setValue(stats.totalInvestment).setNumberFormat(NUMBER_FORMATS.CURRENCY)
  row++
  
  sheet.getRange(row, 1).setValue('Текущая стоимость:')
  sheet.getRange(row, 2).setValue(stats.totalCurrentValue).setNumberFormat(NUMBER_FORMATS.CURRENCY)
  row++
  
  sheet.getRange(row, 1).setValue('Прибыль/убыток:')
  sheet.getRange(row, 2).setValue(stats.totalProfit).setNumberFormat(NUMBER_FORMATS.CURRENCY)
  sheet.getRange(row, 3).setValue(stats.totalProfitPercent).setNumberFormat(NUMBER_FORMATS.PERCENT)
  // Цветовая индикация
  if (stats.totalProfit > 0) {
    sheet.getRange(row, 2, 1, 2).setBackground(COLORS.PROFIT)
  } else if (stats.totalProfit < 0) {
    sheet.getRange(row, 2, 1, 2).setBackground(COLORS.LOSS)
  }
  row++
  
  sheet.getRange(row, 1).setValue('Средняя прибыльность:')
  sheet.getRange(row, 2).setValue(stats.avgProfitability).setNumberFormat(NUMBER_FORMATS.PERCENT)
  row++
  
  sheet.getRange(row, 1).setValue('Активных позиций:')
  sheet.getRange(row, 2).setValue(stats.totalPositions).setNumberFormat(NUMBER_FORMATS.INTEGER)
  row++
  
  // Топ-5 прибыльных
  row += 2
  sheet.getRange(row, 1).setValue('Топ-5 прибыльных').setFontWeight('bold').setFontSize(12)
  row++
  
  // Заголовки таблицы
  sheet.getRange(row, 1, 1, 4).setValues([['Название', 'Прибыль', 'Прибыль %', 'Вложения']])
  formatHeaderRange_(sheet.getRange(row, 1, 1, 4))
  row++
  
  // Данные
  const profitableData = stats.top5Profitable.map(p => [
    p.name,
    p.profit,
    p.profitPercent,
    p.investment
  ])
  if (profitableData.length > 0) {
    sheet.getRange(row, 1, profitableData.length, 4).setValues(profitableData)
    sheet.getRange(row, 2, profitableData.length, 1).setNumberFormat(NUMBER_FORMATS.CURRENCY)
    sheet.getRange(row, 3, profitableData.length, 1).setNumberFormat(NUMBER_FORMATS.PERCENT)
    sheet.getRange(row, 4, profitableData.length, 1).setNumberFormat(NUMBER_FORMATS.CURRENCY)
    sheet.getRange(row, 1, profitableData.length, 4).setBackground(COLORS.PROFIT)
    row += profitableData.length
  } else {
    sheet.getRange(row, 1).setValue('Нет данных')
    row++
  }
  
  // Топ-5 убыточных
  row += 2
  sheet.getRange(row, 1).setValue('Топ-5 убыточных').setFontWeight('bold').setFontSize(12)
  row++
  
  // Заголовки таблицы
  sheet.getRange(row, 1, 1, 4).setValues([['Название', 'Убыток', 'Убыток %', 'Вложения']])
  formatHeaderRange_(sheet.getRange(row, 1, 1, 4))
  row++
  
  // Данные
  const unprofitableData = stats.top5Unprofitable.map(p => [
    p.name,
    p.profit,
    p.profitPercent,
    p.investment
  ])
  if (unprofitableData.length > 0) {
    sheet.getRange(row, 1, unprofitableData.length, 4).setValues(unprofitableData)
    sheet.getRange(row, 2, unprofitableData.length, 1).setNumberFormat(NUMBER_FORMATS.CURRENCY)
    sheet.getRange(row, 3, unprofitableData.length, 1).setNumberFormat(NUMBER_FORMATS.PERCENT)
    sheet.getRange(row, 4, unprofitableData.length, 1).setNumberFormat(NUMBER_FORMATS.CURRENCY)
    sheet.getRange(row, 1, unprofitableData.length, 4).setBackground(COLORS.LOSS)
    row += unprofitableData.length
  } else {
    sheet.getRange(row, 1).setValue('Нет данных')
    row++
  }
  
  // Распределение по фазам
  row += 2
  sheet.getRange(row, 1).setValue('Распределение по фазам').setFontWeight('bold').setFontSize(12)
  row++
  
  sheet.getRange(row, 1, 1, 2).setValues([['Фаза', 'Количество']])
  formatHeaderRange_(sheet.getRange(row, 1, 1, 2))
  row++
  
  const phaseData = Array.from(stats.phaseDistribution.entries())
    .sort((a, b) => b[1] - a[1])
    .map(([phase, count]) => [phase, count])
  
  if (phaseData.length > 0) {
    sheet.getRange(row, 1, phaseData.length, 2).setValues(phaseData)
    sheet.getRange(row, 2, phaseData.length, 1).setNumberFormat(NUMBER_FORMATS.INTEGER)
    row += phaseData.length
  } else {
    sheet.getRange(row, 1).setValue('Нет данных')
    row++
  }
  
  // Распределение по трендам
  row += 2
  sheet.getRange(row, 1).setValue('Распределение по трендам').setFontWeight('bold').setFontSize(12)
  row++
  
  sheet.getRange(row, 1, 1, 2).setValues([['Тренд', 'Количество']])
  formatHeaderRange_(sheet.getRange(row, 1, 1, 2))
  row++
  
  const trendLabels = {
    '🟥': 'Падает',
    '🟩': 'Растет',
    '🟨': 'Боковик',
    '🟪': 'Нет данных'
  }
  
  const trendData = Array.from(stats.trendDistribution.entries())
    .sort((a, b) => b[1] - a[1])
    .map(([trend, count]) => [`${trend} ${trendLabels[trend] || 'Неизвестно'}`, count])
  
  if (trendData.length > 0) {
    sheet.getRange(row, 1, trendData.length, 2).setValues(trendData)
    sheet.getRange(row, 2, trendData.length, 1).setNumberFormat(NUMBER_FORMATS.INTEGER)
  } else {
    sheet.getRange(row, 1).setValue('Нет данных')
  }
  
  // Форматирование колонок
  sheet.setColumnWidth(1, 200)
  sheet.setColumnWidth(2, 150)
  sheet.setColumnWidth(3, 120)
  sheet.setColumnWidth(4, 150)
  
  // Выравнивание
  sheet.getRange(1, 1, row + 10, 4).setVerticalAlignment('middle')
  sheet.getRange(4, 1, row, 1).setHorizontalAlignment('left') // Названия метрик слева
  sheet.getRange(4, 2, row, 3).setHorizontalAlignment('center') // Значения по центру
  
  // Создаем диаграммы
  portfolioStats_createCharts_(sheet, stats)
}

/**
 * Создает диаграммы для визуализации
 */
function portfolioStats_createCharts_(sheet, stats) {
  // Удаляем старые диаграммы
  const charts = sheet.getCharts()
  charts.forEach(chart => sheet.removeChart(chart))
  
  // 1. Круговая диаграмма: распределение прибыльности
  const profitabilityDataRange = sheet.getRange(4, 1, 3, 2) // Прибыльные/убыточные/нейтральные
  // Создаем данные для диаграммы
  const profitabilityRow = 20
  sheet.getRange(profitabilityRow, 6).setValue('Прибыльные')
  sheet.getRange(profitabilityRow, 7).setValue(stats.profitableCount)
  sheet.getRange(profitabilityRow + 1, 6).setValue('Убыточные')
  sheet.getRange(profitabilityRow + 1, 7).setValue(stats.unprofitableCount)
  sheet.getRange(profitabilityRow + 2, 6).setValue('Нейтральные')
  sheet.getRange(profitabilityRow + 2, 7).setValue(stats.neutralCount)
  
  // Создание диаграмм через EmbeddedChartBuilder
  // Примечание: Диаграммы создаются вручную пользователем или через встроенные функции Google Sheets
  // Здесь мы только подготавливаем данные для диаграмм
  
  // Для автоматического создания диаграмм можно использовать:
  // 1. Ручное создание через меню Google Sheets (Вставка → Диаграмма)
  // 2. Или использовать расширенный Sheets API (требует дополнительной настройки)
  
  // Подготавливаем данные для диаграмм в отдельной области листа
  const chartDataRow = 40
  sheet.getRange(chartDataRow, 1).setValue('Данные для диаграмм (можно использовать для создания графиков вручную)')
  sheet.getRange(chartDataRow, 1).setFontWeight('bold')
  
  // Данные для круговой диаграммы прибыльности
  const profitabilityDataRow = chartDataRow + 2
  sheet.getRange(profitabilityDataRow, 1).setValue('Распределение прибыльности:')
  sheet.getRange(profitabilityDataRow + 1, 1, 1, 2).setValues([['Категория', 'Количество']])
  formatHeaderRange_(sheet.getRange(profitabilityDataRow + 1, 1, 1, 2))
  sheet.getRange(profitabilityDataRow + 2, 1, 3, 2).setValues([
    ['Прибыльные', stats.profitableCount],
    ['Убыточные', stats.unprofitableCount],
    ['Нейтральные', stats.neutralCount]
  ])
  
  // Данные для столбчатой диаграммы топ-5 прибыльных
  if (stats.top5Profitable.length > 0) {
    const top5DataRow = profitabilityDataRow + 6
    sheet.getRange(top5DataRow, 1).setValue('Топ-5 прибыльных (для диаграммы):')
    sheet.getRange(top5DataRow + 1, 1, 1, 2).setValues([['Позиция', 'Прибыль']])
    formatHeaderRange_(sheet.getRange(top5DataRow + 1, 1, 1, 2))
    const top5ChartData = stats.top5Profitable.map(p => [p.name, p.profit])
    sheet.getRange(top5DataRow + 2, 1, top5ChartData.length, 2).setValues(top5ChartData)
    sheet.getRange(top5DataRow + 2, 2, top5ChartData.length, 1).setNumberFormat(NUMBER_FORMATS.CURRENCY)
  }
}

/**
 * Очищает лист (используется когда нет данных)
 */
function portfolioStats_clearSheet_(sheet) {
  sheet.clear()
  sheet.getRange(1, 1).setValue('СТАТИСТИКА ПОРТФЕЛЯ')
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold')
  sheet.getRange(3, 1).setValue('Нет данных в портфеле')
}

/**
 * Форматирование листа (вызывается вручную)
 */
function portfolioStats_formatTable() {
  const sheet = getOrCreatePortfolioStatsSheet_()
  portfolioStats_update()
  SpreadsheetApp.getUi().alert('Аналитика портфеля обновлена')
}

