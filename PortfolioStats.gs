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
  
  // Рассчитываем метрики (теперь с именами для изображений)
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
    const trendStr = String(trend || '').trim()
    let trendEmoji = '🟪' // По умолчанию
    
    if (trendStr) {
      // Проверяем начало строки на эмодзи тренда
      if (trendStr.startsWith('🟩')) {
        trendEmoji = '🟩'
      } else if (trendStr.startsWith('🟥')) {
        trendEmoji = '🟥'
      } else if (trendStr.startsWith('🟨')) {
        trendEmoji = '🟨'
      } else if (trendStr.startsWith('🟪')) {
        trendEmoji = '🟪'
      } else {
        // Пробуем через regex как fallback
        const trendMatch = trendStr.match(/^([🟥🟩🟨🟪])/)
        if (trendMatch) {
          trendEmoji = trendMatch[1]
        }
      }
    }
    
    trendDistribution.set(trendEmoji, (trendDistribution.get(trendEmoji) || 0) + 1)
  }
  
  // Сортируем позиции по проценту прибыльности (а не по рублям)
  positions.sort((a, b) => b.profitPercent - a.profitPercent)
  const top5Profitable = positions.filter(p => p.profitPercent > 0).slice(0, 5)
  const top5Unprofitable = positions.filter(p => p.profitPercent < 0).sort((a, b) => a.profitPercent - b.profitPercent).slice(0, 5)
  
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
  
  // Общие показатели - оформляем в таблицу
  let row = 3
  sheet.getRange(row, 1).setValue('Общие показатели').setFontWeight('bold').setFontSize(12)
  row++
  
  // Заголовки таблицы
  sheet.getRange(row, 1, 1, 2).setValues([['Показатель', 'Значение']])
  formatHeaderRange_(sheet.getRange(row, 1, 1, 2))
  row++
  
  // Данные таблицы
  const generalData = [
    ['Общая сумма вложений', stats.totalInvestment],
    ['Текущая стоимость', stats.totalCurrentValue],
    ['Прибыль/убыток', stats.totalProfit],
    ['Прибыль/убыток (%)', stats.totalProfitPercent],
    ['Средняя прибыльность', stats.avgProfitability],
    ['Активных позиций', stats.totalPositions]
  ]
  
  sheet.getRange(row, 1, generalData.length, 2).setValues(generalData)
  // Форматирование чисел
  sheet.getRange(row, 2, 2, 1).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Вложения и стоимость
  sheet.getRange(row + 2, 2, 1, 1).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Прибыль/убыток
  sheet.getRange(row + 3, 2, 1, 1).setNumberFormat(NUMBER_FORMATS.PERCENT) // Прибыль/убыток %
  sheet.getRange(row + 4, 2, 1, 1).setNumberFormat(NUMBER_FORMATS.PERCENT) // Средняя прибыльность
  sheet.getRange(row + 5, 2, 1, 1).setNumberFormat(NUMBER_FORMATS.INTEGER) // Количество позиций
  
  // Цветовая индикация для прибыли/убытка
  if (stats.totalProfit > 0) {
    sheet.getRange(row + 2, 1, 1, 2).setBackground(COLORS.PROFIT)
  } else if (stats.totalProfit < 0) {
    sheet.getRange(row + 2, 1, 1, 2).setBackground(COLORS.LOSS)
  }
  
  row += generalData.length
  
  // Топ-5 прибыльных
  row += 2
  sheet.getRange(row, 1).setValue('Топ-5 прибыльных').setFontWeight('bold').setFontSize(12)
  row++
  
  // Заголовки таблицы (с изображением)
  sheet.getRange(row, 1, 1, 5).setValues([['Изображение', 'Название', 'Прибыль', 'Прибыль %', 'Вложения']])
  formatHeaderRange_(sheet.getRange(row, 1, 1, 5))
  row++
  
  // Данные (с изображениями)
  const profitableData = stats.top5Profitable.map(p => [
    '', // Изображение будет добавлено отдельно
    p.name,
    p.profit,
    p.profitPercent,
    p.investment
  ])
  if (profitableData.length > 0) {
    sheet.getRange(row, 1, profitableData.length, 5).setValues(profitableData)
    // Добавляем изображения
    for (let i = 0; i < profitableData.length; i++) {
      const item = stats.top5Profitable[i]
      if (item && item.name) {
        const imageFormula = buildImageAndLinkFormula_(STEAM_APP_ID, item.name).image
        sheet.getRange(row + i, 1).setFormula(imageFormula)
      }
    }
    sheet.getRange(row, 2, profitableData.length, 1).setNumberFormat(NUMBER_FORMATS.CURRENCY)
    sheet.getRange(row, 3, profitableData.length, 1).setNumberFormat(NUMBER_FORMATS.PERCENT)
    sheet.getRange(row, 4, profitableData.length, 1).setNumberFormat(NUMBER_FORMATS.CURRENCY)
    sheet.getRange(row, 1, profitableData.length, 5).setBackground(COLORS.PROFIT)
    // Устанавливаем ширину колонки изображения
    sheet.setColumnWidth(1, COLUMN_WIDTHS.IMAGE)
    row += profitableData.length
  } else {
    sheet.getRange(row, 1).setValue('Нет данных')
    row++
  }
  
  // Топ-5 убыточных
  row += 2
  sheet.getRange(row, 1).setValue('Топ-5 убыточных').setFontWeight('bold').setFontSize(12)
  row++
  
  // Заголовки таблицы (с изображением)
  sheet.getRange(row, 1, 1, 5).setValues([['Изображение', 'Название', 'Убыток', 'Убыток %', 'Вложения']])
  formatHeaderRange_(sheet.getRange(row, 1, 1, 5))
  row++
  
  // Данные (с изображениями)
  const unprofitableData = stats.top5Unprofitable.map(p => [
    '', // Изображение будет добавлено отдельно
    p.name,
    p.profit,
    p.profitPercent,
    p.investment
  ])
  if (unprofitableData.length > 0) {
    sheet.getRange(row, 1, unprofitableData.length, 5).setValues(unprofitableData)
    // Добавляем изображения
    for (let i = 0; i < unprofitableData.length; i++) {
      const item = stats.top5Unprofitable[i]
      if (item && item.name) {
        const imageFormula = buildImageAndLinkFormula_(STEAM_APP_ID, item.name).image
        sheet.getRange(row + i, 1).setFormula(imageFormula)
      }
    }
    sheet.getRange(row, 2, unprofitableData.length, 1).setNumberFormat(NUMBER_FORMATS.CURRENCY)
    sheet.getRange(row, 3, unprofitableData.length, 1).setNumberFormat(NUMBER_FORMATS.PERCENT)
    sheet.getRange(row, 4, unprofitableData.length, 1).setNumberFormat(NUMBER_FORMATS.CURRENCY)
    sheet.getRange(row, 1, unprofitableData.length, 5).setBackground(COLORS.LOSS)
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
  
  // Форматирование колонок (с учетом колонки изображений в топ-5)
  // Колонка изображений уже установлена выше (COLUMN_WIDTHS.IMAGE)
  // Остальные колонки форматируются автоматически через setColumnWidths в начале функции
  
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

