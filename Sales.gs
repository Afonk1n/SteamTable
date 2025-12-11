// Sales module
// Используем константы из Constants.gs
const SALES_CONFIG = {
  STEAM_APPID: STEAM_APP_ID,
  UPDATE_INTERVAL_MINUTES: UPDATE_INTERVALS.PRICES_MINUTES,
  COLUMNS: SALES_COLUMNS,
}

// Форматирование новой строки Sales (при добавлении из Invest)
function sales_formatNewRow_(sheet, row) {
  const numberFormatConfig = {
    QUANTITY: NUMBER_FORMATS.INTEGER,       // C: Количество
    SELL_PRICE: NUMBER_FORMATS.CURRENCY,    // D: Цена продажи
    CURRENT_PRICE: NUMBER_FORMATS.CURRENCY, // E: Текущая цена
    PRICE_DROP: NUMBER_FORMATS.CURRENCY,    // F: Просадка
    PRICE_DROP_PERCENT: NUMBER_FORMATS.PERCENT, // G: Процент просадки
    MIN_PRICE: NUMBER_FORMATS.CURRENCY,     // I: Min цена
    MAX_PRICE: NUMBER_FORMATS.CURRENCY      // J: Max цена
    // K-S: Buyback Score, Рекомендация, Hero Trend, Метрики, Risk Level
  }
  
  formatNewRowUniversal_(sheet, row, SALES_CONFIG, numberFormatConfig, true)
  
  // Специальное форматирование колонки Потенциал (K) как процент с знаком "+"
  const potentialCol = getColumnIndex(SALES_COLUMNS.POTENTIAL)
  sheet.getRange(row, potentialCol).setNumberFormat('+0%;-0%;"—"')
}

// Функции getSalesSheet_ и getOrCreateSalesSheet_ перенесены в SheetService.gs

function sales_dailyReset() {
  const sheet = getSalesSheet_()
  if (!sheet) return
  const lastRow = sheet.getLastRow()
  if (lastRow <= 1) return

  const rangesToClear = [
    `${SALES_CONFIG.COLUMNS.CURRENT_PRICE}2:${SALES_CONFIG.COLUMNS.PRICE_DROP}${lastRow}`,
  ]
  rangesToClear.forEach(range => sheet.getRange(range).clearContent())

  // ИСПРАВЛЕНИЕ: Синхронизация аналитики убрана отсюда, так как она выполняется в syncPricesFromHistoryToInvestAndSales()
  // Это предотвращает двойную синхронизацию аналитики

  try {
    logAutoAction_('Sales', 'Ежедневный сброс', 'OK')
  } catch (e) {
    console.error('Sales: ошибка при логировании ежедневного сброса:', e)
  }
}

function sales_updateSinglePrice(row) {
  const sheet = getSalesSheet_()
  if (!sheet) return 'error'
  const historySheet = getHistorySheet_()
  if (!historySheet) return 'error'
  
  const itemName = sheet.getRange(`${SALES_CONFIG.COLUMNS.NAME}${row}`).getValue()
  if (!itemName) return 'error'

  const priceResult = getHistoryPriceForPeriod_(historySheet, itemName, getCurrentPricePeriod())
  
  if (!priceResult.found) {
    return 'error'
  }

  const priceColIndex = getColumnIndex(SALES_COLUMNS.CURRENT_PRICE)
  sales_calculateSingle_(sheet, row, priceResult.price)
  
  if (priceResult.isOutdated) {
    sheet.getRange(row, priceColIndex).setBackground(COLORS.STABLE)
  } else {
    sheet.getRange(row, priceColIndex).setBackground(null)
  }
  
  return 'updated'
}

function sales_updateCalculations(row, currentPrice) {
  const sheet = getSalesSheet_()
  if (!sheet) return
  sales_calculateSingle_(sheet, row, currentPrice)
}

function sales_updateImagesAndLinks() {
  updateImagesAndLinksMenu_(SALES_CONFIG, getSalesSheet_, 'Sales')
}

function sales_formatTable() {
  const sheet = getOrCreateSalesSheet_()
  const headers = HEADERS.SALES // 19 колонок (новая структура)
  
  // Базовое форматирование таблицы
  const lastRow = formatTableBase_(sheet, headers, SALES_COLUMNS, getSalesSheet_, 'Sales')
  if (lastRow === 0) return

  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.IMAGE), COLUMN_WIDTHS.IMAGE)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.NAME), COLUMN_WIDTHS.NAME)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.QUANTITY), COLUMN_WIDTHS.MEDIUM) // C
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.SELL_PRICE), COLUMN_WIDTHS.WIDE) // D
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.CURRENT_PRICE), COLUMN_WIDTHS.WIDE) // E
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.PRICE_DROP), COLUMN_WIDTHS.WIDE) // F
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.PRICE_DROP_PERCENT), COLUMN_WIDTHS.WIDE) // G
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.LINK), COLUMN_WIDTHS.NARROW) // H
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.MIN_PRICE), COLUMN_WIDTHS.MEDIUM) // I
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.MAX_PRICE), COLUMN_WIDTHS.MEDIUM) // J
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.BUYBACK_SCORE), 130) // K
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.RECOMMENDATION), COLUMN_WIDTHS.EXTRA_WIDE) // L
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.HERO_TREND), COLUMN_WIDTHS.MEDIUM) // M
  sheet.setColumnWidths(getColumnIndex(SALES_COLUMNS.VOLATILITY_INDEX), 5, COLUMN_WIDTHS.MEDIUM) // N-R (метрики)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.RISK_LEVEL), COLUMN_WIDTHS.MEDIUM) // S

  if (lastRow > 1) {
    sheet.getRange(`C2:F${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // C-F: Количество, Цена продажи, Текущая цена, Просадка
    sheet.getRange(`G2:G${lastRow}`).setNumberFormat(NUMBER_FORMATS.PERCENT) // G: Процент просадки
    sheet.getRange(`I2:J${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // I-J: Min, Max
    // Форматирование метрик (N-R) как число
    sheet.getRange(`N2:R${lastRow}`).setNumberFormat('0.00') // Метрики как числа 0-1

    sheet
      .getRange(DATA_START_ROW, 1, lastRow - 1, headers.length)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')

    sheet.getRange(`B2:B${lastRow}`).setHorizontalAlignment('left')

    const dropRange = sheet.getRange(`F2:G${lastRow}`) // Просадка и Процент просадки
    const recommendationCol = getColumnIndex(SALES_COLUMNS.RECOMMENDATION)
    
    // Условное форматирование для просадки
    const dropPercentCol = getColumnIndex(SALES_COLUMNS.PRICE_DROP_PERCENT)
    const dropPercentRange = sheet.getRange(DATA_START_ROW, dropPercentCol, lastRow - 1, 1)
    
    // Зеленый для положительной просадки (цена выросла)
    const positiveRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([dropPercentRange])
      .whenNumberGreaterThan(0)
      .setBackground(COLORS.PROFIT)
      .build()
    
    // Красный для отрицательной просадки (цена упала)
    const negativeRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([dropPercentRange])
      .whenNumberLessThan(0)
      .setBackground(COLORS.LOSS)
      .build()
    
    sheet.setConditionalFormatRules([positiveRule, negativeRule])
  }
  else {
    sheet.setConditionalFormatRules([])
  }

  // Заморозка строки уже выполнена в formatTableBase_()
  SpreadsheetApp.getUi().alert('Форматирование завершено (Sales)')
}


function sales_findDuplicates() {
  const sheet = getSalesSheet_()
  if (!sheet) return
  const res = highlightDuplicatesByName_(sheet, DATA_START_ROW, COLORS.DUPLICATE)
  SpreadsheetApp.getUi().alert(res.duplicates ? `Найдено повторов: ${res.duplicates}` : 'Повторов не найдено')
}

function sales_syncMinMaxFromHistory(updateAll = true) {
  const sheet = getSalesSheet_()
  if (!sheet) return

  // SALES_COLUMNS.MIN_PRICE = 'I', SALES_COLUMNS.MAX_PRICE = 'J'
  const minColIndex = getColumnIndex(SALES_COLUMNS.MIN_PRICE)
  const maxColIndex = getColumnIndex(SALES_COLUMNS.MAX_PRICE)
  
  return syncMinMaxFromHistoryUniversal_(sheet, minColIndex, maxColIndex, updateAll)
}

// Синхронизация Тренд/Дней смены из листа History по названию
function sales_syncTrendDaysFromHistory(updateAll = true) {
  const sheet = getSalesSheet_()
  if (!sheet) return

  // SALES_COLUMNS.TREND больше не используется в новой структуре (удалена)
  // Тренд синхронизируется из History, но в Sales нет отдельной колонки для него
  return true
  
  return syncTrendFromHistoryUniversal_(sheet, trendColIndex, updateAll)
}

// Синхронизация расширенной аналитики (Фаза/Потенциал/Рекомендация) из History
function sales_syncExtendedAnalyticsFromHistory(updateAll = true) {
  const sheet = getSalesSheet_()
  if (!sheet) return

  // SALES_COLUMNS: PHASE и POTENTIAL больше не используются в новой структуре
  // SALES_COLUMNS.RECOMMENDATION = 'L'
  const recommendationColIndex = getColumnIndex(SALES_COLUMNS.RECOMMENDATION)
  
  // Синхронизируем только Рекомендацию из History
  return syncRecommendationFromHistoryUniversal_(sheet, recommendationColIndex, updateAll)
}

/**
 * Комплексное обновление всей аналитики (Min/Max + Рекомендация)
 */
function sales_updateAllAnalytics() {
  updateAllAnalyticsManual_(
    'Sales',
    sales_syncMinMaxFromHistory,
    sales_syncTrendDaysFromHistory,
    sales_syncExtendedAnalyticsFromHistory
  )
}

// ===== СИСТЕМА ИНВЕСТИЦИОННЫХ РЕКОМЕНДАЦИЙ =====

/**
 * Получение данных из SteamWebAPI для расчета метрик
 * @param {Array<string>} itemNames - Массив названий предметов
 * @returns {Object} Объект {itemName: itemData}
 */
function sales_updateMetricsFromSteamWebAPI(itemNames) {
  const itemsData = {}
  
  // Batch запросы (до 50 предметов за раз)
  const batchSize = API_CONFIG.STEAM_WEB_API.MAX_ITEMS_PER_REQUEST
  for (let i = 0; i < itemNames.length; i += batchSize) {
    const batch = itemNames.slice(i, i + batchSize)
    const result = steamWebAPI_fetchItems(batch, 'dota2')
    if (result.ok && result.items) {
      result.items.forEach(item => {
        if (item.marketname) {
          itemsData[item.marketname] = item
        }
      })
    }
    // Задержка между batch запросами
    if (i + batchSize < itemNames.length) {
      Utilities.sleep(500)
    }
  }
  
  return itemsData
}

/**
 * Расчет всех метрик для позиций в Sales
 */
function sales_calculateAllMetrics() {
  const sheet = getSalesSheet_()
  if (!sheet) return
  
  const lastRow = sheet.getLastRow()
  if (lastRow < DATA_START_ROW) return
  
  const itemNames = sheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.NAME), lastRow - HEADER_ROW, 1).getValues()
  const itemNamesList = itemNames.map(row => String(row[0] || '').trim()).filter(name => name)
  
  // Получаем данные из SteamWebAPI
  const itemsData = sales_updateMetricsFromSteamWebAPI(itemNamesList)
  
  // Получаем маппинги героев
  const mappings = heroMapping_getAllMappings()
  
  // Получаем историю цен из History
  const historySheet = getHistorySheet_()
  
  // Обновляем метрики для каждой строки
  for (let i = 0; i < itemNames.length; i++) {
    const itemName = String(itemNames[i][0] || '').trim()
    if (!itemName) continue
    
    const row = DATA_START_ROW + i
    const itemData = itemsData[itemName]
    if (!itemData) continue
    
    const mapping = mappings[itemName]
    const heroId = mapping && mapping.heroId ? mapping.heroId : null
    const rankCategory = mapping && mapping.heroId ? 'High Rank' : null
    
    // Получаем историю цен
    let historyData = null
    if (historySheet) {
      const historyRow = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), historySheet.getLastRow() - HEADER_ROW, 1).getValues().findIndex(r => String(r[0] || '').trim() === itemName)
      if (historyRow >= 0) {
        historyData = history_getPriceHistoryForItem_(historySheet, historyRow + DATA_START_ROW)
      }
    }
    
    // Рассчитываем метрики
    const liquidityScore = analytics_calculateLiquidityScore(itemData)
    const demandRatio = analytics_calculateDemandRatio(itemData)
    const priceMomentum = analytics_calculatePriceMomentum(itemData, historyData)
    const salesTrend = analytics_calculateSalesTrend(itemData)
    const volatilityIndex = analytics_calculateVolatilityIndex(itemData, historyData)
    
    // Обновляем колонки метрик
    sheet.getRange(row, getColumnIndex(SALES_COLUMNS.LIQUIDITY_SCORE)).setValue(liquidityScore)
    sheet.getRange(row, getColumnIndex(SALES_COLUMNS.DEMAND_RATIO)).setValue(demandRatio)
    sheet.getRange(row, getColumnIndex(SALES_COLUMNS.PRICE_MOMENTUM)).setValue(priceMomentum)
    sheet.getRange(row, getColumnIndex(SALES_COLUMNS.SALES_TREND)).setValue(salesTrend)
    sheet.getRange(row, getColumnIndex(SALES_COLUMNS.VOLATILITY_INDEX)).setValue(volatilityIndex)
    
    // Hero Trend Score (только для Hero Items)
    if (heroId && rankCategory) {
      const latestStats = heroStats_getLatestStats(heroId, rankCategory)
      if (latestStats) {
        const heroStatsObj = {[rankCategory]: latestStats}
        const heroTrendScore = analytics_calculateHeroTrendScore(heroId, rankCategory, heroStatsObj)
        sheet.getRange(row, getColumnIndex(SALES_COLUMNS.HERO_TREND)).setValue(analytics_formatScore(heroTrendScore))
      }
    }
  }
}

/**
 * Расчет Buyback Score для всех позиций в Sales
 */
function sales_updateBuybackScores() {
  const sheet = getSalesSheet_()
  if (!sheet) return
  
  const lastRow = sheet.getLastRow()
  if (lastRow < DATA_START_ROW) return
  
  const itemNames = sheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.NAME), lastRow - HEADER_ROW, 1).getValues()
  const itemNamesList = itemNames.map(row => String(row[0] || '').trim()).filter(name => name)
  
  // Получаем данные из SteamWebAPI
  const itemsData = sales_updateMetricsFromSteamWebAPI(itemNamesList)
  
  // Получаем маппинги героев
  const mappings = heroMapping_getAllMappings()
  
  // Получаем историю цен из History
  const historySheet = getHistorySheet_()
  
  // Обновляем Buyback Score для каждой строки
  for (let i = 0; i < itemNames.length; i++) {
    const itemName = String(itemNames[i][0] || '').trim()
    if (!itemName) continue
    
    const row = DATA_START_ROW + i
    const itemData = itemsData[itemName]
    if (!itemData) continue
    
    const mapping = mappings[itemName]
    const heroId = mapping && mapping.heroId ? mapping.heroId : null
    const rankCategory = mapping && mapping.heroId ? 'High Rank' : null
    
    // Получаем цену продажи и текущую цену
    const sellPrice = Number(sheet.getRange(row, getColumnIndex(SALES_COLUMNS.SELL_PRICE)).getValue()) || 0
    const currentPrice = Number(sheet.getRange(row, getColumnIndex(SALES_COLUMNS.CURRENT_PRICE)).getValue()) || 0
    
    if (!sellPrice || !currentPrice) continue
    
    // Получаем историю цен
    let historyData = null
    if (historySheet) {
      const historyRow = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), historySheet.getLastRow() - HEADER_ROW, 1).getValues().findIndex(r => String(r[0] || '').trim() === itemName)
      if (historyRow >= 0) {
        historyData = history_getPriceHistoryForItem_(historySheet, historyRow + DATA_START_ROW)
      }
    }
    
    // Получаем статистику героя
    let heroStats = null
    if (heroId && rankCategory) {
      const latestStats = heroStats_getLatestStats(heroId, rankCategory)
      if (latestStats) {
        heroStats = {[rankCategory]: latestStats}
      }
    }
    
    // Рассчитываем Buyback Score
    const buybackScore = analytics_calculateBuybackScore(
      itemData,
      heroStats,
      historyData,
      sellPrice,
      currentPrice,
      heroId,
      rankCategory
    )
    
    // Обновляем колонку Buyback Score
    sheet.getRange(row, getColumnIndex(SALES_COLUMNS.BUYBACK_SCORE))
      .setValue(analytics_formatScore(buybackScore))
    
    // Рассчитываем Risk Level
    const volatilityIndex = analytics_calculateVolatilityIndex(itemData, historyData)
    const demandRatio = analytics_calculateDemandRatio(itemData)
    const riskLevel = analytics_calculateRiskLevel(buybackScore, volatilityIndex, demandRatio)
    sheet.getRange(row, getColumnIndex(SALES_COLUMNS.RISK_LEVEL)).setValue(riskLevel)
  }
}

/**
 * Генерация рекомендации на основе Buyback Score
 * @param {number} row - Номер строки
 * @returns {string} Рекомендация
 */
function sales_generateRecommendation(row) {
  const sheet = getSalesSheet_()
  if (!sheet) return '👀 НАБЛЮДАТЬ'
  
  const buybackScoreStr = String(sheet.getRange(row, getColumnIndex(SALES_COLUMNS.BUYBACK_SCORE)).getValue() || '').trim()
  if (!buybackScoreStr || buybackScoreStr === '—') return '👀 НАБЛЮДАТЬ'
  
  // Парсим число из формата "🟩 0.93"
  const scoreMatch = buybackScoreStr.match(/(\d+\.?\d*)/)
  if (!scoreMatch) return '👀 НАБЛЮДАТЬ'
  
  const buybackScore = parseFloat(scoreMatch[1])
  
  const priceDropPercent = Number(sheet.getRange(row, getColumnIndex(SALES_COLUMNS.PRICE_DROP_PERCENT)).getValue()) || 0
  
  if (buybackScore >= 0.75) {
    return `💰 ОТКУПИТЬ (Score: ${(buybackScore * 100).toFixed(0)}%, Просадка: ${(priceDropPercent * 100).toFixed(1)}%)`
  }
  if (buybackScore >= 0.60) {
    return `🟨 РАССМОТРЕТЬ (Score: ${(buybackScore * 100).toFixed(0)}%)`
  }
  if (buybackScore < 0.40) {
    return `🟥 НЕ ОТКУПАТЬ (Score: ${(buybackScore * 100).toFixed(0)}%)`
  }
  return `👀 НАБЛЮДАТЬ (Score: ${(buybackScore * 100).toFixed(0)}%)`
}
