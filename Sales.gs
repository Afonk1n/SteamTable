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
    SELL_PRICE: NUMBER_FORMATS.CURRENCY,    // C: Цена продажи
    CURRENT_PRICE: NUMBER_FORMATS.CURRENCY, // D: Текущая цена
    PRICE_DROP: NUMBER_FORMATS.CURRENCY,    // E: Просадка
    PRICE_DROP_PERCENT: NUMBER_FORMATS.PERCENT, // F: Процент просадки
    MIN_PRICE: NUMBER_FORMATS.CURRENCY,     // H: Min цена
    MAX_PRICE: NUMBER_FORMATS.CURRENCY      // I: Max цена
    // J-M: Buyback Score, Рекомендация, Hero Trend, Risk Level
  }
  
  formatNewRowUniversal_(sheet, row, SALES_CONFIG, numberFormatConfig, true)
  
  // ПРИМЕЧАНИЕ: В Sales нет колонки QUANTITY и POTENTIAL, эти поля были удалены из структуры
  // Если нужно добавить форматирование для других колонок, используйте соответствующие поля из SALES_COLUMNS
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
  const headers = HEADERS.SALES // 13 колонок (без колонки Количество)
  
  if (!headers || !Array.isArray(headers) || headers.length === 0) {
    console.error('Sales: HEADERS.SALES не определен или пуст')
    SpreadsheetApp.getUi().alert('Ошибка: HEADERS.SALES не определен в Constants.gs')
    return
  }
  
  // Проверка SALES_COLUMNS перед использованием
  if (!SALES_COLUMNS) {
    console.error('Sales: SALES_COLUMNS не определен')
    SpreadsheetApp.getUi().alert('Ошибка: SALES_COLUMNS не определен в Constants.gs')
    return
  }
  
  // Базовое форматирование таблицы
  const lastRow = formatTableBase_(sheet, headers, SALES_COLUMNS, getSalesSheet_, 'Sales')
  if (lastRow === 0) return

  // Проверка COLUMN_WIDTHS перед использованием
  if (!COLUMN_WIDTHS) {
    console.error('Sales: COLUMN_WIDTHS не определен')
    SpreadsheetApp.getUi().alert('Ошибка: COLUMN_WIDTHS не определен в Constants.gs')
    return
  }
  
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.IMAGE), COLUMN_WIDTHS.IMAGE)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.NAME), COLUMN_WIDTHS.NAME)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.SELL_PRICE), COLUMN_WIDTHS.WIDE) // C
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.CURRENT_PRICE), COLUMN_WIDTHS.WIDE) // D
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.PRICE_DROP), COLUMN_WIDTHS.WIDE) // E
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.PRICE_DROP_PERCENT), COLUMN_WIDTHS.WIDE) // F
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.LINK), COLUMN_WIDTHS.NARROW) // G
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.MIN_PRICE), COLUMN_WIDTHS.MEDIUM) // H
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.MAX_PRICE), COLUMN_WIDTHS.MEDIUM) // I
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.BUYBACK_SCORE), 130) // J
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.RECOMMENDATION), COLUMN_WIDTHS.EXTRA_WIDE) // K
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.HERO_TREND), COLUMN_WIDTHS.MEDIUM) // L
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.RISK_LEVEL), COLUMN_WIDTHS.MEDIUM) // M

  if (lastRow > 1) {
    // Дополнительная проверка headers перед использованием
    if (!headers || !Array.isArray(headers)) {
      console.error('Sales: headers потеряны после formatTableBase_')
      SpreadsheetApp.getUi().alert('Ошибка: заголовки Sales потеряны. Попробуйте запустить форматирование еще раз.')
      return
    }
    
    sheet.getRange(`C2:F${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // C-F: Количество, Цена продажи, Текущая цена, Просадка
    sheet.getRange(`G2:G${lastRow}`).setNumberFormat(NUMBER_FORMATS.PERCENT) // G: Процент просадки
    sheet.getRange(`I2:J${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // I-J: Min, Max
    // Метрики удалены из отображения (остаются в коде для расчетов)

    sheet
      .getRange(DATA_START_ROW, 1, lastRow - 1, headers.length)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')

    sheet.getRange(`B2:B${lastRow}`).setHorizontalAlignment('left')

    const dropRange = sheet.getRange(`F2:G${lastRow}`) // Просадка и Процент просадки
    
    // Проверка SALES_COLUMNS перед использованием
    if (!SALES_COLUMNS) {
      console.error('Sales: SALES_COLUMNS не определен')
      SpreadsheetApp.getUi().alert('Ошибка: SALES_COLUMNS не определен в Constants.gs')
      return
    }
    
    const recommendationCol = getColumnIndex(SALES_COLUMNS.RECOMMENDATION)
    if (recommendationCol <= 0) {
      console.error('Sales: не удалось определить колонку RECOMMENDATION')
    }
    
    // Условное форматирование для просадки
    const dropPercentCol = getColumnIndex(SALES_COLUMNS.PRICE_DROP_PERCENT)
    if (dropPercentCol <= 0) {
      console.error('Sales: не удалось определить колонку PRICE_DROP_PERCENT')
      return
    }
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
  console.log('Sales: форматирование завершено')
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
}

// Синхронизация расширенной аналитики (Фаза/Потенциал/Рекомендация) из History
function sales_syncExtendedAnalyticsFromHistory(updateAll = true) {
  const sheet = getSalesSheet_()
  if (!sheet) return

  // SALES_COLUMNS: PHASE и POTENTIAL больше не используются в новой структуре
  // SALES_COLUMNS.RECOMMENDATION = 'L'
  const recommendationColIndex = getColumnIndex(SALES_COLUMNS.RECOMMENDATION)
  
  // Синхронизируем только Рекомендацию из History
  // Используем универсальную функцию с null для phase и potential
  return syncExtendedAnalyticsFromHistoryUniversal_(sheet, null, null, recommendationColIndex, updateAll)
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
      Utilities.sleep(LIMITS.METRICS_UPDATE_DELAY_MS)
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
  
  // Подготовка данных для batch-операций
  const liquidityScores = []
  const demandRatios = []
  const priceMomenta = []
  const salesTrends = []
  const volatilityIndices = []
  const heroTrends = []
  const historyNames = historySheet ? historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), historySheet.getLastRow() - HEADER_ROW, 1).getValues() : []
  
  // Рассчитываем метрики для всех строк
  const startedAt = Date.now()
  const MAX_EXECUTION_TIME_MS = 300000 // 5 минут
  
  for (let i = 0; i < itemNames.length; i++) {
    // Проверка таймаута
    if (Date.now() - startedAt > MAX_EXECUTION_TIME_MS) {
      console.warn(`Sales: превышено время выполнения calculateAllMetrics (${MAX_EXECUTION_TIME_MS}ms), прервано на строке ${i + 1}`)
      break
    }
    
    try {
      const itemName = String(itemNames[i][0] || '').trim()
      if (!itemName) {
        liquidityScores.push([null])
        demandRatios.push([null])
        priceMomenta.push([null])
        salesTrends.push([null])
        volatilityIndices.push([null])
        heroTrends.push([null])
        continue
      }
      
      const itemData = itemsData[itemName]
      if (!itemData) {
        console.warn(`Sales: нет данных из SteamWebAPI для "${itemName}"`)
        liquidityScores.push([null])
        demandRatios.push([null])
        priceMomenta.push([null])
        salesTrends.push([null])
        volatilityIndices.push([null])
        heroTrends.push([null])
        continue
      }
      
      // ВАЛИДАЦИЯ: Проверяем валидность цен в itemData
      if (itemData.pricelatest !== undefined && itemData.pricelatest !== null) {
        const priceValidation = validatePrice_(itemData.pricelatest, `${itemName} (pricelatest)`)
        if (!priceValidation.valid) {
          console.warn(`Sales: некорректная цена pricelatest для "${itemName}": ${itemData.pricelatest}, пропускаем расчет метрик`)
          liquidityScores.push([null])
          demandRatios.push([null])
          priceMomenta.push([null])
          salesTrends.push([null])
          volatilityIndices.push([null])
          heroTrends.push([null])
          continue
        }
      }
      
      const mapping = mappings[itemName]
      const heroId = mapping && mapping.heroId ? mapping.heroId : null
      const rankCategory = mapping && mapping.heroId ? 'High Rank' : null
      
      // Получаем историю цен
      let historyData = null
      if (historySheet && historyNames.length > 0) {
        const historyRowIndex = historyNames.findIndex(r => String(r[0] || '').trim() === itemName)
        if (historyRowIndex >= 0) {
          historyData = history_getPriceHistoryForItem_(historySheet, historyRowIndex + DATA_START_ROW)
        } else {
          console.warn(`Sales: предмет "${itemName}" не найден в History`)
        }
      }
      
      // Рассчитываем метрики с обработкой ошибок
      try {
        liquidityScores.push([analytics_calculateLiquidityScore(itemData)])
        demandRatios.push([analytics_calculateDemandRatio(itemData)])
        priceMomenta.push([analytics_calculatePriceMomentum(itemData, historyData)])
        salesTrends.push([analytics_calculateSalesTrend(itemData)])
        volatilityIndices.push([analytics_calculateVolatilityIndex(itemData, historyData)])
      } catch (e) {
        console.error(`Sales: ошибка расчета метрик для "${itemName}":`, e)
        liquidityScores.push([null])
        demandRatios.push([null])
        priceMomenta.push([null])
        salesTrends.push([null])
        volatilityIndices.push([null])
      }
      
      // Hero Trend Score (только для Hero Items)
      let heroTrendValue = null
      if (heroId && rankCategory) {
        try {
          const latestStats = heroStats_getLatestStats(heroId, rankCategory)
          if (latestStats) {
            const heroStatsObj = {[rankCategory]: latestStats}
            const heroTrendScore = analytics_calculateHeroTrendScore(heroId, rankCategory, heroStatsObj)
            heroTrendValue = analytics_formatScore(heroTrendScore)
          }
        } catch (e) {
          console.error(`Sales: ошибка расчета Hero Trend для "${itemName}":`, e)
        }
      }
      heroTrends.push([heroTrendValue])
      
    } catch (e) {
      console.error(`Sales: ошибка обработки строки ${i + 1} в calculateAllMetrics:`, e)
      liquidityScores.push([null])
      demandRatios.push([null])
      priceMomenta.push([null])
      salesTrends.push([null])
      volatilityIndices.push([null])
      heroTrends.push([null])
    }
  }
  
  // Batch-запись Hero Trend (метрики удалены из отображения, но расчеты остаются для Buyback Score)
  const count = heroTrends.length
  if (count > 0) {
    sheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.HERO_TREND), count, 1).setValues(heroTrends)
  }
  // Метрики (liquidityScores, demandRatios, priceMomenta, salesTrends, volatilityIndices) 
  // рассчитываются, но не записываются в таблицу - используются только для расчета Buyback Score
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
  
  // Читаем цены продажи и текущие цены batch-операцией
  const sellPrices = sheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.SELL_PRICE), lastRow - HEADER_ROW, 1).getValues()
  const currentPrices = sheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.CURRENT_PRICE), lastRow - HEADER_ROW, 1).getValues()
  const historyNames = historySheet ? historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), historySheet.getLastRow() - HEADER_ROW, 1).getValues() : []
  
  // Подготовка данных для batch-операций
  const buybackScores = []
  const riskLevels = []
  
  // Рассчитываем Buyback Score и Risk Level для всех строк
  const startedAt = Date.now()
  const MAX_EXECUTION_TIME_MS = 300000 // 5 минут
  
  for (let i = 0; i < itemNames.length; i++) {
    // Проверка таймаута
    if (Date.now() - startedAt > MAX_EXECUTION_TIME_MS) {
      console.warn(`Sales: превышено время выполнения updateBuybackScores (${MAX_EXECUTION_TIME_MS}ms), прервано на строке ${i + 1}`)
      break
    }
    
    try {
      const itemName = String(itemNames[i][0] || '').trim()
      if (!itemName) {
        buybackScores.push([null])
        riskLevels.push([null])
        continue
      }
      
      const itemData = itemsData[itemName]
      if (!itemData) {
        console.warn(`Sales: нет данных из SteamWebAPI для "${itemName}"`)
        buybackScores.push([null])
        riskLevels.push([null])
        continue
      }
      
      // Получаем цену продажи и текущую цену
      const sellPrice = Number(sellPrices[i][0]) || 0
      const currentPrice = Number(currentPrices[i][0]) || 0
      
      // ВАЛИДАЦИЯ: Проверяем цены перед использованием
      const sellPriceValidation = sellPrice > 0 ? validatePrice_(sellPrice, `${itemName} (sellPrice)`) : { valid: false }
      const currentPriceValidation = currentPrice > 0 ? validatePrice_(currentPrice, `${itemName} (currentPrice)`) : { valid: false }
      
      if (!sellPriceValidation.valid || !currentPriceValidation.valid) {
        console.warn(`Sales: некорректные цены для "${itemName}": sellPrice=${sellPrice}, currentPrice=${currentPrice}`)
        buybackScores.push([null])
        riskLevels.push([null])
        continue
      }
      
      const mapping = mappings[itemName]
      const heroId = mapping && mapping.heroId ? mapping.heroId : null
      const rankCategory = mapping && mapping.heroId ? 'High Rank' : null
      
      // Получаем историю цен
      let historyData = null
      if (historySheet && historyNames.length > 0) {
        const historyRowIndex = historyNames.findIndex(r => String(r[0] || '').trim() === itemName)
        if (historyRowIndex >= 0) {
          historyData = history_getPriceHistoryForItem_(historySheet, historyRowIndex + DATA_START_ROW)
        } else {
          console.warn(`Sales: предмет "${itemName}" не найден в History`)
        }
      }
      
      // Получаем статистику героя
      let heroStats = null
      if (heroId && rankCategory) {
        try {
          const latestStats = heroStats_getLatestStats(heroId, rankCategory)
          if (latestStats) {
            heroStats = {[rankCategory]: latestStats}
          }
        } catch (e) {
          console.error(`Sales: ошибка получения статистики героя для "${itemName}":`, e)
        }
      }
      
      // Рассчитываем Buyback Score
      let buybackScore = 0.5 // Значение по умолчанию
      try {
        buybackScore = analytics_calculateBuybackScore(
          itemData,
          heroStats,
          historyData,
          sellPriceValidation.price,
          currentPriceValidation.price,
          heroId,
          rankCategory
        )
        // Валидация Buyback Score
        if (!Number.isFinite(buybackScore) || buybackScore < 0 || buybackScore > 1) {
          console.warn(`Sales: некорректный Buyback Score для "${itemName}": ${buybackScore}, используем значение по умолчанию`)
          buybackScore = 0.5
        }
      } catch (e) {
        console.error(`Sales: ошибка расчета Buyback Score для "${itemName}":`, e)
        buybackScore = 0.5
      }
      
      buybackScores.push([analytics_formatScore(buybackScore)])
      
      // Рассчитываем Risk Level
      let riskLevel = 'Medium' // Значение по умолчанию
      try {
        const volatilityIndex = analytics_calculateVolatilityIndex(itemData, historyData)
        const demandRatio = analytics_calculateDemandRatio(itemData)
        riskLevel = analytics_calculateRiskLevel(buybackScore, volatilityIndex, demandRatio)
      } catch (e) {
        console.error(`Sales: ошибка расчета Risk Level для "${itemName}":`, e)
      }
      riskLevels.push([riskLevel])
      
    } catch (e) {
      console.error(`Sales: ошибка обработки строки ${i + 1} в updateBuybackScores:`, e)
      buybackScores.push([null])
      riskLevels.push([null])
    }
  }
  
  // Batch-запись Buyback Scores и Risk Levels
  const count = buybackScores.length
  if (count > 0) {
    sheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.BUYBACK_SCORE), count, 1).setValues(buybackScores)
    sheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.RISK_LEVEL), count, 1).setValues(riskLevels)
  }
  
  // Проверяем возможности для откупа (критические уведомления)
  try {
    telegram_checkSalesBuybackOpportunities_()
  } catch (e) {
    console.error('Sales: ошибка при проверке возможностей для откупа:', e)
    // Не прерываем выполнение, просто логируем ошибку
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

