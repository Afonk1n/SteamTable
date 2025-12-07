// Общие утилиты для всех модулей

function parseLocalizedPrice_(raw) {
  if (raw == null) return { ok: false, error: 'no_price' }
  const normalized = String(raw).replace(/[^0-9.,]+/g, '').replace(',', '.')
  const price = parseFloat(normalized)
  return Number.isFinite(price) ? { ok: true, price } : { ok: false, error: 'nan' }
}

function fetchLowestPrice_(appid, itemName) {
  const name = String(itemName || '').trim()
  if (!name) return { ok: false, error: 'empty_name' }
  const url = `https://steamcommunity.com/market/priceoverview/?currency=5&appid=${appid}&market_hash_name=${encodeURIComponent(name)}`
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true })
    if (response.getResponseCode() !== 200) return { ok: false, error: 'http_' + response.getResponseCode() }
    const data = JSON.parse(response.getContentText())
    if (data && data.success && data.lowest_price) {
      const parsed = parseLocalizedPrice_(data.lowest_price)
      return parsed.ok ? { ok: true, price: parsed.price } : { ok: false, error: parsed.error }
    }
    return { ok: false, error: 'no_data' }
  } catch (e) {
    return { ok: false, error: 'exception' }
  }
}


// Система блокировки для предотвращения параллельного запуска
function acquireLock_(lockKey, timeoutSeconds = LIMITS.LOCK_TIMEOUT_SEC) {
  const props = PropertiesService.getScriptProperties()
  const now = Date.now()
  const lockInfo = props.getProperty(lockKey)
  
  if (lockInfo) {
    const { timestamp, pid } = JSON.parse(lockInfo)
    const lockAge = (now - timestamp) / 1000
    
    // Если блокировка старше timeout - считаем мертвой и перехватываем
    if (lockAge < timeoutSeconds) {
      return { locked: true, age: lockAge }
    }
  }
  
  // Создаем блокировку
  const lockData = JSON.stringify({ timestamp: now, pid: Utilities.getUuid() })
  props.setProperty(lockKey, lockData)
  return { locked: false }
}

function releaseLock_(lockKey) {
  const props = PropertiesService.getScriptProperties()
  props.deleteProperty(lockKey)
}

// === СИСТЕМА УПРАВЛЕНИЯ ПЕРИОДАМИ СБОРА ЦЕН ===

// Определяет текущий период сбора цен на основе времени
function getCurrentPricePeriod() {
  const now = new Date()
  const hour = now.getHours()
  const minutes = now.getMinutes()
  const currentTimeMinutes = hour * 60 + minutes
  
  const morningStartMinutes = UPDATE_INTERVALS.MORNING_HOUR * 60 + UPDATE_INTERVALS.MORNING_MINUTE
  const eveningStartMinutes = UPDATE_INTERVALS.EVENING_HOUR * 60 + UPDATE_INTERVALS.EVENING_MINUTE
  
  // Период "ночь" (для пользователя): с 00:10 до 12:00
  // Проверка с учетом минут для точного определения периода
  if (currentTimeMinutes >= morningStartMinutes && currentTimeMinutes < eveningStartMinutes) {
    return PRICE_COLLECTION_PERIODS.MORNING
  }
  // Период "день" (для пользователя): с 12:00 до 00:10 следующего дня
  return PRICE_COLLECTION_PERIODS.EVENING
}

// Получает состояние периода сбора цен
function getPriceCollectionState() {
  const props = PropertiesService.getScriptProperties()
  const period = props.getProperty('price_collection_period') || ''
  const date = props.getProperty('price_collection_date') || ''
  const completed = props.getProperty('price_collection_completed') === 'true'
  
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  const isToday = date === todayStr
  
  return {
    period,
    date,
    completed,
    isToday
  }
}

// Сохраняет состояние периода сбора цен
function setPriceCollectionState(period, completed = false) {
  const props = PropertiesService.getScriptProperties()
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  
  props.setProperty('price_collection_period', period)
  props.setProperty('price_collection_date', todayStr)
  props.setProperty('price_collection_completed', completed ? 'true' : 'false')
}

// Проверяет, нужно ли собирать цены в данный момент
function shouldCollectPrices() {
  const currentPeriod = getCurrentPricePeriod()
  const state = getPriceCollectionState()
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  const now = new Date()
  const hour = now.getHours()
  
  // Дополнительная защита: проверяем, что период соответствует текущему времени
  // Это предотвращает сбор для неправильного периода
  if (state.period && state.period !== currentPeriod && state.date === todayStr) {
    // Если период из состояния не соответствует текущему времени - игнорируем состояние
    console.log(`ShouldCollect: период из состояния (${state.period}) не соответствует текущему периоду (${currentPeriod}). Сбрасываем состояние.`)
    // Сбрасываем состояние, так как оно устарело
    setPriceCollectionState(currentPeriod, false)
  }
  
  // Если сменился период или день - начинаем новый сбор
  if (state.period !== currentPeriod || state.date !== todayStr) {
    return { should: true, reason: 'new_period_or_day', period: currentPeriod }
  }
  
  // Если текущий период уже завершен - не собираем
  if (state.completed && state.isToday) {
    return { should: false, reason: 'period_completed', period: currentPeriod }
  }
  
  // Продолжаем сбор текущего периода
  return { should: true, reason: 'continue_collection', period: currentPeriod }
}

// Получает цену из History для указанного предмета и периода
function getHistoryPriceForPeriod_(historySheet, itemName, period = null) {
  if (!historySheet || !itemName) return { found: false }
  
  if (!period) {
    period = getCurrentPricePeriod()
  }
  
  const lastRow = historySheet.getLastRow()
  const lastCol = historySheet.getLastColumn()
  if (lastRow <= 1 || lastCol < HISTORY_COLUMNS.FIRST_DATE_COL) {
    return { found: false, reason: 'no_data' }
  }
  
  // Находим строку предмета
  const names = historySheet.getRange(DATA_START_ROW, 2, lastRow - 1, 1).getValues()
  const target = String(itemName || '').trim()
  const idx = names.findIndex(r => String(r[0] || '').trim() === target)
  
  if (idx === -1) {
    return { found: false, reason: 'item_not_found' }
  }
  
  const row = idx + DATA_START_ROW
  
  // Получаем все колонки дат
  const firstDateCol = HISTORY_COLUMNS.FIRST_DATE_COL
  const dateCount = lastCol - firstDateCol + 1
  if (dateCount <= 0) {
    return { found: false, reason: 'no_dates' }
  }
  
  // Читаем заголовки дат и значения цен
  const dateHeaders = historySheet.getRange(HEADER_ROW, firstDateCol, 1, dateCount).getDisplayValues()[0]
  const priceValues = historySheet.getRange(row, firstDateCol, 1, dateCount).getValues()[0]
  
  const now = new Date()
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yy')
  const periodLabel = period === PRICE_COLLECTION_PERIODS.MORNING ? 'ночь' : 'день'
  const targetHeader = `${todayStr} ${periodLabel}`
  
  // Ищем цену за текущий период (сегодня + ночь/день)
  let currentPrice = null
  let currentColIndex = -1
  let currentDate = null
  
  for (let i = dateHeaders.length - 1; i >= 0; i--) {
    const header = String(dateHeaders[i] || '').trim()
    const price = priceValues[i]
    
    // Ищем точное совпадение с текущим периодом
    if (header === targetHeader && typeof price === 'number' && !isNaN(price) && price > 0) {
      currentPrice = price
      currentColIndex = firstDateCol + i
      currentDate = header
      break
    }
  }
  
  // Если нашли цену за текущий период - возвращаем её
  if (currentPrice !== null) {
    return {
      found: true,
      price: currentPrice,
      date: currentDate,
      colIndex: currentColIndex,
      isCurrentPeriod: true
    }
  }
  
  // Если цена за текущий период отсутствует - ищем предыдущую цену (любого периода)
  // Важно: ищем самую последнюю заполненную цену, пропуская только текущий период
  // Это позволяет найти цену за другой период за сегодня (например, "ночь" если ищем "день")
  let previousPrice = null
  let previousDate = null
  let previousColIndex = -1
  
  // Ищем самую правую (последнюю) заполненную цену, пропуская только текущий период
  // Логика: идём справа налево и берём первую заполненную цену, которая не является текущим периодом
  for (let i = dateHeaders.length - 1; i >= 0; i--) {
    const header = String(dateHeaders[i] || '').trim()
    if (!header) continue  // Пропускаем пустые заголовки
    
    // Пропускаем только текущий период (если он пустой, мы уже пропустили его выше)
    if (header === targetHeader) continue
    
    // Пропускаем только действительно будущие даты (не все периоды за сегодня)
    // Сравниваем только по дате, не учитывая период (ночь/день)
    // Формат заголовка: "dd.MM.yy период" или "dd.MM.yy"
    const headerParts = header.split(' ')
    if (headerParts.length === 0) continue
    const headerDate = headerParts[0]  // Берём только дату (до пробела)
    
    // Проверяем, что дата не будущая (строковое сравнение формата dd.MM.yy работает корректно)
    // Например: "31.10.25" < "01.11.25" корректно
    if (headerDate > todayStr) continue
    
    // Проверяем, что цена валидна (число, не NaN, больше 0)
    const price = priceValues[i]
    if (typeof price === 'number' && !isNaN(price) && price > 0) {
      previousPrice = price
      previousDate = header
      previousColIndex = firstDateCol + i
      break  // Нашли последнюю заполненную цену
    }
    // Если цена пустая или невалидна - продолжаем поиск дальше налево
  }
  
  if (previousPrice !== null) {
    return {
      found: true,
      price: previousPrice,
      date: previousDate,
      colIndex: previousColIndex,
      isCurrentPeriod: false,
      isOutdated: true
    }
  }
  
  return { found: false, reason: 'no_price_found' }
}

// === УНИФИЦИРОВАННАЯ СИСТЕМА ОБНОВЛЕНИЯ ЦЕН ===

// Создание колонки периода, обновление цен и синхронизация с Invest/Sales
function history_createPeriodAndUpdate() {
  // Используем период из состояния (устанавливается триггерами)
  // Если состояние не установлено, используем текущий период
  const state = getPriceCollectionState()
  const requestedPeriod = state.period || getCurrentPricePeriod()
  const currentPeriod = getCurrentPricePeriod()
  
  // Дополнительная проверка: период должен соответствовать текущему времени
  // Это предотвращает преждевременное создание колонок
  if (requestedPeriod !== currentPeriod) {
    console.log(`Unified: несоответствие периода. Запрошен: ${requestedPeriod}, текущий: ${currentPeriod}. Используем текущий период.`)
  }
  
  const period = currentPeriod  // Используем текущий период, а не запрошенный
  
  console.log(`Unified: начало обработки периода ${period}`)
  
  const lockKey = 'unified_price_update_lock'
  const lockCheck = acquireLock_(lockKey, 600)
  if (lockCheck.locked) {
    console.log(`Unified: пропуск - обновление уже выполняется`)
    return
  }
  
  let updateExecuted = false
  try {
    // Создаём колонку для текущего периода (с проверкой времени внутри функции)
    history_ensurePeriodColumn(period)
    
    // Обновляем цены в History для текущего периода
    const historyCompleted = history_updatePricesForPeriod(period)
    
    // Сбрасываем Invest/Sales и синхронизируем цены
    invest_dailyReset()
    sales_dailyReset()
    syncPricesFromHistoryToInvestAndSales()
    
    // Если сбор завершен - отмечаем период как завершенный и обновляем аналитику
    if (historyCompleted) {
      setPriceCollectionState(period, true)
      console.log(`Unified: сбор цен для периода ${period} завершен`)
      
      // Обновляем тренды, выделение min/max и новые столбцы в History
      try {
        history_updateAllAnalytics_()
        console.log(`Unified: тренды, выделение min/max и текущая цена/min/max обновлены`)
      } catch (e) {
        console.error('Unified: ошибка при обновлении трендов/выделения:', e)
        logAutoAction_('Unified', 'Ошибка обновления трендов', 'ERROR')
      }
      
      // Обновляем аналитику в Invest/Sales после завершения периода
      try {
        invest_syncMinMaxFromHistory()
        invest_syncTrendDaysFromHistory()
        invest_syncExtendedAnalyticsFromHistory()
        sales_syncMinMaxFromHistory()
        sales_syncTrendDaysFromHistory()
        sales_syncExtendedAnalyticsFromHistory()
        console.log(`Unified: аналитика в Invest/Sales обновлена`)
      } catch (e) {
        console.error('Unified: ошибка при обновлении аналитики в Invest/Sales:', e)
        logAutoAction_('Unified', 'Ошибка обновления аналитики', 'ERROR')
      }
      
      logAutoAction_('Unified', `Сбор цен завершен (${period})`, 'OK')
    } else {
      setPriceCollectionState(period, false)
    }
    
    updateExecuted = true
  } catch (e) {
    console.error('Unified: ошибка при обновлении:', e)
    logAutoAction_('Unified', 'Ошибка обновления', 'ERROR')
  } finally {
    if (updateExecuted) {
      releaseLock_(lockKey)
    }
  }
}

// Основная функция обновления цен (продолжение сбора)
function unified_priceUpdate() {
  const check = shouldCollectPrices()
  
  if (!check.should) {
    console.log(`Unified: сбор цен не требуется (${check.reason})`)
    return
  }
  
  const period = check.period
  const lockKey = 'unified_price_update_lock'
  const lockCheck = acquireLock_(lockKey, 600)
  if (lockCheck.locked) {
    console.log(`Unified: пропуск - обновление уже выполняется`)
    return
  }
  
  let updateExecuted = false
  try {
    const historyCompleted = history_updatePricesForPeriod(period)
    
    // Если historyCompleted === false, это может означать:
    // 1. Колонка еще не создана (нормально для unified_priceUpdate, который может сработать до создания колонки)
    // 2. Нет данных для обновления
    // В этом случае просто выходим без ошибки
    if (historyCompleted === false) {
      console.log(`Unified: обновление цен не выполнено (колонка еще не создана или нет данных)`)
      return
    }
    
    syncPricesFromHistoryToInvestAndSales()
    
    if (historyCompleted) {
      setPriceCollectionState(period, true)
      console.log(`Unified: сбор цен для периода ${period} завершен`)
      
      // Обновляем аналитику в History
      try {
        history_updateAllAnalytics_()
        console.log(`Unified: аналитика History обновлена`)
      } catch (e) {
        console.error('Unified: ошибка при обновлении трендов/выделения:', e)
        logAutoAction_('Unified', 'Ошибка обновления трендов', 'ERROR')
      }
      
      // Обновляем аналитику в Invest/Sales после завершения периода
      try {
        invest_syncMinMaxFromHistory()
        invest_syncTrendDaysFromHistory()
        invest_syncExtendedAnalyticsFromHistory()
        sales_syncMinMaxFromHistory()
        sales_syncTrendDaysFromHistory()
        sales_syncExtendedAnalyticsFromHistory()
        console.log(`Unified: аналитика в Invest/Sales обновлена`)
      } catch (e) {
        console.error('Unified: ошибка при обновлении аналитики в Invest/Sales:', e)
        logAutoAction_('Unified', 'Ошибка обновления аналитики', 'ERROR')
      }
      
      logAutoAction_('Unified', `Сбор цен завершен (${period})`, 'OK')
    } else {
      setPriceCollectionState(period, false)
    }
    
    updateExecuted = true
  } catch (e) {
    console.error('Unified: ошибка при обновлении:', e)
    logAutoAction_('Unified', 'Ошибка обновления', 'ERROR')
  } finally {
    if (updateExecuted) {
      releaseLock_(lockKey)
    }
  }
}

// Синхронизирует цены из History в Invest и Sales
function syncPricesFromHistoryToInvestAndSales() {
  const historySheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.HISTORY)
  if (!historySheet) {
    console.error('Sync: лист History не найден')
    return
  }
  
    // Синхронизируем Invest
    syncInvestPricesFromHistory_(historySheet)
    
    // Синхронизируем Sales
    syncSalesPricesFromHistory_(historySheet)
  
  // Обновляем аналитику в Invest/Sales после синхронизации цен
  try {
    invest_syncMinMaxFromHistory()
    invest_syncTrendDaysFromHistory()
    invest_syncExtendedAnalyticsFromHistory()
    sales_syncMinMaxFromHistory()
    sales_syncTrendDaysFromHistory()
    sales_syncExtendedAnalyticsFromHistory()
    console.log(`Sync: аналитика в Invest/Sales обновлена`)
  } catch (e) {
    console.error('Sync: ошибка при обновлении аналитики в Invest/Sales:', e)
  }
}

// Синхронизирует цены для Invest из History
function syncInvestPricesFromHistory_(historySheet) {
  const investSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.INVEST)
  if (!investSheet) return
  
  const lastRow = investSheet.getLastRow()
  if (lastRow <= 1) return
  
  const count = lastRow - 1
  
  const nameColIndex = getColumnIndex(INVEST_COLUMNS.NAME)
  const priceColIndex = getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE)
  // статус-колонка удалена
  
  const names = investSheet.getRange(DATA_START_ROW, nameColIndex, count, 1).getValues()
  const currentPrices = investSheet.getRange(DATA_START_ROW, priceColIndex, count, 1).getValues()
  // const statuses = investSheet.getRange(DATA_START_ROW, statusColIndex, count, 1).getValues()
  
  let updatedCount = 0
  let errorCount = 0
  
  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue
    
    const priceResult = getHistoryPriceForPeriod_(historySheet, name, getCurrentPricePeriod())
    
    if (!priceResult.found) {
      if (priceResult.reason === 'item_not_found') {
        currentPrices[i][0] = null
        errorCount++
      }
      continue
    }
    
    currentPrices[i][0] = priceResult.price
    // статусов больше нет
    updatedCount++
    
    const row = i + DATA_START_ROW
    if (priceResult.isOutdated) {
      investSheet.getRange(row, priceColIndex).setBackground(COLORS.STABLE)
    } else {
      investSheet.getRange(row, priceColIndex).setBackground(null)
    }
  }
  
  investSheet.getRange(DATA_START_ROW, priceColIndex, count, 1).setValues(currentPrices)
  // статусов больше нет
  
  invest_calculateBatch_(investSheet, currentPrices)
  
  console.log(`Sync Invest: обновлено ${updatedCount}, ошибок ${errorCount}`)
}

// Синхронизирует цены для Sales из History
function syncSalesPricesFromHistory_(historySheet) {
  const salesSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.SALES)
  if (!salesSheet) return
  
  const lastRow = salesSheet.getLastRow()
  if (lastRow <= 1) return
  
  const count = lastRow - 1
  
  const nameColIndex = getColumnIndex(SALES_COLUMNS.NAME)
  const priceColIndex = getColumnIndex(SALES_COLUMNS.CURRENT_PRICE)
  const dropColIndex = getColumnIndex(SALES_COLUMNS.PRICE_DROP)
  // статус-колонка удалена
  const sellColIndex = getColumnIndex(SALES_COLUMNS.SELL_PRICE)
  
  const names = salesSheet.getRange(DATA_START_ROW, nameColIndex, count, 1).getValues()
  const currentPrices = salesSheet.getRange(DATA_START_ROW, priceColIndex, count, 1).getValues()
  const priceDrops = salesSheet.getRange(DATA_START_ROW, dropColIndex, count, 1).getValues()
  // const statuses = salesSheet.getRange(DATA_START_ROW, statusColIndex, count, 1).getValues()
  const sellPrices = salesSheet.getRange(DATA_START_ROW, sellColIndex, count, 1).getValues()
  
  let updatedCount = 0
  let errorCount = 0
  
  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue
    
    const priceResult = getHistoryPriceForPeriod_(historySheet, name, getCurrentPricePeriod())
    
    if (!priceResult.found) {
      if (priceResult.reason === 'item_not_found') {
        currentPrices[i][0] = null
        errorCount++
      }
      continue
    }
    
    currentPrices[i][0] = priceResult.price
    // статусов больше нет
    
    const sellPrice = Number(sellPrices[i][0])
    if (Number.isFinite(sellPrice) && sellPrice > 0 && priceResult.price > 0) {
      priceDrops[i][0] = (sellPrice - priceResult.price) / sellPrice
    } else {
      priceDrops[i][0] = null
    }
    
    updatedCount++
    
    const row = i + DATA_START_ROW
    if (priceResult.isOutdated) {
      salesSheet.getRange(row, priceColIndex).setBackground(COLORS.STABLE)
    } else {
      salesSheet.getRange(row, priceColIndex).setBackground(null)
    }
  }
  
  salesSheet.getRange(DATA_START_ROW, priceColIndex, count, 1).setValues(currentPrices)
  salesSheet.getRange(DATA_START_ROW, dropColIndex, count, 1).setValues(priceDrops)
  // статусов больше нет
  
  console.log(`Sync Sales: обновлено ${updatedCount}, ошибок ${errorCount}`)
}


function fetchLowestPriceWithBackoff_(appid, itemName, options) {
  const name = String(itemName || '').trim()
  if (!name) return { ok: false, error: 'empty_name' }
  const {
    attempts = 1,  // Одна попытка на предмет
    baseDelayMs = 200,
    betweenItemsMs = 1000,  // 1 секунда между предметами
    timeBudgetMs = 330000,
    startedAt = Date.now(),
  } = options || {}

  // Проверка бюджета времени перед началом
  if (Date.now() - startedAt > timeBudgetMs - 5000) {
    return { ok: false, error: 'time_budget_low' }
  }

  let delay = baseDelayMs
  let lastResult = null
  
  for (let attempt = 1; attempt <= attempts; attempt++) {
    // Проверка бюджета времени перед каждой попыткой
    if (Date.now() - startedAt > timeBudgetMs - 1000) {
      return lastResult || { ok: false, error: 'time_budget_exceeded' }
    }
    
    const res = fetchLowestPrice_(appid, name)
    lastResult = res
    
    if (res.ok) {
      // Адаптивная пауза - оптимизированная стратегия
      const props = PropertiesService.getScriptProperties()
      const now = Date.now()
      const lastFetchTimeStr = props.getProperty('last_fetch_time')
      const lastFetchTime = lastFetchTimeStr ? Number(lastFetchTimeStr) : 0
      const timeSinceLastFetch = now - lastFetchTime
      
      // Базовый интервал: междуItemsMs (обычно 150мс)
      // Адаптивно увеличивается если предыдущий запрос был недавно
      let minInterval = options.betweenItemsMs || 150
      
      // Если прошло больше 200мс с последнего запроса, можно ускориться
      if (timeSinceLastFetch > 200) {
        minInterval = Math.max(100, minInterval - 50)
      }
      
      // Защита от слишком быстрых запросов
      if (timeSinceLastFetch < minInterval) {
        Utilities.sleep(minInterval - timeSinceLastFetch)
      }
      
      props.setProperty('last_fetch_time', String(now))
      
      return res
    }
    
    // Умные задержки между попытками для одного предмета
    if (attempt < attempts) {
      const errorDelay = Math.min(delay, 250) // 200-250мс между попытками
      Utilities.sleep(errorDelay)
      delay = Math.min(delay * 1.5, 250)
    }
  }
  
  return { ok: false, error: 'max_attempts' }
}

function shouldStopByTimeBudget_(startedAt, timeBudgetMs) {
  return Date.now() - startedAt > timeBudgetMs
}

/**
 * Универсальная функция для форматирования заголовков
 * @param {Range} headerRange - Диапазон заголовков
 */
function formatHeaderRange_(headerRange) {
  headerRange
    .setBackground(COLORS.BACKGROUND)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
}

function setImageAndLink_(sheet, row, appId, itemName, columns) {
  const name = String(itemName || '').trim()
  if (!name) return false
  try {
    const marketUrl = `https://steamcommunity.com/market/listings/${appId}/${encodeURIComponent(name)}`
    const imageUrl = `https://api.steamapis.com/image/item/${appId}/${encodeURIComponent(name)}`
    if (columns && columns.LINK) {
      sheet.getRange(`${columns.LINK}${row}`).setFormula(`=HYPERLINK("${marketUrl}";"Ссылка")`)
    }
    if (columns && columns.IMAGE) {
      sheet.getRange(`${columns.IMAGE}${row}`).setFormula(`=IMAGE("${imageUrl}")`)
    }
    return true
  } catch (e) {
    return false
  }
}

function buildImageAndLinkFormula_(appId, itemName) {
  const name = String(itemName || '').trim()
  if (!name) return { link: '', image: '' }
  const marketUrl = `https://steamcommunity.com/market/listings/${appId}/${encodeURIComponent(name)}`
  const imageUrl = `https://api.steamapis.com/image/item/${appId}/${encodeURIComponent(name)}`
  return {
    link: `=HYPERLINK("${marketUrl}";"Ссылка")`,
    image: `=IMAGE("${imageUrl}")`,
  }
}

function getHistoryMinMaxByName_(historySheet, itemName) {
  const lastRow = historySheet.getLastRow()
  const lastCol = historySheet.getLastColumn()
  if (lastRow <= 1 || lastCol <= 4) return { noItem: true }

  const names = historySheet.getRange(2, 2, lastRow - 1, 1).getValues().map(r => String(r[0] || '').trim())
  const target = String(itemName || '').trim()
  const idx = names.findIndex(v => v === target)
  if (idx === -1) return { noItem: true }
  const row = idx + 2

  const firstDateCol = getHistoryFirstDateCol_(historySheet) || 8 // по умолчанию H
  const span = lastCol - (firstDateCol - 1)
  if (span <= 0) return { noValues: true }
  const values = historySheet
    .getRange(row, firstDateCol, 1, span)
    .getValues()[0]
    .filter(v => typeof v === 'number' && !isNaN(v))

  if (!values.length) return { noValues: true }
  return { min: Math.min.apply(null, values), max: Math.max.apply(null, values) }
}

function getHistoryFirstDateCol_(historySheet) {
  const lastCol = historySheet.getLastColumn()
  if (lastCol < HISTORY_COLUMNS.FIRST_DATE_COL) return null
  
  // Ищем первую колонку, где заголовок выглядит как дата
  // Поддерживаем форматы: "dd.MM.yy", "dd.MM.yy ночь", "dd.MM.yy день"
  for (let col = HISTORY_COLUMNS.FIRST_DATE_COL; col <= lastCol; col++) {
    const cell = historySheet.getRange(HEADER_ROW, col, 1, 1)
    const dv = String(cell.getDisplayValue() || '').trim()
    const v = cell.getValue()
    
    if (v instanceof Date) return col
    if (/^\d{2}\.\d{2}\.\d{2}(\s+(ночь|день))?$/.test(dv)) return col
  }
  return null
}

function parseRuNumber_(input) {
  if (input == null) return { ok: false }
  const s = String(input).trim().replace(/\s+/g, '')
  if (!s) return { ok: false }
  // Запятая как десятичный разделитель допустима
  const normalized = s.replace(',', '.')
  const n = Number(normalized)
  return Number.isFinite(n) ? { ok: true, value: n } : { ok: false }
}

function findRowByName_(sheet, name, nameColIndex) {
  const lastRow = sheet.getLastRow()
  if (lastRow <= 1) return -1
  const col = nameColIndex || 2 // по умолчанию колонка B
  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues()
  const target = String(name || '').trim()
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === target) return i + 2 // смещение на заголовок
  }
  return -1
}

function getOrCreateAutoLogSheet_() {
  const ss = SpreadsheetApp.getActive()
  let sheet = ss.getSheetByName(SHEET_NAMES.AUTO_LOG)
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.AUTO_LOG)
    sheet.getRange(1, 1, 1, 4).setValues([
      ['Дата/время', 'Лист', 'Действие', 'Статус']
    ])
    sheet.setFrozenRows(HEADER_ROW)
    // Формат шапки
    formatHeaderRange_(sheet.getRange(1, 1, 1, 4))
    // Ширины
    sheet.setColumnWidth(1, 150) // Дата/время
    sheet.setColumnWidth(2, 100) // Лист
    sheet.setColumnWidth(3, 200) // Действие
    sheet.setColumnWidth(4, 100) // Статус
  }
  return sheet
}

function logAutoAction_(sheetName, action, status = 'OK') {
  const sheet = getOrCreateAutoLogSheet_()
  // Вставляем новую строку сразу после заголовка (строка 2)
  // Существующие строки автоматически сдвигаются вниз
  const insertRow = HEADER_ROW + 1
  sheet.insertRowAfter(HEADER_ROW)
  const now = new Date()
  sheet.getRange(insertRow, 1, 1, 4).setValues([[now, sheetName, action, status]])
  sheet.getRange(insertRow, 1).setNumberFormat('dd.MM.yyyy HH:mm')
  sheet.getRange(insertRow, 1, 1, 4).setVerticalAlignment('middle').setHorizontalAlignment('center')
}


function getOrCreateLogSheet_() {
  const ss = SpreadsheetApp.getActive()
  let sheet = ss.getSheetByName(SHEET_NAMES.LOG)
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.LOG)
    sheet.getRange(1, 1, 1, 8).setValues([
      ['Дата/время', 'Тип', 'Изображение', 'Предмет', 'Кол-во', 'Цена за шт', 'Сумма', 'Источник']
    ])
    sheet.setFrozenRows(HEADER_ROW)
    // Формат шапки
    formatHeaderRange_(sheet.getRange(1, 1, 1, 8))
    // Ширины и формат чисел
    sheet.setColumnWidth(1, 150)
    sheet.setColumnWidth(2, 90)
    sheet.setColumnWidth(3, 120) // Изображение (покрупнее)
    sheet.setColumnWidth(4, 250)
    sheet.setColumnWidth(5, 90)
    sheet.setColumnWidth(6, 120)
    sheet.setColumnWidth(7, 120)
    sheet.setColumnWidth(8, 120)
  }
  return sheet
}

function logOperation_(type, itemName, quantity, pricePerUnit, total, source) {
  const sheet = getOrCreateLogSheet_()
  // Вставляем новую строку сразу после заголовка (строка 2)
  // Существующие строки автоматически сдвигаются вниз
  const insertRow = HEADER_ROW + 1
  sheet.insertRowAfter(HEADER_ROW)
  const now = new Date()
  // Порядок колонок: A Дата, B Тип, C Изображение, D Предмет, E Кол-во, F Цена за шт, G Сумма, H Источник
  sheet.getRange(insertRow, 1, 1, 2).setValues([[now, type]])
  
  // Используем тот же механизм изображений, что и в других листах
  const logConfig = {
    IMAGE: 'C',
    NAME: 'D',
    LINK: 'G' // Используем колонку G для ссылки
  }
  setImageAndLink_(sheet, insertRow, 570, itemName, logConfig)
  
  sheet.getRange(insertRow, 4, 1, 4).setValues([[itemName, quantity, pricePerUnit, total]])
  sheet.getRange(insertRow, 8).setValue(source)
  
  // Форматы и выравнивания
  sheet.getRange(insertRow, 1).setNumberFormat('dd.MM.yyyy HH:mm')
  sheet.getRange(insertRow, 1, 1, 8).setVerticalAlignment('middle').setHorizontalAlignment('center')
  sheet.getRange(insertRow, 4).setHorizontalAlignment('left')
  sheet.getRange(insertRow, 5).setNumberFormat('0')
  sheet.getRange(insertRow, 6).setNumberFormat('#,##0.00 ₽')
  sheet.getRange(insertRow, 7).setNumberFormat('#,##0.00 ₽')
  
  // Устанавливаем высоту строки для изображения
  sheet.setRowHeight(insertRow, 85)
}

function highlightDuplicatesByName_(sheet, nameColIndex, color) {
  const lastRow = sheet.getLastRow()
  if (lastRow <= 1) return { duplicates: 0, rows: [] }
  const col = nameColIndex || 2
  const names = sheet.getRange(2, col, lastRow - 1, 1).getValues().map(r => String(r[0] || '').trim())
  const seen = new Map()
  const dupRows = []
  for (let i = 0; i < names.length; i++) {
    const n = names[i]
    if (!n) continue
    if (seen.has(n)) {
      dupRows.push(i + 2)
    } else {
      seen.set(n, i + 2)
    }
  }
  // Сброс фона на всей колонке названий
  sheet.getRange(2, col, lastRow - 1, 1).setBackground(null)
  if (dupRows.length) {
    const bg = color || '#e3f2fd'
    dupRows.forEach(r => sheet.getRange(r, col).setBackground(bg))
  }
  return { duplicates: dupRows.length, rows: dupRows }
}

// === УНИВЕРСАЛЬНАЯ ФУНКЦИЯ ОБНОВЛЕНИЯ ИЗОБРАЖЕНИЙ И ССЫЛОК ===
/**
 * Универсальная функция для обновления изображений и ссылок в любом листе
 * @param {Object} config - Конфигурация модуля (INVEST_CONFIG, SALES_CONFIG и т.д.)
 * @param {Sheet} sheet - Лист для обновления
 * @param {boolean} updateAll - Обновлять все (true) или только отсутствующие (false)
 * @param {string} moduleName - Название модуля для логирования
 * @returns {Object} {updatedCount, errorCount}
 */
function updateImagesAndLinksUniversal_(config, sheet, updateAll, moduleName) {
  const lastRow = sheet.getLastRow()
  let updatedCount = 0
  let errorCount = 0

  if (lastRow <= 1) return { updatedCount: 0, errorCount: 0 }

  const count = lastRow - 1
  const names = sheet.getRange(DATA_START_ROW, 2, count, 1).getValues() // B
  const imageCol = sheet.getRange(DATA_START_ROW, 1, count, 1) // A
  const linkCol = config.COLUMNS.LINK ? sheet.getRange(DATA_START_ROW, getColumnIndex(config.COLUMNS.LINK), count, 1) : null
  
  const imageFormulas = imageCol.getFormulas()
  const linkFormulas = linkCol ? linkCol.getFormulas() : null

  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue

    const hasImage = imageFormulas[i][0]
    const hasLink = linkFormulas && linkFormulas[i][0]
    const needsUpdate = updateAll || !hasImage || !hasLink

    if (!needsUpdate) continue

    try {
      const built = buildImageAndLinkFormula_(config.STEAM_APPID, name)
      imageFormulas[i][0] = built.image
      if (linkFormulas) {
        linkFormulas[i][0] = built.link
      }
      updatedCount++
      Utilities.sleep(100)
    } catch (e) {
      console.error(`${moduleName}: ошибка при обновлении изображения ${name}:`, e)
      errorCount++
    }
  }

  // Применяем изменения batch
  imageCol.setFormulas(imageFormulas)
  if (linkCol && linkFormulas) {
    linkCol.setFormulas(linkFormulas)
  }

  return { updatedCount, errorCount }
}

// Вспомогательная функция для получения индекса колонки
function getColumnIndex(columnLetter) {
  let result = 0
  for (let i = 0; i < columnLetter.length; i++) {
    result = result * 26 + (columnLetter.charCodeAt(i) - 'A'.charCodeAt(0) + 1)
  }
  return result
}

// === УНИВЕРСАЛЬНАЯ ФУНКЦИЯ СИНХРОНИЗАЦИИ MIN/MAX ИЗ HISTORY ===
/**
 * Универсальная функция для синхронизации Min/Max цен из History
 * @param {Sheet} targetSheet - Лист назначения (Invest или Sales)
 * @param {number} minColIndex - Индекс колонки Min
 * @param {number} maxColIndex - Индекс колонки Max
 * @param {boolean} updateAll - Обновлять все или только пустые
 * @returns {Object} {updatedCount}
 */
function syncMinMaxFromHistoryUniversal_(targetSheet, minColIndex, maxColIndex, updateAll) {
  const lastRow = targetSheet.getLastRow()
  if (lastRow <= 1) return { updatedCount: 0 }

  const history = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.HISTORY)
  if (!history) return { updatedCount: 0 }

  const count = lastRow - 1
  const names = targetSheet.getRange(DATA_START_ROW, 2, count, 1).getValues()
  const minCol = targetSheet.getRange(DATA_START_ROW, minColIndex, count, 1).getValues()
  const maxCol = targetSheet.getRange(DATA_START_ROW, maxColIndex, count, 1).getValues()

  const outMin = minCol.map(r => [r[0]])
  const outMax = maxCol.map(r => [r[0]])
  let updatedCount = 0

  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue

    if (!updateAll) {
      const hasMin = outMin[i][0] != null && outMin[i][0] !== ''
      const hasMax = outMax[i][0] != null && outMax[i][0] !== ''
      if (hasMin && hasMax) continue
    }

    const mm = getHistoryMinMaxByName_(history, name)
    if (mm && mm.min != null && mm.max != null) {
      outMin[i][0] = mm.min
      outMax[i][0] = mm.max
      updatedCount++
    } else if (mm && mm.noItem) {
      outMin[i][0] = STATUS.MISSING
      outMax[i][0] = STATUS.MISSING
      updatedCount++
    } else if (mm && mm.noValues) {
      outMin[i][0] = STATUS.NO_VALUES
      outMax[i][0] = STATUS.NO_VALUES
      updatedCount++
    }
  }

  targetSheet.getRange(DATA_START_ROW, minColIndex, count, 1).setValues(outMin)
  targetSheet.getRange(DATA_START_ROW, maxColIndex, count, 1).setValues(outMax)

  return { updatedCount }
}

// === УНИВЕРСАЛЬНАЯ ФУНКЦИЯ СИНХРОНИЗАЦИИ ТРЕНД/ДНЕЙ СМЕНЫ ИЗ HISTORY ===
/**
 * Универсальная функция для синхронизации Тренд/Дней смены из History
 * @param {Sheet} targetSheet - Лист назначения (Invest или Sales)
 * @param {number} trendColIndex - Индекс колонки Тренд
 * @param {number} daysColIndex - Индекс колонки Дней смены
 * @param {boolean} updateAll - Обновлять все или только пустые
 * @returns {Object} {updatedCount}
 */
function syncTrendDaysFromHistoryUniversal_(targetSheet, trendColIndex, daysColIndex, updateAll) {
  const lastRow = targetSheet.getLastRow()
  if (lastRow <= 1) return { updatedCount: 0 }

  const history = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.HISTORY)
  if (!history) return { updatedCount: 0 }

  const count = lastRow - 1
  const names = targetSheet.getRange(DATA_START_ROW, 2, count, 1).getValues()
  const trendCol = targetSheet.getRange(DATA_START_ROW, trendColIndex, count, 1).getValues()
  const daysCol = targetSheet.getRange(DATA_START_ROW, daysColIndex, count, 1).getValues()

  const outTrend = trendCol.map(r => [r[0]])
  const outDays = daysCol.map(r => [r[0]])
  let updatedCount = 0

  // Получаем данные из History
  const historyLastRow = history.getLastRow()
  if (historyLastRow <= 1) {
    // Если History пустой, заполняем значениями по умолчанию
    for (let i = 0; i < count; i++) {
      const name = String(names[i][0] || '').trim()
      if (!name) continue
      
      if (updateAll || !outTrend[i][0]) {
        outTrend[i][0] = '🟪'
        outDays[i][0] = 0
        updatedCount++
      }
    }
  } else {
    // Читаем все данные из History одним батчем
    const historyNames = history.getRange(DATA_START_ROW, 2, historyLastRow - 1, 1).getValues()
    const historyTrendCol = getColumnIndex(HISTORY_COLUMNS.TREND)
    const historyDaysCol = getColumnIndex(HISTORY_COLUMNS.DAYS_CHANGE)
    const historyTrends = history.getRange(DATA_START_ROW, historyTrendCol, historyLastRow - 1, 1).getValues()
    const historyDays = history.getRange(DATA_START_ROW, historyDaysCol, historyLastRow - 1, 1).getValues()

    // Создаем мапу для быстрого поиска
    const historyMap = new Map()
    for (let i = 0; i < historyNames.length; i++) {
      const hName = String(historyNames[i][0] || '').trim()
      if (hName) {
        historyMap.set(hName, {
          trend: historyTrends[i][0] || '🟪',
          days: historyDays[i][0] || 0
        })
      }
    }

    // Синхронизируем данные
    for (let i = 0; i < count; i++) {
      const name = String(names[i][0] || '').trim()
      if (!name) continue

      if (!updateAll) {
        const hasTrend = outTrend[i][0] != null && outTrend[i][0] !== ''
        const hasDays = outDays[i][0] != null && outDays[i][0] !== ''
        if (hasTrend && hasDays) continue
      }

      const historyData = historyMap.get(name)
      if (historyData) {
        outTrend[i][0] = historyData.trend
        outDays[i][0] = historyData.days
        updatedCount++
      } else {
        // Предмет не найден в History - ставим значения по умолчанию
        outTrend[i][0] = '🟪'
        outDays[i][0] = 0
        updatedCount++
      }
    }
  }

  targetSheet.getRange(DATA_START_ROW, trendColIndex, count, 1).setValues(outTrend)
  targetSheet.getRange(DATA_START_ROW, daysColIndex, count, 1).setValues(outDays)

  return { updatedCount }
}

// === УНИВЕРСАЛЬНАЯ ФУНКЦИЯ СИНХРОНИЗАЦИИ РАСШИРЕННОЙ АНАЛИТИКИ ИЗ HISTORY ===
/**
 * Универсальная функция для синхронизации Фаза/Потенциал/Рекомендация из History
 * @param {Sheet} targetSheet - Лист назначения (Invest или Sales)
 * @param {number} phaseColIndex - Индекс колонки Фаза
 * @param {number} potentialColIndex - Индекс колонки Потенциал
 * @param {number} recommendationColIndex - Индекс колонки Рекомендация
 * @param {boolean} updateAll - Обновлять все или только пустые
 * @returns {Object} {updatedCount}
 */
function syncExtendedAnalyticsFromHistoryUniversal_(targetSheet, phaseColIndex, potentialColIndex, recommendationColIndex, updateAll) {
  const lastRow = targetSheet.getLastRow()
  if (lastRow <= 1) return { updatedCount: 0 }

  const history = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.HISTORY)
  if (!history) return { updatedCount: 0 }

  const count = lastRow - 1
  const names = targetSheet.getRange(DATA_START_ROW, 2, count, 1).getValues()
  const phaseCol = targetSheet.getRange(DATA_START_ROW, phaseColIndex, count, 1).getValues()
  const potentialCol = targetSheet.getRange(DATA_START_ROW, potentialColIndex, count, 1).getValues()
  const recommendationCol = targetSheet.getRange(DATA_START_ROW, recommendationColIndex, count, 1).getValues()

  const outPhase = phaseCol.map(r => [r[0]])
  const outPotential = potentialCol.map(r => [r[0]])
  const outRecommendation = recommendationCol.map(r => [r[0]])
  let updatedCount = 0

  // Получаем данные из History
  const historyLastRow = history.getLastRow()
  
  if (historyLastRow <= 1) {
    // Если History пустой, заполняем значениями по умолчанию
    for (let i = 0; i < count; i++) {
      const name = String(names[i][0] || '').trim()
      if (!name) continue
      
      if (updateAll || !outPhase[i][0]) {
        outPhase[i][0] = '❓'
        outPotential[i][0] = null  // null вместо '—' для числового формата
        outRecommendation[i][0] = '👀 НАБЛЮДАТЬ'
        updatedCount++
      }
    }
  } else {
    // Читаем данные из History одним батчем
    const historyNames = history.getRange(DATA_START_ROW, 2, historyLastRow - 1, 1).getValues()
    
    // Колонки K/L/M (11/12/13) - Фаза, Потенциал, Рекомендация
    const phaseColHistory = getColumnIndex(HISTORY_COLUMNS.PHASE)
    const potentialColHistory = getColumnIndex(HISTORY_COLUMNS.POTENTIAL)
    const recommendationColHistory = getColumnIndex(HISTORY_COLUMNS.RECOMMENDATION)
    
    const historyPhase = history.getRange(DATA_START_ROW, phaseColHistory, historyLastRow - 1, 1).getValues()
    const historyPotential = history.getRange(DATA_START_ROW, potentialColHistory, historyLastRow - 1, 1).getValues()
    const historyRecommendation = history.getRange(DATA_START_ROW, recommendationColHistory, historyLastRow - 1, 1).getValues()

    // Создаем мапу для быстрого поиска
    const historyMap = new Map()
    for (let i = 0; i < historyNames.length; i++) {
      const hName = String(historyNames[i][0] || '').trim()
      if (hName) {
        // Потенциал теперь хранится как число (процент в виде десятичной дроби, например 0.14 для +14%)
        // Если значение null или строка '—', используем null
        let potentialValue = historyPotential[i][0]
        if (potentialValue === '—' || potentialValue === null || potentialValue === '') {
          potentialValue = null
        }
        historyMap.set(hName, {
          phase: historyPhase[i][0] || '❓',
          potential: potentialValue,
          recommendation: historyRecommendation[i][0] || '👀 НАБЛЮДАТЬ'
        })
      }
    }

    // Синхронизируем данные
    for (let i = 0; i < count; i++) {
      const name = String(names[i][0] || '').trim()
      if (!name) continue

      if (!updateAll) {
        const hasPhase = outPhase[i][0] != null && outPhase[i][0] !== ''
        const hasPotential = outPotential[i][0] != null && outPotential[i][0] !== ''
        const hasRecommendation = outRecommendation[i][0] != null && outRecommendation[i][0] !== ''
        if (hasPhase && hasPotential && hasRecommendation) continue
      }

      const historyData = historyMap.get(name)
      if (historyData) {
        outPhase[i][0] = historyData.phase
        outPotential[i][0] = historyData.potential
        outRecommendation[i][0] = historyData.recommendation
        updatedCount++
      } else {
        // Предмет не найден в History - ставим значения по умолчанию
        outPhase[i][0] = '❓'
        outPotential[i][0] = null  // null вместо '—' для числового формата
        outRecommendation[i][0] = '👀 НАБЛЮДАТЬ'
        updatedCount++
      }
    }
  }

  targetSheet.getRange(DATA_START_ROW, phaseColIndex, count, 1).setValues(outPhase)
  targetSheet.getRange(DATA_START_ROW, potentialColIndex, count, 1).setValues(outPotential)
  targetSheet.getRange(DATA_START_ROW, recommendationColIndex, count, 1).setValues(outRecommendation)

  return { updatedCount }
}

/**
 * Универсальная функция для обновления всей аналитики из History через меню
 * @param {string} moduleName - Название модуля (Invest/Sales)
 * @param {Function} syncMinMax - Функция синхронизации Min/Max
 * @param {Function} syncTrendDays - Функция синхронизации Тренд/Дней смены
 * @param {Function} syncExtended - Функция синхронизации расширенной аналитики
 */
function updateAllAnalyticsManual_(moduleName, syncMinMax, syncTrendDays, syncExtended) {
  const ui = SpreadsheetApp.getUi()
  const response = ui.alert(
    `Обновить всю аналитику (${moduleName})`,
    'Обновить Min/Max, Тренд/Дней смены, Фазу/Потенциал/Рекомендацию?\n\n' +
      'Да - Для всех строк\n' +
      'Нет - Только для пустых значений',
    ui.ButtonSet.YES_NO_CANCEL
  )
  if (response === ui.Button.CANCEL) return
  const modeAll = response === ui.Button.YES
  
  syncMinMax(modeAll)
  syncTrendDays(modeAll)
  syncExtended(modeAll)
  
  ui.alert(`Вся аналитика обновлена из History (${moduleName})`)
}

/**
 * Универсальная функция для обновления изображений и ссылок через меню
 * @param {Object} config - Конфигурация модуля (INVEST_CONFIG, SALES_CONFIG и т.д.)
 * @param {Function} getSheetFunction - Функция получения листа (getInvestSheet_, getSalesSheet_ и т.д.)
 * @param {string} moduleName - Название модуля для отображения
 */
function updateImagesAndLinksMenu_(config, getSheetFunction, moduleName) {
  const ui = SpreadsheetApp.getUi()
  const response = ui.alert(
    `Обновить изображения и ссылки (${moduleName})`,
    'Как вы хотите обновить изображения и ссылки?\n\n' +
      'Да - Обновить все изображения и ссылки\n' +
      'Нет - Обновить только отсутствующие\n' +
      'Отмена - Отменить обновление',
    ui.ButtonSet.YES_NO_CANCEL
  )
  if (response === ui.Button.CANCEL) return

  const sheet = getSheetFunction()
  if (!sheet) return

  const updateAll = response === ui.Button.YES
  const result = updateImagesAndLinksUniversal_(config, sheet, updateAll, moduleName)

  try {
    ui.alert(
      `${moduleName} — результат обновления`,
      `Обновлено строк: ${result.updatedCount}\nОшибок: ${result.errorCount}`,
      ui.ButtonSet.OK
    )
  } catch (e) {
    console.log(`${moduleName}: невозможно показать UI в данном контексте`)
  }
}

