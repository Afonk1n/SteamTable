// Общие утилиты для всех модулей

/**
 * Универсальная функция для выполнения HTTP запросов с retry логикой
 * @param {string} url - URL для запроса
 * @param {Object} options - Опции для UrlFetchApp.fetch
 * @param {Object} config - Конфигурация retry (maxRetries, retryDelayMs, apiName)
 * @returns {Object} {ok: boolean, response?: Object, error?: string, details?: string}
 */
function fetchWithRetry_(url, options = {}, config = {}) {
  const maxRetries = config.maxRetries || API_CONFIG.STEAM_WEB_API.MAX_RETRIES
  const retryDelayMs = config.retryDelayMs || API_CONFIG.STEAM_WEB_API.RETRY_DELAY_MS
  const apiName = config.apiName || 'API'
  
  let attempts = 0
  
  while (attempts < maxRetries) {
    try {
      const response = UrlFetchApp.fetch(url, options)
      const responseCode = response.getResponseCode()
      
      if (responseCode === HTTP_STATUS.RATE_LIMIT) {
        const delay = retryDelayMs * Math.pow(2, attempts)
        console.warn(`${apiName}: лимит запросов, повтор через ${delay}ms (попытка ${attempts + 1}/${maxRetries})`)
        Utilities.sleep(delay)
        attempts++
        continue
      }
      
      if (responseCode !== HTTP_STATUS.OK) {
        const responseText = response.getContentText()
        const errorPreview = responseText.substring(0, LIMITS.ERROR_MESSAGE_MAX_LENGTH)
        console.error(`${apiName}: HTTP ошибка ${responseCode}: ${errorPreview}`)
        return { ok: false, error: `http_${responseCode}`, details: errorPreview }
      }
      
      return { ok: true, response: response }
      
    } catch (e) {
      attempts++
      if (attempts >= maxRetries) {
        console.error(`${apiName}: исключение после всех попыток:`, e)
        return { ok: false, error: 'exception', details: e.message }
      }
      
      const delay = retryDelayMs * Math.pow(2, attempts - 1)
      console.warn(`${apiName}: ошибка, повтор через ${delay}ms (попытка ${attempts}/${maxRetries}):`, e.message)
      Utilities.sleep(delay)
    }
  }
  
  return { ok: false, error: 'max_retries_exceeded' }
}

function parseLocalizedPrice_(raw) {
  if (raw == null) return { ok: false, error: 'no_price' }
  const normalized = String(raw).replace(/[^0-9.,]+/g, '').replace(',', '.')
  const price = parseFloat(normalized)
  return Number.isFinite(price) ? { ok: true, price } : { ok: false, error: 'nan' }
}

/**
 * Валидация цены перед записью
 * @param {number} price - Цена для проверки
 * @param {string} itemName - Название предмета (для логирования)
 * @returns {Object} {valid: boolean, price?: number, error?: string}
 */
function validatePrice_(price, itemName = '') {
  // Проверка на тип данных
  if (typeof price !== 'number') {
    console.warn(`Price validation: некорректный тип для "${itemName}": ${typeof price}, значение: ${price}`)
    return { valid: false, error: 'invalid_type' }
  }
  
  // Проверка на NaN или Infinity
  if (!Number.isFinite(price)) {
    console.warn(`Price validation: некорректное значение (NaN/Infinity) для "${itemName}": ${price}`)
    return { valid: false, error: 'not_finite' }
  }
  
  // Проверка на отрицательные значения
  if (price < 0) {
    console.warn(`Price validation: отрицательная цена для "${itemName}": ${price}`)
    return { valid: false, error: 'negative_price' }
  }
  
  // Проверка на ноль (может быть допустимо, но подозрительно)
  if (price === 0) {
    console.warn(`Price validation: нулевая цена для "${itemName}"`)
    return { valid: false, error: 'zero_price' }
  }
  
  // Проверка на слишком большие значения (аномальные)
  if (price > MAX_REASONABLE_PRICE) {
    console.warn(`Price validation: аномально высокая цена для "${itemName}": ${price} (максимум: ${MAX_REASONABLE_PRICE})`)
    return { valid: false, error: 'price_too_high', price, maxPrice: MAX_REASONABLE_PRICE }
  }
  
  return { valid: true, price }
}

/**
 * Получает минимальную цену предмета с приоритетом SteamWebAPI.ru и fallback на priceoverview
 * @param {number} appid - ID приложения Steam (570 для Dota 2)
 * @param {string} itemName - Название предмета
 * @returns {Object} {ok: boolean, price?: number, error?: string, source?: string}
 */
function fetchLowestPrice_(appid, itemName) {
  const name = String(itemName || '').trim()
  if (!name) return { ok: false, error: 'empty_name' }
  
  // Пробуем сначала SteamWebAPI.ru (более надежный источник)
  try {
    const itemData = steamWebAPI_getItemData(name, 'dota2')
    if (itemData && itemData.ok && itemData.price && itemData.price > 0) {
      return {
        ok: true,
        price: itemData.price,
        source: 'steamwebapi'
      }
    }
  } catch (e) {
    console.warn(`SteamWebAPI failed for "${name}", falling back to priceoverview:`, e.message)
  }
  
  // Fallback на старый метод (priceoverview)
  return fetchLowestPrice_legacy_(appid, name)
}

/**
 * Старый метод получения цены через priceoverview (fallback)
 * @param {number} appid - ID приложения Steam
 * @param {string} itemName - Название предмета
 * @returns {Object} {ok: boolean, price?: number, error?: string, source?: string}
 */
function fetchLowestPrice_legacy_(appid, itemName) {
  const name = String(itemName || '').trim()
  if (!name) return { ok: false, error: 'empty_name' }
  
  const url = `https://steamcommunity.com/market/priceoverview/?currency=5&appid=${appid}&market_hash_name=${encodeURIComponent(name)}`
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true })
    if (response.getResponseCode() !== HTTP_STATUS.OK) {
      return { ok: false, error: 'http_' + response.getResponseCode(), source: 'priceoverview' }
    }
    const data = JSON.parse(response.getContentText())
    if (data && data.success && data.lowest_price) {
      const parsed = parseLocalizedPrice_(data.lowest_price)
      return parsed.ok 
        ? { ok: true, price: parsed.price, source: 'priceoverview' } 
        : { ok: false, error: parsed.error, source: 'priceoverview' }
    }
    return { ok: false, error: 'no_data', source: 'priceoverview' }
  } catch (e) {
    return { ok: false, error: 'exception', source: 'priceoverview', details: e.message }
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
  
  // ВАЖНО: Создаем лист AutoLog заранее, чтобы он точно существовал для логирования
  getOrCreateAutoLogSheet_()
  
  const lockKey = 'unified_price_update_lock'
  const lockCheck = acquireLock_(lockKey, LIMITS.LOCK_TIMEOUT_SEC)
  if (lockCheck.locked) {
    console.log(`Unified: пропуск - обновление уже выполняется`)
    return
  }
  
  let updateExecuted = false
  try {
    // Сохраняем lastCol до создания колонки, чтобы определить, была ли создана новая
    const historySheet = getHistorySheet_()
    const lastColBefore = historySheet ? historySheet.getLastColumn() : 0
    
    // Создаём колонку для текущего периода (с проверкой времени внутри функции)
    history_ensurePeriodColumn(period)
    
    // Логируем начало обновления цен и обновляем текущие цены (окраска желтым) только если была создана новая колонка
    const lastColAfter = historySheet ? historySheet.getLastColumn() : 0
    if (lastColAfter > lastColBefore) {
      const periodLabel = period === PRICE_COLLECTION_PERIODS.MORNING ? 'ночь' : 'день'
      logAutoAction_('History', `Начало обновления цен (${periodLabel})`, 'OK')
      
      // Обновляем текущие цены сразу после создания новой колонки, чтобы устаревшие цены окрасились в желтый
      try {
        history_updateCurrentPriceMinMax_(historySheet)
        console.log(`Unified: текущие цены обновлены и окрашены после создания новой колонки`)
      } catch (e) {
        console.error('Unified: ошибка при обновлении текущих цен:', e)
      }
    }
    
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
      
      // Сохраняем историю портфеля и обновляем Investment Scores (только после дневного сбора)
      // Примечание: аналитика Invest/Sales уже обновлена в syncPricesFromHistoryToInvestAndSales()
      try {
        if (period === PRICE_COLLECTION_PERIODS.EVENING) {
          portfolioStats_saveHistory_()
          console.log(`Unified: история портфеля сохранена`)
          
          // Обновляем Investment Scores после завершения дневного сбора
          try {
            history_updateInvestmentScores()
            invest_updateInvestmentScores()
            sales_updateBuybackScores()
            console.log(`Unified: Investment/Buyback Scores обновлены после дневного сбора`)
          } catch (e) {
            console.error('Unified: ошибка при обновлении Investment Scores:', e)
            logAutoAction_('Unified', 'Ошибка обновления Investment Scores', 'ERROR')
          }
        }
      } catch (e) {
        console.error('Unified: ошибка при сохранении истории портфеля:', e)
        logAutoAction_('Unified', 'Ошибка сохранения истории портфеля', 'ERROR')
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
  
  // ВАЖНО: Создаем лист AutoLog заранее, чтобы он точно существовал для логирования
  getOrCreateAutoLogSheet_()
  
  const lockKey = 'unified_price_update_lock'
  const lockCheck = acquireLock_(lockKey, LIMITS.LOCK_TIMEOUT_SEC)
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
      updateExecuted = true // Отмечаем, что обработка завершена, чтобы освободить блокировку
      return
    }
    
    syncPricesFromHistoryToInvestAndSales()
    
    // Убрали вызов telegram_checkPriceTargets() - теперь используется ежедневный триггер в 13:00
    
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
      
      // Проверяем резкие изменения цены (критические уведомления)
      try {
        telegram_checkPriceChanges_()
      } catch (e) {
        console.error('Unified: ошибка при проверке резких изменений цены:', e)
        // Не прерываем выполнение, просто логируем ошибку
      }
      
      // Сохраняем историю портфеля и обновляем Investment Scores (только после дневного сбора)
      // Примечание: аналитика Invest/Sales уже обновлена в syncPricesFromHistoryToInvestAndSales()
      try {
        if (period === PRICE_COLLECTION_PERIODS.EVENING) {
          portfolioStats_saveHistory_()
          console.log(`Unified: история портфеля сохранена`)
          
          // Обновляем Investment Scores после завершения дневного сбора
          try {
            history_updateInvestmentScores()
            invest_updateInvestmentScores()
            sales_updateBuybackScores()
            console.log(`Unified: Investment/Buyback Scores обновлены после дневного сбора`)
          } catch (e) {
            console.error('Unified: ошибка при обновлении Investment Scores:', e)
            logAutoAction_('Unified', 'Ошибка обновления Investment Scores', 'ERROR')
          }
        }
      } catch (e) {
        console.error('Unified: ошибка при сохранении истории портфеля:', e)
        logAutoAction_('Unified', 'Ошибка сохранения истории портфеля', 'ERROR')
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

/**
 * Читает данные из History для синхронизации цен
 * @param {Sheet} historySheet - Лист History
 * @returns {Object|null} Объект с данными History или null если данных нет
 */
function readHistoryDataForSync_(historySheet) {
  if (!historySheet) return null
  
  const historyLastRow = historySheet.getLastRow()
  const historyLastCol = historySheet.getLastColumn()
  const historyFirstDateCol = HISTORY_COLUMNS.FIRST_DATE_COL
  
  if (historyLastRow <= 1 || historyLastCol < historyFirstDateCol) {
    return null
  }
  
  const historyNames = historySheet.getRange(DATA_START_ROW, 2, historyLastRow - 1, 1).getValues()
  const historyDateCount = historyLastCol - historyFirstDateCol + 1
  const historyDateHeaders = historySheet.getRange(HEADER_ROW, historyFirstDateCol, 1, historyDateCount).getDisplayValues()[0]
  const historyPriceData = historySheet.getRange(DATA_START_ROW, historyFirstDateCol, historyLastRow - 1, historyDateCount).getValues()
  
  // Находим индекс колонки с текущим периодом
  const period = getCurrentPricePeriod()
  const now = new Date()
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yy')
  const periodLabel = period === PRICE_COLLECTION_PERIODS.MORNING ? 'ночь' : 'день'
  const targetHeader = `${todayStr} ${periodLabel}`
  
  let currentPeriodColIndex = -1
  for (let j = historyDateHeaders.length - 1; j >= 0; j--) {
    if (String(historyDateHeaders[j] || '').trim() === targetHeader) {
      currentPeriodColIndex = j
      break
    }
  }
  
  // Создаем Map для быстрого поиска цен по имени
  const historyPriceMap = new Map()
  for (let h = 0; h < historyLastRow - 1; h++) {
    const historyName = String(historyNames[h][0] || '').trim()
    if (historyName && historyPriceData[h]) {
      historyPriceMap.set(historyName, {
        row: h,
        prices: historyPriceData[h],
        currentPeriodColIndex: currentPeriodColIndex
      })
    }
  }
  
  return {
    historyPriceMap,
    currentPeriodColIndex,
    lastRow: historyLastRow
  }
}

/**
 * Определяет, устарела ли цена (не за текущий период)
 * @param {Object} historyData - Данные из historyPriceMap
 * @returns {boolean} true если цена устарела
 */
function isPriceOutdated_(historyData) {
  return historyData.currentPeriodColIndex < 0 || !historyData.prices[historyData.currentPeriodColIndex]
}

/**
 * Получает текущую цену из данных History
 * @param {Object} historyData - Данные из historyPriceMap
 * @returns {number|null} Цена или null если не найдена
 */
function getCurrentPriceFromHistoryData_(historyData) {
  let price = null
  
  // Сначала пытаемся получить цену за текущий период
  if (historyData.currentPeriodColIndex >= 0 && historyData.prices[historyData.currentPeriodColIndex]) {
    const value = historyData.prices[historyData.currentPeriodColIndex]
    if (typeof value === 'number' && !isNaN(value) && value > 0) {
      price = value
    }
  }
  
  // Если цена за текущий период отсутствует, ищем последнюю заполненную цену
  if (price === null) {
    for (let j = historyData.prices.length - 1; j >= 0; j--) {
      const value = historyData.prices[j]
      if (typeof value === 'number' && !isNaN(value) && value > 0) {
        price = value
        break
      }
    }
  }
  
  return price
}

/**
 * Универсальное обновление всех метрик для Invest и Sales
 * Вызывает расчет метрик (Liquidity, Demand, Momentum, Sales Trend, Volatility, Hero Trend)
 * для всех позиций в Invest и Sales
 */
function updateAllMetricsForInvestSales() {
  try {
    console.log('Utils: начало обновления метрик для Invest/Sales')
    
    // Обновляем метрики для Invest
    invest_calculateAllMetrics()
    console.log('Utils: метрики Invest обновлены')
    
    // Обновляем метрики для Sales
    sales_calculateAllMetrics()
    console.log('Utils: метрики Sales обновлены')
    
    console.log('Utils: обновление метрик завершено')
  } catch (e) {
    console.error('Utils: ошибка при обновлении метрик Invest/Sales:', e)
  }
}

// Синхронизирует цены из History в Invest и Sales
function syncPricesFromHistoryToInvestAndSales() {
  const historySheet = getHistorySheet_()
  if (!historySheet) {
    console.error('Sync: лист History не найден')
    return
  }
  
  // Читаем данные History один раз для обеих функций
  const historyData = readHistoryDataForSync_(historySheet)
  if (!historyData) {
    console.log('Sync: History не содержит данных')
    return
  }
  
  // Синхронизируем Invest
  syncInvestPricesFromHistory_(historySheet, historyData)
  
  // Синхронизируем Sales
  syncSalesPricesFromHistory_(historySheet, historyData)
  
  // Обновляем аналитику в Invest/Sales после синхронизации цен
  try {
    invest_syncMinMaxFromHistory()
    invest_syncTrendDaysFromHistory()
    invest_syncExtendedAnalyticsFromHistory()
    sales_syncMinMaxFromHistory()
    sales_syncTrendDaysFromHistory()
    sales_syncExtendedAnalyticsFromHistory()
    // Обновляем аналитику портфеля
    portfolioStats_update()
    console.log(`Sync: аналитика в Invest/Sales и портфеля обновлена`)
  } catch (e) {
    console.error('Sync: ошибка при обновлении аналитики в Invest/Sales:', e)
  }
  
  // Автоматический расчет метрик после синхронизации цен
  try {
    updateAllMetricsForInvestSales()
    console.log('Sync: метрики Invest/Sales обновлены')
  } catch (e) {
    console.error('Sync: ошибка при обновлении метрик Invest/Sales:', e)
  }
}

// Синхронизирует цены для Invest из History
function syncInvestPricesFromHistory_(historySheet, historyData = null) {
  const investSheet = getInvestSheet_()
  if (!investSheet) return
  
  const lastRow = investSheet.getLastRow()
  if (lastRow <= 1) return
  
  const count = lastRow - 1
  
  const nameColIndex = getColumnIndex(INVEST_COLUMNS.NAME)
  const priceColIndex = getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE)
  
  const names = investSheet.getRange(DATA_START_ROW, nameColIndex, count, 1).getValues()
  const currentPrices = investSheet.getRange(DATA_START_ROW, priceColIndex, count, 1).getValues()
  
  // Если данные History не переданы, читаем их
  if (!historyData) {
    historyData = readHistoryDataForSync_(historySheet)
    if (!historyData) {
      console.log('Sync Invest: History не содержит данных')
      return
    }
  }
  
  const { historyPriceMap } = historyData
  const backgroundsToSet = [] // Массив для batch-применения фонов
  
  let updatedCount = 0
  let errorCount = 0
  
  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue
    
    const itemHistoryData = historyPriceMap.get(name)
    if (!itemHistoryData) {
      currentPrices[i][0] = null
      errorCount++
      continue
    }
    
    // Получаем цену используя универсальную функцию
    const price = getCurrentPriceFromHistoryData_(itemHistoryData)
    currentPrices[i][0] = price
    updatedCount++
    
    const row = i + DATA_START_ROW
    // Определяем, устарела ли цена используя универсальную функцию
    const isOutdated = isPriceOutdated_(itemHistoryData)
    backgroundsToSet.push({ row, col: priceColIndex, color: isOutdated ? COLORS.STABLE : null })
  }
  
  // Batch-применение фонов
  backgroundsToSet.forEach(bg => {
    investSheet.getRange(bg.row, bg.col).setBackground(bg.color)
  })
  
  investSheet.getRange(DATA_START_ROW, priceColIndex, count, 1).setValues(currentPrices)
  
  invest_calculateBatch_(investSheet, currentPrices)
  
  console.log(`Sync Invest: обновлено ${updatedCount}, ошибок ${errorCount}`)
}

// Синхронизирует цены для Sales из History
function syncSalesPricesFromHistory_(historySheet, historyData = null) {
  const salesSheet = getSalesSheet_()
  if (!salesSheet) return
  
  const lastRow = salesSheet.getLastRow()
  if (lastRow <= 1) return
  
  const count = lastRow - 1
  
  const nameColIndex = getColumnIndex(SALES_COLUMNS.NAME)
  const priceColIndex = getColumnIndex(SALES_COLUMNS.CURRENT_PRICE)
  const sellColIndex = getColumnIndex(SALES_COLUMNS.SELL_PRICE)
  
  const names = salesSheet.getRange(DATA_START_ROW, nameColIndex, count, 1).getValues()
  const currentPrices = salesSheet.getRange(DATA_START_ROW, priceColIndex, count, 1).getValues()
  const sellPrices = salesSheet.getRange(DATA_START_ROW, sellColIndex, count, 1).getValues()
  
  // Если данные History не переданы, читаем их
  if (!historyData) {
    historyData = readHistoryDataForSync_(historySheet)
    if (!historyData) {
      console.log('Sync Sales: History не содержит данных')
      return
    }
  }
  
  const { historyPriceMap } = historyData
  const backgroundsToSet = [] // Массив для batch-применения фонов
  
  let updatedCount = 0
  let errorCount = 0
  
  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue
    
    const itemHistoryData = historyPriceMap.get(name)
    if (!itemHistoryData) {
      currentPrices[i][0] = null
      errorCount++
      continue
    }
    
    // Получаем цену используя универсальную функцию
    const price = getCurrentPriceFromHistoryData_(itemHistoryData)
    currentPrices[i][0] = price
    updatedCount++
    
    const row = i + DATA_START_ROW
    // Определяем, устарела ли цена используя универсальную функцию
    const isOutdated = isPriceOutdated_(itemHistoryData)
    backgroundsToSet.push({ row, col: priceColIndex, color: isOutdated ? COLORS.STABLE : null })
  }
  
  // Batch-применение фонов
  backgroundsToSet.forEach(bg => {
    salesSheet.getRange(bg.row, bg.col).setBackground(bg.color)
  })
  
  salesSheet.getRange(DATA_START_ROW, priceColIndex, count, 1).setValues(currentPrices)
  
  // ИСПРАВЛЕНИЕ: Используем sales_calculateBatch_() вместо ручного расчета priceDrop для консистентности
  sales_calculateBatch_(salesSheet, currentPrices, sellPrices)
  
  console.log(`Sync Sales: обновлено ${updatedCount}, ошибок ${errorCount}`)
}


function fetchLowestPriceWithBackoff_(appid, itemName, options) {
  const name = String(itemName || '').trim()
  if (!name) return { ok: false, error: 'empty_name' }
  const {
    attempts = 1,  // Одна попытка на предмет
    baseDelayMs = LIMITS.BASE_DELAY_MS,
    betweenItemsMs = LIMITS.API_BETWEEN_ITEMS_MS,
    timeBudgetMs = LIMITS.TIME_BUDGET_MS,
    startedAt = Date.now(),
  } = options || {}

  // Проверка бюджета времени перед началом
    if (Date.now() - startedAt > timeBudgetMs - (LIMITS.SAFETY_BUFFER_MS * 5)) {
    return { ok: false, error: 'time_budget_low' }
  }

  let delay = baseDelayMs
  let lastResult = null
  
  for (let attempt = 1; attempt <= attempts; attempt++) {
    // Проверка бюджета времени перед каждой попыткой
    if (Date.now() - startedAt > timeBudgetMs - LIMITS.SAFETY_BUFFER_MS) {
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
      let minInterval = options.betweenItemsMs || LIMITS.BETWEEN_ITEMS_MS
      
      // Если прошло больше 200мс с последнего запроса, можно ускориться
      if (timeSinceLastFetch > LIMITS.BASE_DELAY_MS) {
        minInterval = Math.max(LIMITS.MIN_INTERVAL_MS, minInterval - 50)
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
      const errorDelay = Math.min(delay, LIMITS.MAX_ERROR_DELAY_MS)
      Utilities.sleep(errorDelay)
      delay = Math.min(delay * 1.5, LIMITS.MAX_ERROR_DELAY_MS)
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
  return createLogSheet_(
    SHEET_NAMES.AUTO_LOG,
    ['Дата/время', 'Лист', 'Действие', 'Статус'],
    LIMITS.LOG_COLUMN_WIDTHS.AUTO_LOG
  )
}

/**
 * Универсальная функция для вставки новой строки в лист лога
 * @param {Sheet} sheet - Лист для вставки
 * @param {Array} values - Массив значений для вставки
 * @param {Object} formats - Объект с форматами {dateFormat: string, alignments: {horizontal: string, vertical: string}}
 */
function insertLogRowUniversal_(sheet, values, formats = {}) {
  // Вставляем новую строку сразу после заголовка (строка 2)
  // Существующие строки автоматически сдвигаются вниз
  const insertRow = HEADER_ROW + 1
  sheet.insertRowAfter(HEADER_ROW)
  
  // Сбрасываем форматирование заголовка для новой строки
  // После insertRowAfter новая строка может наследовать форматирование заголовка
  const numCols = values.length
  const newRowRange = sheet.getRange(insertRow, 1, 1, numCols)
  newRowRange.setBackground(null) // Сбрасываем фон
  newRowRange.setFontWeight('normal') // Сбрасываем жирный шрифт
  
  // Устанавливаем значения
  sheet.getRange(insertRow, 1, 1, numCols).setValues([values])
  
  // Применяем форматы
  if (formats.dateFormat) {
    sheet.getRange(insertRow, 1).setNumberFormat(formats.dateFormat)
  }
  
  const horizontalAlign = formats.alignments?.horizontal || 'center'
  const verticalAlign = formats.alignments?.vertical || 'middle'
  sheet.getRange(insertRow, 1, 1, numCols).setVerticalAlignment(verticalAlign).setHorizontalAlignment(horizontalAlign)
}

function logAutoAction_(sheetName, action, status = 'OK') {
  const sheet = getOrCreateAutoLogSheet_()
  const now = new Date()
  
  insertLogRowUniversal_(sheet, [now, sheetName, action, status], {
    dateFormat: 'dd.MM.yyyy HH:mm',
    alignments: { horizontal: 'center', vertical: 'middle' }
  })
}


function getOrCreateLogSheet_() {
  return createLogSheet_(
    SHEET_NAMES.LOG,
    ['Дата/время', 'Тип', 'Изображение', 'Предмет', 'Кол-во', 'Цена за шт', 'Сумма', 'Источник'],
    LIMITS.LOG_COLUMN_WIDTHS.LOG
  )
}

function logOperation_(type, itemName, quantity, pricePerUnit, total, source) {
  const sheet = getOrCreateLogSheet_()
  const now = new Date()
  const insertRow = HEADER_ROW + 1
  
  // Порядок колонок: A Дата, B Тип, C Изображение, D Предмет, E Кол-во, F Цена за шт, G Сумма, H Источник
  // Используем универсальную функцию для базовой вставки (изображение будет null, заполним потом)
  insertLogRowUniversal_(sheet, [now, type, null, itemName, quantity, pricePerUnit, total, source], {
    dateFormat: 'dd.MM.yyyy HH:mm',
    alignments: { horizontal: 'center', vertical: 'middle' }
  })
  
  // Используем тот же механизм изображений, что и в других листах
  const logConfig = {
    IMAGE: 'C',
    NAME: 'D',
    LINK: 'G' // Используем колонку G для ссылки
  }
  setImageAndLink_(sheet, insertRow, 570, itemName, logConfig)
  
  // Дополнительное форматирование для специфичных колонок
  sheet.getRange(insertRow, 4).setHorizontalAlignment('left') // Предмет - выравнивание влево
  sheet.getRange(insertRow, 5).setNumberFormat('0') // Кол-во - целое число
  sheet.getRange(insertRow, 6).setNumberFormat('#,##0.00 ₽') // Цена за шт
  sheet.getRange(insertRow, 7).setNumberFormat('#,##0.00 ₽') // Сумма
  
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
      Utilities.sleep(LIMITS.HISTORY_UPDATE_DELAY_MS)
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
  if (!columnLetter) {
    console.error('getColumnIndex: columnLetter не определен (undefined или null)')
    return 0
  }
  
  if (typeof columnLetter !== 'string') {
    console.error(`getColumnIndex: columnLetter должен быть строкой, получен ${typeof columnLetter}`)
    return 0
  }
  
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

  const history = getHistorySheet_()
  if (!history) return { updatedCount: 0 }

  const count = lastRow - 1
  const names = targetSheet.getRange(DATA_START_ROW, 2, count, 1).getValues()
  const minCol = targetSheet.getRange(DATA_START_ROW, minColIndex, count, 1).getValues()
  const maxCol = targetSheet.getRange(DATA_START_ROW, maxColIndex, count, 1).getValues()

  const outMin = minCol.map(r => [r[0]])
  const outMax = maxCol.map(r => [r[0]])
  let updatedCount = 0

  // ОПТИМИЗАЦИЯ: Читаем все данные History одним batch-запросом
  const historyLastRow = history.getLastRow()
  const historyLastCol = history.getLastColumn()
  
  if (historyLastRow <= 1 || historyLastCol < HISTORY_COLUMNS.FIRST_DATE_COL) {
    // Если History пустой, заполняем значениями по умолчанию
    for (let i = 0; i < count; i++) {
      const name = String(names[i][0] || '').trim()
      if (!name) continue
      
      if (updateAll || !outMin[i][0] || !outMax[i][0]) {
        outMin[i][0] = STATUS.NO_VALUES
        outMax[i][0] = STATUS.NO_VALUES
        updatedCount++
      }
    }
  } else {
    // Читаем все данные History одним batch-запросом
    const historyNames = history.getRange(DATA_START_ROW, 2, historyLastRow - 1, 1).getValues()
    const historyFirstDateCol = HISTORY_COLUMNS.FIRST_DATE_COL
    const historyDateCount = historyLastCol - historyFirstDateCol + 1
    const historyPriceData = history.getRange(DATA_START_ROW, historyFirstDateCol, historyLastRow - 1, historyDateCount).getValues()
    
    // Создаем Map для быстрого поиска Min/Max по имени
    const historyMinMaxMap = new Map()
    for (let h = 0; h < historyLastRow - 1; h++) {
      const historyName = String(historyNames[h][0] || '').trim()
      if (historyName && historyPriceData[h]) {
        // Вычисляем Min/Max из всех цен для этой строки
        const prices = historyPriceData[h].filter(v => typeof v === 'number' && !isNaN(v) && v > 0)
        if (prices.length > 0) {
          historyMinMaxMap.set(historyName, {
            min: Math.min(...prices),
            max: Math.max(...prices)
          })
        } else {
          historyMinMaxMap.set(historyName, { noValues: true })
        }
      }
    }
    
    // Синхронизируем данные
    for (let i = 0; i < count; i++) {
      const name = String(names[i][0] || '').trim()
      if (!name) continue

      if (!updateAll) {
        const hasMin = outMin[i][0] != null && outMin[i][0] !== ''
        const hasMax = outMax[i][0] != null && outMax[i][0] !== ''
        if (hasMin && hasMax) continue
      }

      const mm = historyMinMaxMap.get(name)
      if (mm && mm.min != null && mm.max != null) {
        outMin[i][0] = mm.min
        outMax[i][0] = mm.max
        updatedCount++
      } else if (!mm) {
        // Предмет не найден в History
        outMin[i][0] = STATUS.MISSING
        outMax[i][0] = STATUS.MISSING
        updatedCount++
      } else if (mm.noValues) {
        outMin[i][0] = STATUS.NO_VALUES
        outMax[i][0] = STATUS.NO_VALUES
        updatedCount++
      }
    }
  }

  targetSheet.getRange(DATA_START_ROW, minColIndex, count, 1).setValues(outMin)
  targetSheet.getRange(DATA_START_ROW, maxColIndex, count, 1).setValues(outMax)

  return { updatedCount }
}

// === УНИВЕРСАЛЬНАЯ ФУНКЦИЯ СИНХРОНИЗАЦИИ ТРЕНД ИЗ HISTORY ===
/**
 * Универсальная функция для синхронизации Тренд из History
 * Теперь Тренд содержит объединенный формат: "🟥 Падает 35 дн."
 * @param {Sheet} targetSheet - Лист назначения (Invest или Sales)
 * @param {number} trendColIndex - Индекс колонки Тренд
 * @param {boolean} updateAll - Обновлять все или только пустые
 * @returns {Object} {updatedCount}
 */
function syncTrendFromHistoryUniversal_(targetSheet, trendColIndex, updateAll) {
  const lastRow = targetSheet.getLastRow()
  if (lastRow <= 1) return { updatedCount: 0 }

  const history = getHistorySheet_()
  if (!history) return { updatedCount: 0 }

  const count = lastRow - 1
  const names = targetSheet.getRange(DATA_START_ROW, 2, count, 1).getValues()
  const trendCol = targetSheet.getRange(DATA_START_ROW, trendColIndex, count, 1).getValues()

  const outTrend = trendCol.map(r => [r[0]])
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
        updatedCount++
      }
    }
  } else {
    // Читаем все данные из History одним батчем
    const historyNames = history.getRange(DATA_START_ROW, 2, historyLastRow - 1, 1).getValues()
    const historyTrendCol = getColumnIndex(HISTORY_COLUMNS.TREND)
    const historyTrends = history.getRange(DATA_START_ROW, historyTrendCol, historyLastRow - 1, 1).getValues()

    // Создаем мапу для быстрого поиска
    const historyMap = new Map()
    for (let i = 0; i < historyNames.length; i++) {
      const hName = String(historyNames[i][0] || '').trim()
      if (hName) {
        historyMap.set(hName, {
          trend: historyTrends[i][0] || '🟪'
        })
      }
    }

    // Синхронизируем данные
    for (let i = 0; i < count; i++) {
      const name = String(names[i][0] || '').trim()
      if (!name) continue

      if (!updateAll) {
        const hasTrend = outTrend[i][0] != null && outTrend[i][0] !== ''
        if (hasTrend) continue
      }

      const historyData = historyMap.get(name)
      if (historyData) {
        outTrend[i][0] = historyData.trend
        updatedCount++
      } else {
        // Предмет не найден в History - ставим значения по умолчанию
        outTrend[i][0] = '🟪'
        updatedCount++
      }
    }
  }

  targetSheet.getRange(DATA_START_ROW, trendColIndex, count, 1).setValues(outTrend)

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

  const history = getHistorySheet_()
  if (!history) return { updatedCount: 0 }

  const count = lastRow - 1
  const names = targetSheet.getRange(DATA_START_ROW, 2, count, 1).getValues()
  
  // Поддержка null для phase и potential (для Sales, где они не используются)
  const phaseCol = phaseColIndex ? targetSheet.getRange(DATA_START_ROW, phaseColIndex, count, 1).getValues() : null
  const potentialCol = potentialColIndex ? targetSheet.getRange(DATA_START_ROW, potentialColIndex, count, 1).getValues() : null
  const recommendationCol = targetSheet.getRange(DATA_START_ROW, recommendationColIndex, count, 1).getValues()

  const outPhase = phaseCol ? phaseCol.map(r => [r[0]]) : null
  const outPotential = potentialCol ? potentialCol.map(r => [r[0]]) : null
  const outRecommendation = recommendationCol.map(r => [r[0]])
  let updatedCount = 0

  // Получаем данные из History
  const historyLastRow = history.getLastRow()
  
  if (historyLastRow <= 1) {
    // Если History пустой, заполняем значениями по умолчанию
    for (let i = 0; i < count; i++) {
      const name = String(names[i][0] || '').trim()
      if (!name) continue
      
      if (updateAll || !outRecommendation[i][0]) {
        if (outPhase) outPhase[i][0] = '❓'
        if (outPotential) outPotential[i][0] = null  // null вместо '—' для числового формата
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
        const hasPhase = outPhase ? (outPhase[i][0] != null && outPhase[i][0] !== '') : true
        const hasPotential = outPotential ? (outPotential[i][0] != null && outPotential[i][0] !== '') : true
        const hasRecommendation = outRecommendation[i][0] != null && outRecommendation[i][0] !== ''
        if (hasPhase && hasPotential && hasRecommendation) continue
      }

      const historyData = historyMap.get(name)
      if (historyData) {
        if (outPhase) outPhase[i][0] = historyData.phase
        if (outPotential) outPotential[i][0] = historyData.potential
        outRecommendation[i][0] = historyData.recommendation
        updatedCount++
      } else {
        // Предмет не найден в History - ставим значения по умолчанию
        if (outPhase) outPhase[i][0] = '❓'
        if (outPotential) outPotential[i][0] = null  // null вместо '—' для числового формата
        outRecommendation[i][0] = '👀 НАБЛЮДАТЬ'
        updatedCount++
      }
    }
  }

  if (outPhase && phaseColIndex) {
    targetSheet.getRange(DATA_START_ROW, phaseColIndex, count, 1).setValues(outPhase)
  }
  if (outPotential && potentialColIndex) {
    targetSheet.getRange(DATA_START_ROW, potentialColIndex, count, 1).setValues(outPotential)
  }
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

