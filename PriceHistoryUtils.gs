/**
 * PriceHistoryUtils - Работа с pricehistory эндпоинтом Steam
 * 
 * Утилита для расчета Min/Max цен из всей истории предмета через pricehistory
 */

/**
 * Получает историю цен для предмета через pricehistory эндпоинт
 * @param {string} itemName - Название предмета
 * @param {number} appid - ID приложения Steam (570 для Dota 2)
 * @returns {Object} {ok: boolean, prices?: Array, error?: string}
 */
function priceHistory_fetchHistoryForItem(itemName, appid = STEAM_APP_ID) {
  const name = String(itemName || '').trim()
  if (!name) {
    return { ok: false, error: 'empty_name' }
  }
  
  const url = `${API_CONFIG.STEAM_PRICE_HISTORY.BASE_URL}/?country=RU&currency=5&appid=${appid}&market_hash_name=${encodeURIComponent(name)}`
  
  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      timeout: API_CONFIG.STEAM_PRICE_HISTORY.TIMEOUT_MS
    })
    
    const responseCode = response.getResponseCode()
    if (responseCode !== 200) {
      console.error(`PriceHistory: HTTP ошибка ${responseCode} для "${name}"`)
      return { ok: false, error: `http_${responseCode}` }
    }
    
    const responseText = response.getContentText()
    const data = JSON.parse(responseText)
    
    if (!data || !data.success || !Array.isArray(data.prices)) {
      console.warn(`PriceHistory: некорректный ответ для "${name}"`)
      return { ok: false, error: 'invalid_format' }
    }
    
    return { ok: true, prices: data.prices }
    
  } catch (e) {
    console.error(`PriceHistory: исключение при получении истории для "${name}":`, e)
    return { ok: false, error: 'exception', details: e.message }
  }
}

/**
 * Парсит данные истории цен из формата pricehistory
 * @param {Array} pricesData - Массив цен из pricehistory API
 * @returns {Array} Массив объектов {date: Date, price: number, volume: number}
 */
function priceHistory_parseHistoryData(pricesData) {
  if (!Array.isArray(pricesData)) {
    return []
  }
  
  const parsed = []
  
  pricesData.forEach(record => {
    if (!Array.isArray(record) || record.length < 2) {
      return // Пропускаем некорректные записи
    }
    
    try {
      const dateStr = String(record[0] || '').trim() // "Sep 01 2020 01: +0"
      const price = Number(record[1])
      const volume = record.length > 2 ? parseInt(String(record[2] || '0'), 10) : 0
      
      if (!Number.isFinite(price) || price <= 0) {
        return // Пропускаем некорректные цены
      }
      
      // Парсим дату из формата "Sep 01 2020 01: +0"
      // Формат: "MMM DD YYYY HH: +0" или "MMM DD YYYY HH:MM"
      const dateParts = dateStr.split(' ')
      if (dateParts.length >= 3) {
        const monthStr = dateParts[0] // "Sep"
        const day = parseInt(dateParts[1], 10)
        const year = parseInt(dateParts[2], 10)
        
        const monthMap = {
          'Jan': 0, 'Feb': 1, 'Mar': 2, 'Apr': 3, 'May': 4, 'Jun': 5,
          'Jul': 6, 'Aug': 7, 'Sep': 8, 'Oct': 9, 'Nov': 10, 'Dec': 11
        }
        
        const month = monthMap[monthStr] !== undefined ? monthMap[monthStr] : 0
        
        // Создаем дату (используем полдень для избежания проблем с часовыми поясами)
        const date = new Date(year, month, day, 12, 0, 0)
        
        parsed.push({
          date: date,
          price: price,
          volume: volume
        })
      }
    } catch (e) {
      console.warn('PriceHistory: ошибка парсинга записи:', record, e)
    }
  })
  
  return parsed
}

/**
 * Рассчитывает Min и Max цены из истории
 * @param {Array} parsedHistory - Массив объектов {date, price, volume}
 * @returns {Object} {min: number, max: number, minDate: Date, maxDate: Date}
 */
function priceHistory_calculateMinMax(parsedHistory) {
  if (!Array.isArray(parsedHistory) || parsedHistory.length === 0) {
    return { min: null, max: null, minDate: null, maxDate: null }
  }
  
  let min = Infinity
  let max = -Infinity
  let minDate = null
  let maxDate = null
  
  parsedHistory.forEach(record => {
    if (record.price < min) {
      min = record.price
      minDate = record.date
    }
    if (record.price > max) {
      max = record.price
      maxDate = record.date
    }
  })
  
  return {
    min: min === Infinity ? null : min,
    max: max === -Infinity ? null : max,
    minDate: minDate,
    maxDate: maxDate
  }
}

/**
 * Обновляет Min/Max для одного предмета в History
 * @param {string} itemName - Название предмета
 * @param {number} minPrice - Минимальная цена
 * @param {number} maxPrice - Максимальная цена
 * @returns {boolean} true если обновлено успешно
 */
function priceHistory_updateMinMaxInHistory(itemName, minPrice, maxPrice) {
  const historySheet = getHistorySheet_()
  if (!historySheet) {
    return false
  }
  
  const lastRow = historySheet.getLastRow()
  if (lastRow < DATA_START_ROW) {
    return false
  }
  
  // Ищем строку с предметом
  const names = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), lastRow - HEADER_ROW, 1).getValues()
  
  for (let i = 0; i < names.length; i++) {
    const rowName = String(names[i][0] || '').trim()
    if (rowName === itemName) {
      const row = DATA_START_ROW + i
      
      // Обновляем Min и Max только если новые значения валидны
      if (Number.isFinite(minPrice) && minPrice > 0) {
        historySheet.getRange(row, getColumnIndex(HISTORY_COLUMNS.MIN_PRICE)).setValue(minPrice)
      }
      if (Number.isFinite(maxPrice) && maxPrice > 0) {
        historySheet.getRange(row, getColumnIndex(HISTORY_COLUMNS.MAX_PRICE)).setValue(maxPrice)
      }
      
      return true
    }
  }
  
  return false
}

/**
 * Рассчитывает Min/Max для всех предметов из History через pricehistory
 * Можно вызывать через меню или автоматически при добавлении нового предмета
 */
function priceHistory_calculateMinMaxForAllItems() {
  console.log('PriceHistory: начало расчета Min/Max для всех предметов')
  
  const historySheet = getHistorySheet_()
  if (!historySheet) {
    console.error('PriceHistory: лист History не найден')
    SpreadsheetApp.getUi().alert('Ошибка: лист History не найден')
    return
  }
  
  const lastRow = historySheet.getLastRow()
  if (lastRow < DATA_START_ROW) {
    console.log('PriceHistory: нет предметов в History')
    SpreadsheetApp.getUi().alert('Нет предметов в History для расчета Min/Max')
    return
  }
  
  // Получаем все названия предметов
  const names = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), lastRow - HEADER_ROW, 1).getValues()
  const uniqueNames = []
  const nameSet = new Set()
  
  names.forEach(row => {
    const name = String(row[0] || '').trim()
    if (name && !nameSet.has(name)) {
      uniqueNames.push(name)
      nameSet.add(name)
    }
  })
  
  console.log(`PriceHistory: найдено ${uniqueNames.length} уникальных предметов`)
  
  if (uniqueNames.length === 0) {
    SpreadsheetApp.getUi().alert('Нет предметов для расчета Min/Max')
    return
  }
  
  let successCount = 0
  let errorCount = 0
  let skippedCount = 0
  const startedAt = Date.now()
  
  // Показываем прогресс
  const ui = SpreadsheetApp.getUi()
  const progressResponse = ui.alert(
    'Расчет Min/Max',
    `Будет обработано ${uniqueNames.length} предметов.\nЭто может занять несколько минут.\n\nПродолжить?`,
    ui.ButtonSet.YES_NO
  )
  
  if (progressResponse !== ui.Button.YES) {
    return
  }
  
  // Обрабатываем каждый предмет
  for (let i = 0; i < uniqueNames.length; i++) {
    const itemName = uniqueNames[i]
    
    // Проверка бюджета времени (максимум 10 минут на операцию)
    if (Date.now() - startedAt > 600000) {
      console.warn(`PriceHistory: превышен бюджет времени, обработано ${i} из ${uniqueNames.length}`)
      break
    }
    
    try {
      // Получаем историю цен
      const historyResult = priceHistory_fetchHistoryForItem(itemName, STEAM_APP_ID)
      
      if (!historyResult.ok || !historyResult.prices) {
        errorCount++
        console.warn(`PriceHistory: не удалось получить историю для "${itemName}": ${historyResult.error}`)
        Utilities.sleep(500) // Пауза между запросами
        continue
      }
      
      // Парсим данные
      const parsedHistory = priceHistory_parseHistoryData(historyResult.prices)
      
      if (parsedHistory.length === 0) {
        skippedCount++
        console.warn(`PriceHistory: нет валидных данных в истории для "${itemName}"`)
        continue
      }
      
      // Рассчитываем Min/Max
      const minMax = priceHistory_calculateMinMax(parsedHistory)
      
      if (!minMax.min || !minMax.max) {
        skippedCount++
        console.warn(`PriceHistory: не удалось рассчитать Min/Max для "${itemName}"`)
        continue
      }
      
      // Обновляем в History
      const updated = priceHistory_updateMinMaxInHistory(itemName, minMax.min, minMax.max)
      
      if (updated) {
        successCount++
        console.log(`PriceHistory: обновлено Min/Max для "${itemName}": Min=${minMax.min}, Max=${minMax.max}`)
      } else {
        errorCount++
        console.warn(`PriceHistory: не найдена строка для "${itemName}" в History`)
      }
      
      // Пауза между запросами (чтобы не перегружать Steam API)
      Utilities.sleep(800)
      
    } catch (e) {
      errorCount++
      console.error(`PriceHistory: исключение при обработке "${itemName}":`, e)
    }
  }
  
  // Показываем результат
  const message = `Расчет Min/Max завершен!\n\n` +
    `Успешно: ${successCount}\n` +
    `Ошибок: ${errorCount}\n` +
    `Пропущено: ${skippedCount}\n` +
    `Всего: ${uniqueNames.length}`
  
  console.log(`PriceHistory: ${message}`)
  
  try {
    logAutoAction_('PriceHistory', 'Расчет Min/Max для всех предметов', `OK (успешно: ${successCount}, ошибок: ${errorCount}, пропущено: ${skippedCount})`)
    ui.alert(message)
  } catch (e) {
    console.log('PriceHistory: невозможно показать UI')
  }
}

/**
 * Рассчитывает Min/Max для одного предмета (вспомогательная функция)
 * @param {string} itemName - Название предмета
 * @returns {Object} {ok: boolean, min?: number, max?: number, error?: string}
 */
function priceHistory_calculateMinMaxForSingleItem(itemName) {
  const name = String(itemName || '').trim()
  if (!name) {
    return { ok: false, error: 'empty_name' }
  }
  
  const historyResult = priceHistory_fetchHistoryForItem(name, STEAM_APP_ID)
  
  if (!historyResult.ok || !historyResult.prices) {
    return { ok: false, error: historyResult.error || 'fetch_failed' }
  }
  
  const parsedHistory = priceHistory_parseHistoryData(historyResult.prices)
  
  if (parsedHistory.length === 0) {
    return { ok: false, error: 'no_valid_data' }
  }
  
  const minMax = priceHistory_calculateMinMax(parsedHistory)
  
  if (!minMax.min || !minMax.max) {
    return { ok: false, error: 'calculation_failed' }
  }
  
  return {
    ok: true,
    min: minMax.min,
    max: minMax.max,
    minDate: minMax.minDate,
    maxDate: minMax.maxDate
  }
}

