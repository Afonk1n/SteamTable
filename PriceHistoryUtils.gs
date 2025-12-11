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
    if (responseCode !== HTTP_STATUS.OK) {
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
 * Проверяет, есть ли уже Min/Max у предмета в History
 * @param {string} itemName - Название предмета
 * @returns {Object} {hasMin: boolean, hasMax: boolean, minValue: number|null, maxValue: number|null}
 */
function priceHistory_checkMinMaxExists(itemName) {
  const historySheet = getHistorySheet_()
  if (!historySheet) {
    return { hasMin: false, hasMax: false, minValue: null, maxValue: null }
  }
  
  const lastRow = historySheet.getLastRow()
  if (lastRow < DATA_START_ROW) {
    return { hasMin: false, hasMax: false, minValue: null, maxValue: null }
  }
  
  const names = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), lastRow - HEADER_ROW, 1).getValues()
  
  for (let i = 0; i < names.length; i++) {
    const rowName = String(names[i][0] || '').trim()
    if (rowName === itemName) {
      const row = DATA_START_ROW + i
      const minCell = historySheet.getRange(row, getColumnIndex(HISTORY_COLUMNS.MIN_PRICE))
      const maxCell = historySheet.getRange(row, getColumnIndex(HISTORY_COLUMNS.MAX_PRICE))
      const minValue = minCell.getValue()
      const maxValue = maxCell.getValue()
      
      const hasMin = minValue !== null && minValue !== '' && Number.isFinite(Number(minValue)) && Number(minValue) > 0
      const hasMax = maxValue !== null && maxValue !== '' && Number.isFinite(Number(maxValue)) && Number(maxValue) > 0
      
      return {
        hasMin: hasMin,
        hasMax: hasMax,
        minValue: hasMin ? Number(minValue) : null,
        maxValue: hasMax ? Number(maxValue) : null
      }
    }
  }
  
  return { hasMin: false, hasMax: false, minValue: null, maxValue: null }
}

/**
 * Рассчитывает Min/Max для всех предметов из History через SteamWebAPI.ru
 * Использует SteamWebAPI.ru вместо pricehistory (быстрее, надежнее, до 50 предметов за раз)
 * @param {boolean} onlyMissing - Если true, обновляет только предметы без Min/Max
 */
function priceHistory_calculateMinMaxForAllItems(onlyMissing = false) {
  console.log('PriceHistory: начало расчета Min/Max для всех предметов (через SteamWebAPI.ru)')
  
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
  
  // Оптимизированная проверка Min/Max: batch чтение вместо проверки по одной строке
  let itemsToProcess = uniqueNames
  if (onlyMissing) {
    // Читаем все Min/Max значения за один раз
    const historySheet = getHistorySheet_()
    const lastRow = historySheet.getLastRow()
    
    if (lastRow >= DATA_START_ROW) {
      const count = lastRow - HEADER_ROW
      const namesBatch = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), count, 1).getValues()
      const minBatch = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.MIN_PRICE), count, 1).getValues()
      const maxBatch = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.MAX_PRICE), count, 1).getValues()
      
      // Создаем карту существующих Min/Max
      const existingMap = new Map()
      for (let i = 0; i < count; i++) {
        const rowName = String(namesBatch[i][0] || '').trim()
        if (rowName) {
          const minValue = minBatch[i][0]
          const maxValue = maxBatch[i][0]
          const hasMin = minValue !== null && minValue !== '' && Number.isFinite(Number(minValue)) && Number(minValue) > 0
          const hasMax = maxValue !== null && maxValue !== '' && Number.isFinite(Number(maxValue)) && Number(maxValue) > 0
          existingMap.set(rowName, { hasMin, hasMax })
        }
      }
      
      // Фильтруем только предметы без Min/Max
      itemsToProcess = uniqueNames.filter(name => {
        const exists = existingMap.get(name)
        return !exists || !exists.hasMin || !exists.hasMax
      })
    } else {
      itemsToProcess = uniqueNames
    }
    
    console.log(`PriceHistory: найдено ${itemsToProcess.length} предметов без Min/Max (из ${uniqueNames.length} всего)`)
  } else {
    console.log(`PriceHistory: найдено ${uniqueNames.length} уникальных предметов`)
  }
  
  if (itemsToProcess.length === 0) {
    const message = onlyMissing 
      ? 'Все предметы уже имеют Min/Max значения'
      : 'Нет предметов для расчета Min/Max'
    SpreadsheetApp.getUi().alert(message)
    return
  }
  
  // Диалог с подтверждением
  const ui = SpreadsheetApp.getUi()
  const modeText = onlyMissing ? 'только у отсутствующих' : 'для всех предметов'
  const confirmResponse = ui.alert(
    'Расчет Min/Max',
    `Будет обработано ${itemsToProcess.length} предметов (${modeText}).\nЭто может занять несколько минут.\n\nПродолжить?`,
    ui.ButtonSet.YES_NO
  )
  
  if (confirmResponse !== ui.Button.YES) {
    return
  }
  
  const finalItemsToProcess = itemsToProcess
  
  let successCount = 0
  let errorCount = 0
  let skippedCount = 0
  const startedAt = Date.now()
  
  // Обрабатываем предметы пакетами по 50 штук (SteamWebAPI.ru поддерживает до 50 предметов за раз!)
  const batchSize = API_CONFIG.STEAM_WEB_API.MAX_ITEMS_PER_REQUEST
  let processedCount = 0
  
  for (let batchStart = 0; batchStart < finalItemsToProcess.length; batchStart += batchSize) {
    // Проверка бюджета времени (максимум 5 минут на операцию)
    if (Date.now() - startedAt > 300000) {
      console.warn(`PriceHistory: превышен бюджет времени, обработано ${processedCount} из ${finalItemsToProcess.length}`)
      break
    }
    
    const batch = finalItemsToProcess.slice(batchStart, batchStart + batchSize)
    console.log(`PriceHistory: обработка пакета ${Math.floor(batchStart / batchSize) + 1} (${batch.length} предметов из ${finalItemsToProcess.length})`)
    
    try {
      // Получаем данные для пакета предметов через SteamWebAPI.ru
      // Используем fetchItemsBatch с fallback для автоматической обработки не найденных предметов
      const batchResult = steamWebAPI_fetchItemsBatch(batch, 'dota2', true)
      
      if (!batchResult.ok || !batchResult.items) {
        errorCount += batch.length
        console.warn(`PriceHistory: ошибка получения данных для пакета: ${batchResult.error || 'unknown'}`)
        Utilities.sleep(2000) // Пауза перед следующим пакетом
        continue
      }
      
      // batchResult.items может быть пустым объектом, если все предметы не найдены
      // Это нормально, fallback уже попытался их найти
      
      // Создаем карту результатов с несколькими ключами для гибкого поиска
      // batchResult.items - это объект {key: item}, нужно преобразовать в массив
      const itemsMap = {}
      const itemsArray = Object.values(batchResult.items || {})
      itemsArray.forEach(item => {
        // Добавляем предмет по нескольким ключам для гибкого поиска
        const keys = [
          (item.normalizedname || '').toLowerCase().trim(),
          (item.marketname || '').toLowerCase().trim(),
          (item.markethashname || '').toLowerCase().trim()
        ].filter(k => k && k.length > 0)
        
        // Также добавляем ключ с удалением специальных символов (апострофы, кавычки и т.д.)
        keys.forEach(key => {
          const normalizedKey = key.replace(/[''""`]/g, '').replace(/\s+/g, ' ').trim()
          if (normalizedKey) {
            keys.push(normalizedKey)
          }
        })
        
        // Убираем дубликаты и добавляем в карту
        const uniqueKeys = [...new Set(keys)]
        uniqueKeys.forEach(key => {
          if (!itemsMap[key]) {
            itemsMap[key] = item
          }
        })
      })
      
      // Обрабатываем каждый предмет из пакета
      batch.forEach(itemName => {
        try {
          // Пробуем найти предмет по разным вариантам названия
          const searchVariants = [
            itemName.toLowerCase().trim(),
            itemName.toLowerCase().trim().replace(/[''""`]/g, ''), // Без апострофов
            itemName.toLowerCase().trim().replace(/[''""`]/g, '').replace(/\s+/g, ' ') // Без апострофов и лишних пробелов
          ]
          
          let itemData = null
          for (const searchKey of searchVariants) {
            if (itemsMap[searchKey]) {
              itemData = itemsMap[searchKey]
              break
            }
          }
          
          // Если не нашли точное совпадение, ищем похожее (fuzzy match)
          if (!itemData) {
            const searchKeyBase = itemName.toLowerCase().trim().replace(/[''""`]/g, '').replace(/\s+/g, ' ')
            for (const key in itemsMap) {
              const keyBase = key.replace(/[''""`]/g, '').replace(/\s+/g, ' ')
              if (keyBase === searchKeyBase || key.includes(searchKeyBase) || searchKeyBase.includes(key)) {
                itemData = itemsMap[key]
                console.log(`PriceHistory: найдено приблизительное совпадение "${itemName}" -> "${itemsMap[key].marketname || itemsMap[key].markethashname}"`)
                break
              }
            }
          }
          
          // Fallback уже выполнен в steamWebAPI_fetchItemsBatch, проверяем результат еще раз
          // Пробуем найти в batchResult.items напрямую (fallback мог добавить предмет)
          if (!itemData && batchResult.items) {
            const searchKeys = [
              itemName.toLowerCase().trim(),
              itemName.toLowerCase().trim().replace(/[''""`]/g, ''),
              itemName.toLowerCase().trim().replace(/[''""`]/g, '').replace(/\s+/g, ' ')
            ]
            
            for (const key of searchKeys) {
              if (batchResult.items[key]) {
                itemData = batchResult.items[key]
                console.log(`PriceHistory: предмет "${itemName}" найден через fallback в batchResult`)
                break
              }
            }
          }
          
          if (!itemData) {
            errorCount++
            console.warn(`PriceHistory: предмет "${itemName}" не найден ни через основной endpoint, ни через item_by_nameid`)
            return
          }
          
          // Извлекаем Min/Max напрямую из ответа SteamWebAPI.ru (pricemin и pricemax)
          // Проверяем наличие полей (могут быть строками или числами)
          let minPrice = null
          let maxPrice = null
          
          if (itemData.pricemin !== null && itemData.pricemin !== undefined && itemData.pricemin !== '') {
            minPrice = Number(itemData.pricemin)
          }
          
          if (itemData.pricemax !== null && itemData.pricemax !== undefined && itemData.pricemax !== '') {
            maxPrice = Number(itemData.pricemax)
          }
          
          // Логируем для отладки, если данные некорректны
          if (!Number.isFinite(minPrice) || !Number.isFinite(maxPrice) || minPrice <= 0 || maxPrice <= 0) {
            skippedCount++
            console.warn(`PriceHistory: некорректные Min/Max для "${itemName}": Min=${minPrice}, Max=${maxPrice}. Данные предмета: marketname="${itemData.marketname || ''}", markethashname="${itemData.markethashname || ''}", pricemin="${itemData.pricemin}" (type: ${typeof itemData.pricemin}), pricemax="${itemData.pricemax}" (type: ${typeof itemData.pricemax})`)
            return
          }
          
          // Обновляем в History
          const updated = priceHistory_updateMinMaxInHistory(itemName, minPrice, maxPrice)
          
          if (updated) {
            successCount++
            processedCount++
            console.log(`PriceHistory: обновлено Min/Max для "${itemName}": Min=${minPrice}, Max=${maxPrice}`)
          } else {
            errorCount++
            console.warn(`PriceHistory: не найдена строка для "${itemName}" в History`)
          }
          
        } catch (e) {
          errorCount++
          console.error(`PriceHistory: исключение при обработке "${itemName}":`, e)
        }
      })
      
      // Пауза между пакетами (SteamWebAPI.ru может иметь лимиты)
      if (batchStart + batchSize < itemsToProcess.length) {
        Utilities.sleep(1000) // 1 сек между пакетами достаточно
      }
      
    } catch (e) {
      errorCount += batch.length
      console.error(`PriceHistory: исключение при обработке пакета:`, e)
      Utilities.sleep(2000)
    }
  }
  
  // Показываем результат
  const modeTextResult = onlyMissing ? ' (только отсутствующие)' : ''
  const message = `Расчет Min/Max завершен${modeTextResult}!\n\n` +
    `Успешно: ${successCount}\n` +
    `Ошибок: ${errorCount}\n` +
    `Пропущено: ${skippedCount}\n` +
    `Всего обработано: ${itemsToProcess.length}`
  
  console.log(`PriceHistory: ${message}`)
  
  try {
    logAutoAction_('PriceHistory', `Расчет Min/Max${modeTextResult}`, `OK (успешно: ${successCount}, ошибок: ${errorCount}, пропущено: ${skippedCount})`)
    ui.alert(message)
  } catch (e) {
    console.log('PriceHistory: невозможно показать UI')
  }
}

/**
 * Рассчитывает Min/Max только для предметов, у которых Min/Max отсутствуют
 * Полезно для продолжения после принудительного завершения скрипта
 */
function priceHistory_calculateMinMaxForMissingItems() {
  priceHistory_calculateMinMaxForAllItems(true)
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

