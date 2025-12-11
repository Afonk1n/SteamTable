/**
 * SteamWebAPI - Интеграция с SteamWebAPI.ru
 * 
 * Замена текущих запросов к Steam + получение дополнительных данных
 * Работает без API ключа, бесплатно
 */

/**
 * Получает данные о предметах через SteamWebAPI.ru
 * @param {Array<string>} itemNames - Массив названий предметов (до 50)
 * @param {string} game - Игра ('dota2', 'cs2', 'rust', 'tf2')
 * @returns {Object} {ok: boolean, items?: Array, error?: string}
 */
function steamWebAPI_fetchItems(itemNames, game = 'dota2') {
  if (!itemNames || itemNames.length === 0) {
    return { ok: false, error: 'empty_items' }
  }
  
  if (itemNames.length > API_CONFIG.STEAM_WEB_API.MAX_ITEMS_PER_REQUEST) {
    return { ok: false, error: 'too_many_items', maxItems: API_CONFIG.STEAM_WEB_API.MAX_ITEMS_PER_REQUEST }
  }
  
  // Формируем URL с параметрами
  // Согласно документации API: параметр name - массив со стилем form, explode: true
  // Формат: name=Item1&name=Item2 (без квадратных скобок, несколько раз)
  const baseUrl = `${API_CONFIG.STEAM_WEB_API.BASE_URL}/api/item`
  const params = []
  params.push(`game=${encodeURIComponent(game)}`)
  
  // Добавляем все названия предметов как отдельные параметры name=
  itemNames.forEach(name => {
    params.push(`name=${encodeURIComponent(name)}`)
  })
  
  const url = `${baseUrl}?${params.join('&')}`
  
  const options = {
    method: 'GET',
    muteHttpExceptions: true
  }
  
  const result = fetchWithRetry_(url, options, {
    maxRetries: API_CONFIG.STEAM_WEB_API.MAX_RETRIES,
    retryDelayMs: API_CONFIG.STEAM_WEB_API.RETRY_DELAY_MS,
    apiName: 'SteamWebAPI'
  })
  
  if (!result.ok) {
    return result
  }
  
  try {
    const responseText = result.response.getContentText()
    const items = JSON.parse(responseText)
    
    if (!Array.isArray(items)) {
      console.error('SteamWebAPI: неожиданный формат ответа (не массив)')
      return { ok: false, error: 'invalid_format', details: 'Response is not an array' }
    }
    
    return { ok: true, items: items }
  } catch (e) {
    console.error('SteamWebAPI: ошибка парсинга ответа:', e)
    return { ok: false, error: 'parse_error', details: e.message }
  }
}

/**
 * Получает данные об одном предмете
 * @param {string} itemName - Название предмета
 * @param {string} game - Игра ('dota2', 'cs2', 'rust', 'tf2')
 * @returns {Object} {ok: boolean, item?: Object, error?: string}
 */
function steamWebAPI_fetchSingleItem(itemName, game = 'dota2') {
  const result = steamWebAPI_fetchItems([itemName], game)
  
  if (!result.ok) {
    return result
  }
  
  if (result.items.length === 0) {
    return { ok: false, error: 'item_not_found' }
  }
  
  return { ok: true, item: result.items[0] }
}

/**
 * Унифицированный интерфейс для получения данных о предмете
 * Возвращает данные в формате совместимом с текущим кодом
 * @param {string} itemName - Название предмета
 * @param {string} game - Игра ('dota2', 'cs2', 'rust', 'tf2')
 * @returns {Object} {ok: boolean, price?: number, data?: Object, error?: string}
 */
function steamWebAPI_getItemData(itemName, game = 'dota2') {
  const result = steamWebAPI_fetchSingleItem(itemName, game)
  
  if (!result.ok || !result.item) {
    return result
  }
  
  const item = result.item
  
  // Извлекаем основную цену (pricelatest)
  const price = item.pricelatest || item.price || null
  
  return {
    ok: true,
    price: price,
    data: item
  }
}

/**
 * Получает имя героя из тегов предмета (tag5 часто содержит имя героя)
 * @param {Object} itemData - Данные предмета из SteamWebAPI.ru
 * @returns {string|null} Имя героя или null
 */
function steamWebAPI_getHeroNameFromTags(itemData) {
  if (!itemData) {
    return null
  }
  
  // tag5 часто содержит имя героя
  if (itemData.tag5 && typeof itemData.tag5 === 'string' && itemData.tag5.trim().length > 0) {
    return itemData.tag5.trim()
  }
  
  // Можно также проверить другие теги, но tag5 - основной
  return null
}

/**
 * Парсит данные предмета из ответа API в удобный формат
 * @param {Object} itemData - Сырые данные из API
 * @returns {Object} Обработанные данные
 */
function steamWebAPI_parseItemData(itemData) {
  if (!itemData) {
    return null
  }
  
  return {
    // Основные данные
    name: itemData.marketname || itemData.markethashname || null,
    normalizedName: itemData.normalizedname || null,
    
    // Цены
    priceLatest: itemData.pricelatest || null,
    priceAvg24h: itemData.priceavg24h || null,
    priceAvg7d: itemData.priceavg7d || null,
    priceAvg30d: itemData.priceavg30d || null,
    priceAvg90d: itemData.priceavg90d || null,
    priceMin: itemData.pricemin || null,
    priceMax: itemData.pricemax || null,
    priceMedian: itemData.pricemedian || null,
    
    // Объемы продаж
    sold24h: itemData.sold24h || 0,
    sold7d: itemData.sold7d || 0,
    sold30d: itemData.sold30d || 0,
    sold90d: itemData.sold90d || 0,
    soldToday: itemData.soldtoday || 0,
    
    // Ордера и глубина рынка
    buyOrderPrice: itemData.buyorderprice || null,
    buyOrderVolume: itemData.buyordervolume || 0,
    offerVolume: itemData.offervolume || 0,
    
    // Теги (для определения героя)
    tag1: itemData.tag1 || null,
    tag2: itemData.tag2 || null,
    tag3: itemData.tag3 || null,
    tag4: itemData.tag4 || null,
    tag5: itemData.tag5 || null,  // Часто содержит имя героя
    
    // Последние продажи
    latestSales: itemData.latest10steamsales || [],
    
    // Метаданные
    classId: itemData.classid || null,
    instanceId: itemData.instanceid || null,
    imageUrl: itemData.itemimage || null,
    steamUrl: itemData.steamurl || null
  }
}

/**
 * Получает данные для нескольких предметов пакетами (batch)
 * @param {Array<string>} itemNames - Массив названий предметов
 * @param {string} game - Игра ('dota2', 'cs2', 'rust', 'tf2')
 * @returns {Object} {ok: boolean, items?: Object, errors?: Array}
 */
function steamWebAPI_fetchItemsBatch(itemNames, game = 'dota2') {
  if (!itemNames || itemNames.length === 0) {
    return { ok: false, error: 'empty_items' }
  }
  
  const maxItemsPerRequest = API_CONFIG.STEAM_WEB_API.MAX_ITEMS_PER_REQUEST
  const result = {}
  const errors = []
  
  // Разбиваем на пакеты по 50 предметов
  for (let i = 0; i < itemNames.length; i += maxItemsPerRequest) {
    const batch = itemNames.slice(i, i + maxItemsPerRequest)
    
    const batchResult = steamWebAPI_fetchItems(batch, game)
    
    if (batchResult.ok && batchResult.items) {
      // Добавляем все предметы в результат (по normalizedName или marketname как ключ)
      batchResult.items.forEach(item => {
        const key = item.normalizedname || item.marketname || item.markethashname
        if (key) {
          result[key.toLowerCase()] = item
        }
      })
    } else {
      errors.push({
        batch: batch,
        error: batchResult.error || 'unknown'
      })
    }
    
    // Пауза между пакетами
    if (i + maxItemsPerRequest < itemNames.length) {
      Utilities.sleep(API_CONFIG.STEAM_WEB_API.RETRY_DELAY_MS)
    }
  }
  
  return {
    ok: errors.length === 0,
    items: result,
    errors: errors.length > 0 ? errors : undefined
  }
}

/**
 * Получает данные о предмете по NameID через SteamWebAPI.ru
 * @param {number|string} nameId - NameID предмета (уникальный идентификатор)
 * @returns {Object} {ok: boolean, item?: Object, error?: string}
 */
function steamWebAPI_fetchItemByNameId(nameId) {
  if (!nameId) {
    return { ok: false, error: 'empty_nameid' }
  }
  
  const baseUrl = `${API_CONFIG.STEAM_WEB_API.BASE_URL}/api/item_by_nameid`
  const url = `${baseUrl}?nameid=${encodeURIComponent(nameId)}`
  
  const options = {
    method: 'GET',
    muteHttpExceptions: true,
    headers: {
      'Accept': 'application/json'
    }
  }
  
  const result = fetchWithRetry_(url, options, {
    maxRetries: API_CONFIG.STEAM_WEB_API.MAX_RETRIES,
    retryDelayMs: API_CONFIG.STEAM_WEB_API.RETRY_DELAY_MS,
    apiName: 'SteamWebAPI (item_by_nameid)'
  })
  
  if (!result.ok) {
    return result
  }
  
  try {
    const responseText = result.response.getContentText()
    const item = JSON.parse(responseText)
    
    if (!item || typeof item !== 'object') {
      console.error('SteamWebAPI: неожиданный формат ответа (item_by_nameid, не объект)')
      return { ok: false, error: 'invalid_format', details: 'Response is not an object' }
    }
    
    return { ok: true, item: item }
  } catch (e) {
    console.error('SteamWebAPI: ошибка парсинга ответа (item_by_nameid):', e)
    return { ok: false, error: 'parse_error', details: e.message }
  }
}

/**
 * Получает данные о предмете по названию через item_by_nameid endpoint
 * API автоматически определяет NameID из базы данных по названию
 * Использует параметр `name` для поиска по Market Hash Name
 * @param {string} itemName - Название предмета
 * @param {string} game - Игра ('dota2', 'cs2', 'rust', 'tf2')
 * @returns {Object} {ok: boolean, item?: Object, error?: string}
 */
function steamWebAPI_fetchItemByNameIdViaName(itemName, game = 'dota2') {
  if (!itemName || itemName.trim().length === 0) {
    return { ok: false, error: 'empty_name' }
  }
  
  const baseUrl = `${API_CONFIG.STEAM_WEB_API.BASE_URL}/api/item_by_nameid`
  const url = `${baseUrl}?game=${encodeURIComponent(game)}&name=${encodeURIComponent(itemName)}`
  
  const options = {
    method: 'GET',
    muteHttpExceptions: true,
    headers: {
      'Accept': 'application/json'
    }
  }
  
  const result = fetchWithRetry_(url, options, {
    maxRetries: API_CONFIG.STEAM_WEB_API.MAX_RETRIES,
    retryDelayMs: API_CONFIG.STEAM_WEB_API.RETRY_DELAY_MS,
    apiName: 'SteamWebAPI (item_by_nameid via name)'
  })
  
  if (!result.ok) {
    return result
  }
  
  try {
    const responseText = result.response.getContentText()
    const data = JSON.parse(responseText)
    
    // API возвращает массив, берем первый элемент
    if (Array.isArray(data) && data.length > 0) {
      return { ok: true, item: data[0] }
    }
    
    return { ok: false, error: 'no_data', details: 'Empty response array' }
  } catch (e) {
    console.error('SteamWebAPI: ошибка парсинга ответа (item_by_nameid via name):', e)
    return { ok: false, error: 'parse_error', details: e.message }
  }
}

/**
 * Тест подключения к SteamWebAPI.ru
 * Показывает результат в UI
 */
function steamWebAPI_testConnection() {
  const ui = SpreadsheetApp.getUi()
  
  try {
    // Тестируем на примере предмета "Infernal Menace"
    const testItemName = 'Infernal Menace'
    
    console.log('SteamWebAPI: начало теста подключения')
    console.log(`SteamWebAPI: тестовый предмет: "${testItemName}"`)
    
    // Используем существующую функцию для корректного формирования запроса
    const result = steamWebAPI_fetchSingleItem(testItemName, 'dota2')
    
    if (!result.ok) {
      let errorMessage = `❌ Ошибка подключения\n\n`
      
      if (result.error === 'item_not_found') {
        errorMessage += `Предмет "${testItemName}" не найден в базе SteamWebAPI.ru.\n\n`
        errorMessage += `API работает, но тестовый предмет отсутствует.\n`
        errorMessage += `Попробуйте другой предмет или проверьте правильность названия.`
      } else if (result.error === 'empty_items') {
        errorMessage += `Пустое название предмета.`
      } else if (result.error && result.error.startsWith('http_')) {
        errorMessage += `HTTP ошибка: ${result.error}\n\n`
        errorMessage += `${result.details || 'Нет дополнительной информации'}`
      } else {
        errorMessage += `Ошибка: ${result.error || 'unknown'}\n\n`
        errorMessage += `${result.details || 'Нет дополнительной информации'}`
      }
      
      console.error('SteamWebAPI test:', errorMessage)
      ui.alert('Тест SteamWebAPI.ru', errorMessage, ui.ButtonSet.OK)
      return result
    }
    
    if (!result.item) {
      const errorMessage = `❌ Предмет не найден\n\n` +
                          `Предмет "${testItemName}" не найден в базе SteamWebAPI.ru.`
      console.warn('SteamWebAPI test: предмет не найден')
      ui.alert('Тест SteamWebAPI.ru', errorMessage, ui.ButtonSet.OK)
      return { ok: false, error: 'item_not_found' }
    }
    
    // Парсим данные предмета
    const parsed = steamWebAPI_parseItemData(result.item)
    
    const successMessage = `✅ Подключение успешно!\n\n` +
                         `Тестовый предмет: "${testItemName}"\n\n` +
                         `Получено данных:\n` +
                         `• Название: ${parsed.name || 'N/A'}\n` +
                         `• Цена: ${parsed.priceLatest ? parsed.priceLatest + ' ₽' : 'N/A'}\n` +
                         `• Продано за 24ч: ${parsed.sold24h || 0}\n` +
                         `• Герой (tag5): ${parsed.tag5 || 'N/A'}\n\n` +
                         `SteamWebAPI.ru работает и готов к использованию!`
    
    console.log(`SteamWebAPI test: успешно, получены данные для "${testItemName}"`)
    ui.alert('Тест SteamWebAPI.ru', successMessage, ui.ButtonSet.OK)
    return { ok: true, message: successMessage, item: parsed }
    
  } catch (e) {
    const errorMessage = `❌ Ошибка: ${e.message}\n\n${e.stack || ''}`
    console.error('SteamWebAPI test: исключение', e)
    ui.alert('Тест SteamWebAPI.ru', errorMessage, ui.ButtonSet.OK)
    return { ok: false, error: 'exception', details: e.message }
  }
}

