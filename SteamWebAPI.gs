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
  const baseUrl = `${API_CONFIG.STEAM_WEB_API.BASE_URL}/api/item`
  const params = []
  params.push(`game=${encodeURIComponent(game)}`)
  
  // Добавляем все названия предметов
  itemNames.forEach(name => {
    params.push(`name[]=${encodeURIComponent(name)}`)
  })
  
  const url = `${baseUrl}?${params.join('&')}`
  
  let attempts = 0
  const maxRetries = API_CONFIG.STEAM_WEB_API.MAX_RETRIES
  
  while (attempts < maxRetries) {
    try {
      const options = {
        method: 'GET',
        muteHttpExceptions: true
      }
      
      const response = UrlFetchApp.fetch(url, options)
      const responseCode = response.getResponseCode()
      const responseText = response.getContentText()
      
      if (responseCode === 429) {
        const delay = API_CONFIG.STEAM_WEB_API.RETRY_DELAY_MS * Math.pow(2, attempts)
        console.warn(`SteamWebAPI: лимит запросов, повтор через ${delay}ms (попытка ${attempts + 1}/${maxRetries})`)
        Utilities.sleep(delay)
        attempts++
        continue
      }
      
      if (responseCode !== 200) {
        console.error(`SteamWebAPI: HTTP ошибка ${responseCode}: ${responseText.substring(0, 200)}`)
        return { ok: false, error: `http_${responseCode}`, details: responseText.substring(0, 200) }
      }
      
      const items = JSON.parse(responseText)
      
      if (!Array.isArray(items)) {
        console.error('SteamWebAPI: неожиданный формат ответа (не массив)')
        return { ok: false, error: 'invalid_format', details: 'Response is not an array' }
      }
      
      return { ok: true, items: items }
      
    } catch (e) {
      attempts++
      if (attempts >= maxRetries) {
        console.error('SteamWebAPI: исключение после всех попыток:', e)
        return { ok: false, error: 'exception', details: e.message }
      }
      
      const delay = API_CONFIG.STEAM_WEB_API.RETRY_DELAY_MS * Math.pow(2, attempts - 1)
      console.warn(`SteamWebAPI: ошибка, повтор через ${delay}ms (попытка ${attempts}/${maxRetries}):`, e.message)
      Utilities.sleep(delay)
    }
  }
  
  return { ok: false, error: 'max_retries_exceeded' }
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

