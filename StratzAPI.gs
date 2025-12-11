/**
 * StratzAPI - Интеграция с Stratz GraphQL API
 * 
 * Функции для получения статистики героев через GraphQL
 * API ключ хранится в PropertiesService
 */

/**
 * Получает конфигурацию Stratz API
 * @returns {Object} {apiKey} или null если не настроено
 */
function stratz_getConfig() {
  const props = PropertiesService.getScriptProperties()
  const apiKey = props.getProperty('stratz_api_key')
  
  if (!apiKey) {
    return null
  }
  
  return { apiKey }
}

/**
 * Сохраняет конфигурацию Stratz API
 * @param {string} apiKey - API ключ от Stratz
 */
function stratz_setConfig(apiKey) {
  if (!apiKey || apiKey.trim().length === 0) {
    throw new Error('Stratz: apiKey обязателен')
  }
  
  const props = PropertiesService.getScriptProperties()
  props.setProperty('stratz_api_key', apiKey.trim())
  
  console.log('Stratz: конфигурация сохранена')
  return { ok: true }
}

/**
 * Выполняет GraphQL запрос к Stratz API
 * @param {string} query - GraphQL запрос
 * @param {Object} variables - Переменные запроса (опционально)
 * @returns {Object} {ok: boolean, data?: Object, error?: string}
 */
function stratz_executeGraphQL(query, variables = {}) {
  const config = stratz_getConfig()
  
  if (!config) {
    console.error('Stratz: конфигурация не настроена')
    return { ok: false, error: 'not_configured' }
  }
  
  if (!query || query.trim().length === 0) {
    console.error('Stratz: пустой запрос')
    return { ok: false, error: 'empty_query' }
  }
  
  const url = API_CONFIG.STRATZ.BASE_URL
  
  let attempts = 0
  const maxRetries = API_CONFIG.STRATZ.MAX_RETRIES
  
  while (attempts < maxRetries) {
    try {
      const payload = {
        query: query,
        variables: variables
      }
      
      const options = {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${config.apiKey}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
      
      const response = UrlFetchApp.fetch(url, options)
      const responseCode = response.getResponseCode()
      const responseText = response.getContentText()
      
      // Проверка лимитов (429 Too Many Requests)
      if (responseCode === 429) {
        const delay = API_CONFIG.STRATZ.RETRY_DELAY_MS * Math.pow(2, attempts)
        console.warn(`Stratz: лимит запросов, повтор через ${delay}ms (попытка ${attempts + 1}/${maxRetries})`)
        Utilities.sleep(delay)
        attempts++
        continue
      }
      
      if (responseCode !== 200) {
        console.error(`Stratz: HTTP ошибка ${responseCode}: ${responseText}`)
        
        // Детальная обработка 403 (Forbidden)
        if (responseCode === 403) {
          let errorMessage = 'HTTP 403 Forbidden - Доступ запрещен'
          try {
            const errorData = JSON.parse(responseText)
            if (errorData.message) {
              errorMessage += `: ${errorData.message}`
            }
          } catch (e) {
            // Если не JSON, используем текст как есть
            if (responseText && responseText.length < 200) {
              errorMessage += `: ${responseText}`
            }
          }
          return { 
            ok: false, 
            error: 'http_403', 
            details: responseText,
            message: errorMessage
          }
        }
        
        return { ok: false, error: `http_${responseCode}`, details: responseText }
      }
      
      const result = JSON.parse(responseText)
      
      // Проверка ошибок GraphQL
      if (result.errors && result.errors.length > 0) {
        const errorMessages = result.errors.map(e => e.message).join('; ')
        console.error(`Stratz: GraphQL ошибки: ${errorMessages}`)
        return { ok: false, error: 'graphql_error', details: errorMessages }
      }
      
      return { ok: true, data: result.data }
      
    } catch (e) {
      attempts++
      if (attempts >= maxRetries) {
        console.error('Stratz: исключение после всех попыток:', e)
        return { ok: false, error: 'exception', details: e.message }
      }
      
      const delay = API_CONFIG.STRATZ.RETRY_DELAY_MS * Math.pow(2, attempts - 1)
      console.warn(`Stratz: ошибка, повтор через ${delay}ms (попытка ${attempts}/${maxRetries}):`, e.message)
      Utilities.sleep(delay)
    }
  }
  
  return { ok: false, error: 'max_retries_exceeded' }
}

/**
 * Получает статистику всех героев
 * @param {string} rankCategory - 'high' для High Rank (Divine/Immortal) или 'all' для всех рангов
 * @returns {Object} {ok: boolean, heroStats?: Array, error?: string}
 */
function stratz_fetchHeroStats(rankCategory = 'all') {
  // GraphQL запрос для получения статистики героев
  // Примечание: точный запрос может потребовать корректировки после изучения документации
  const query = `
    query GetHeroStats($rank: Rank) {
      constants {
        heroes {
          id
          displayName
          shortName
        }
      }
      heroStats {
        heroId
        pickRate
        winRate
        banRate
        matchCount
      }
    }
  `
  
  const variables = rankCategory === 'high' ? { rank: 'IMMORTAL' } : {}
  
  const result = stratz_executeGraphQL(query, variables)
  
  if (!result.ok) {
    return result
  }
  
  // Обработка данных
  const heroes = result.data?.constants?.heroes || []
  const heroStats = result.data?.heroStats || []
  
  // Объединение данных героев со статистикой
  const combinedStats = heroStats.map(stat => {
    const hero = heroes.find(h => h.id === stat.heroId)
    return {
      heroId: stat.heroId,
      heroName: hero?.displayName || hero?.shortName || `Hero ${stat.heroId}`,
      pickRate: stat.pickRate || 0,
      winRate: stat.winRate || 0,
      banRate: stat.banRate || 0,
      contestRate: (stat.pickRate || 0) + (stat.banRate || 0),
      matchCount: stat.matchCount || 0
    }
  })
  
  return { ok: true, heroStats: combinedStats }
}

/**
 * Получает статистику всех героев для обеих категорий (High Rank + All Ranks)
 * @returns {Object} {ok: boolean, highRank?: Array, allRanks?: Array, error?: string}
 */
function stratz_fetchAllHeroStats() {
  const highRankResult = stratz_fetchHeroStats('high')
  Utilities.sleep(500) // Пауза между запросами
  
  const allRanksResult = stratz_fetchHeroStats('all')
  
  if (!highRankResult.ok && !allRanksResult.ok) {
    return { ok: false, error: 'both_requests_failed' }
  }
  
  return {
    ok: true,
    highRank: highRankResult.ok ? highRankResult.heroStats : null,
    allRanks: allRanksResult.ok ? allRanksResult.heroStats : null,
    errors: {
      highRank: highRankResult.ok ? null : highRankResult.error,
      allRanks: allRanksResult.ok ? null : allRanksResult.error
    }
  }
}

/**
 * Тест подключения к Stratz API
 * @returns {Object} {ok: boolean, message?: string, error?: string}
 */
function stratz_testConnection() {
  const config = stratz_getConfig()
  
  if (!config) {
    return { ok: false, error: 'not_configured', message: 'API ключ не настроен' }
  }
  
  // Простой тестовый запрос (получение списка героев)
  const query = `
    query TestConnection {
      constants {
        heroes {
          id
          displayName
        }
      }
    }
  `
  
  const result = stratz_executeGraphQL(query)
  
  if (result.ok && result.data?.constants?.heroes) {
    const heroCount = result.data.constants.heroes.length
    return { 
      ok: true, 
      message: `✅ Подключение успешно! Получено ${heroCount} героев.` 
    }
  }
  
  // Улучшенное сообщение об ошибке
  let errorMessage = `❌ Ошибка: ${result.error || 'Неизвестная ошибка'}`
  
  if (result.error === 'http_403') {
    errorMessage = `❌ Ошибка HTTP 403 (Доступ запрещен)\n\n` +
      `Возможные причины:\n` +
      `1. Неправильный API ключ\n` +
      `2. API ключ не активирован на сайте Stratz\n` +
      `3. API ключ не имеет нужных прав доступа\n\n` +
      `Проверьте:\n` +
      `- Правильность введенного ключа\n` +
      `- Статус ключа на https://stratz.com/api\n` +
      `- Что ключ активен и имеет доступ к GraphQL API`
    
    if (result.details) {
      try {
        const errorData = JSON.parse(result.details)
        if (errorData.message) {
          errorMessage += `\n\nДетали от API: ${errorData.message}`
        }
      } catch (e) {
        // Не JSON, оставляем как есть
      }
    }
  } else if (result.error === 'graphql_error' && result.details) {
    errorMessage += `\n\nДетали: ${result.details}`
  } else if (result.details && result.details.length < 300) {
    errorMessage += `\n\nДетали: ${result.details}`
  }
  
  return { 
    ok: false, 
    error: result.error || 'unknown',
    details: result.details,
    message: errorMessage
  }
}

/**
 * Настройка Stratz API через диалог
 */
function stratz_setup() {
  const ui = SpreadsheetApp.getUi()
  
  // Проверяем текущую конфигурацию
  const currentConfig = stratz_getConfig()
  let promptText = 'Настройка Stratz API\n\n'
  
  if (currentConfig) {
    promptText += 'Текущий API ключ настроен.\n\n'
  }
  
  promptText += 'Введите API ключ от Stratz:'
  
  const response = ui.prompt(
    'Настройка Stratz API',
    promptText,
    ui.ButtonSet.OK_CANCEL
  )
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return
  }
  
  const apiKey = response.getResponseText().trim()
  
  if (!apiKey) {
    ui.alert('Ошибка', 'API ключ не может быть пустым', ui.ButtonSet.OK)
    return
  }
  
  try {
    stratz_setConfig(apiKey)
    
    // Тест подключения
    const testResult = stratz_testConnection()
    
    if (testResult.ok) {
      ui.alert('✅ Конфигурация сохранена!\n\n' + testResult.message + '\n\nИспользуйте "Тест Stratz API" для проверки.')
    } else {
      ui.alert('⚠️ Конфигурация сохранена, но тест подключения не прошел:\n\n' + (testResult.message || testResult.error) + '\n\nПроверьте правильность API ключа.')
    }
  } catch (e) {
    ui.alert('Ошибка', 'Не удалось сохранить конфигурацию: ' + e.message, ui.ButtonSet.OK)
  }
}

