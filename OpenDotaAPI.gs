/**
 * OpenDotaAPI - Интеграция с OpenDota API
 * 
 * Основной источник статистики героев (пикрейт, винрейт, банрейт)
 * OpenDota API бесплатный и не требует API ключа (но есть лимиты)
 * Документация: https://docs.opendota.com/
 */

/**
 * Получает статистику героев через OpenDota API
 * @param {string} rankTier - Ранг (1-8, где 8 = Immortal). null = все ранги
 * @returns {Object} {ok: boolean, heroStats?: Array, error?: string}
 */
function openDota_fetchHeroStats(rankTier = null) {
  // OpenDota API endpoint для статистики героев
  // https://docs.opendota.com/#tag/heroes%2Fpaths%2F~1heroes%2Fget
  let url = 'https://api.opendota.com/api/heroStats'
  
  // Если указан ранг, можно использовать фильтры (но в базовом API их нет)
  // Для фильтрации по рангу нужно использовать другой endpoint или фильтровать на стороне клиента
  
  const options = {
    method: 'GET',
    muteHttpExceptions: true,
    headers: {
      'Accept': 'application/json'
    }
  }
  
  const result = fetchWithRetry_(url, options, {
    maxRetries: API_CONFIG.OPENDOTA.MAX_RETRIES,
    retryDelayMs: API_CONFIG.OPENDOTA.RETRY_DELAY_MS,
    apiName: 'OpenDota'
  })
  
  if (!result.ok) {
    return result
  }
  
  try {
    const responseText = result.response.getContentText()
    const heroes = JSON.parse(responseText)
    
    if (!Array.isArray(heroes)) {
      console.error('OpenDota: неожиданный формат ответа (не массив)')
      return { ok: false, error: 'invalid_format', details: 'Response is not an array' }
    }
    
    // Преобразуем данные OpenDota в формат, совместимый с нашим кодом
    // OpenDota предоставляет статистику по рангам:
    // 1 = Herald, 2 = Guardian, 3 = Crusader, 4 = Archon, 5 = Legend
    // 6 = Ancient, 7 = Divine, 8 = Immortal
    const heroStats = heroes.map(hero => {
      // Вычисляем статистику для "High Rank" (Ancient + Divine + Immortal = 6+7+8)
      const highRankPick = (hero['6_pick'] || 0) + (hero['7_pick'] || 0) + (hero['8_pick'] || 0)
      const highRankWin = (hero['6_win'] || 0) + (hero['7_win'] || 0) + (hero['8_win'] || 0)
      const highRankWinRate = highRankPick > 0 ? (highRankWin / highRankPick) * 100 : 0
      
      // Вычисляем статистику для "All Ranks" (все ранги 1-8)
      const allRanksPick = (hero['1_pick'] || 0) + (hero['2_pick'] || 0) + (hero['3_pick'] || 0) +
                          (hero['4_pick'] || 0) + (hero['5_pick'] || 0) + highRankPick
      const allRanksWin = (hero['1_win'] || 0) + (hero['2_win'] || 0) + (hero['3_win'] || 0) +
                         (hero['4_win'] || 0) + (hero['5_win'] || 0) + highRankWin
      const allRanksWinRate = allRanksPick > 0 ? (allRanksWin / allRanksPick) * 100 : 0
      
      // Про-статистика (pro_pick, pro_ban)
      const proPick = hero.pro_pick || 0
      const proBan = hero.pro_ban || 0
      const proBanRate = proPick > 0 ? (proBan / proPick) * 100 : 0
      const proContestRate = proPick + proBan // Контест рейт для про-матчей
      
      // Pick rate - это абсолютное количество пиков (для нормализации потребуется общая сумма всех пиков)
      // Для текущей реализации используем абсолютные значения
      // Контест rate = pick + ban
      
      return {
        heroId: hero.id,
        heroName: hero.localized_name || hero.name,
        // High Rank статистика
        highRank: {
          pickRate: highRankPick, // Абсолютное значение (пиков)
          winRate: highRankWinRate, // Процент побед
          banRate: proBanRate, // Процент баннов (из про-матчей)
          contestRate: highRankPick + proBan, // Абсолютное значение (пики + баны)
          matchCount: highRankPick,
          // Про-статистика
          proPick: proPick,
          proBan: proBan,
          proContestRate: proContestRate
        },
        // All Ranks статистика
        allRanks: {
          pickRate: allRanksPick, // Абсолютное значение (пиков)
          winRate: allRanksWinRate, // Процент побед
          banRate: proBanRate, // Процент баннов (из про-матчей)
          contestRate: allRanksPick + proBan, // Абсолютное значение (пики + баны)
          matchCount: allRanksPick,
          // Про-статистика
          proPick: proPick,
          proBan: proBan,
          proContestRate: proContestRate
        }
      }
    })
    
    return { ok: true, heroStats: heroStats }
  } catch (e) {
    console.error('OpenDota: ошибка парсинга ответа:', e)
    return { ok: false, error: 'parse_error', details: e.message }
  }
}

/**
 * Получает статистику всех героев для обеих категорий (High Rank + All Ranks)
 * Аналогично stratz_fetchAllHeroStats() для совместимости
 * @returns {Object} {ok: boolean, highRank?: Array, allRanks?: Array, error?: string}
 */
function openDota_fetchAllHeroStats() {
  const result = openDota_fetchHeroStats()
  
  if (!result.ok || !result.heroStats) {
    return result
  }
  
  // Разделяем статистику на High Rank и All Ranks
  const highRank = []
  const allRanks = []
  
  result.heroStats.forEach(hero => {
    if (hero.highRank) {
      highRank.push({
        heroId: hero.heroId,
        heroName: hero.heroName,
        pickRate: hero.highRank.pickRate,
        winRate: hero.highRank.winRate,
        banRate: hero.highRank.banRate,
        contestRate: hero.highRank.contestRate,
        matchCount: hero.highRank.matchCount,
        // Про-статистика
        proPick: hero.highRank.proPick || 0,
        proBan: hero.highRank.proBan || 0,
        proContestRate: hero.highRank.proContestRate || 0
      })
    }
    
    if (hero.allRanks) {
      allRanks.push({
        heroId: hero.heroId,
        heroName: hero.heroName,
        pickRate: hero.allRanks.pickRate,
        winRate: hero.allRanks.winRate,
        banRate: hero.allRanks.banRate,
        contestRate: hero.allRanks.contestRate,
        matchCount: hero.allRanks.matchCount,
        // Про-статистика
        proPick: hero.allRanks.proPick || 0,
        proBan: hero.allRanks.proBan || 0,
        proContestRate: hero.allRanks.proContestRate || 0
      })
    }
  })
  
  return {
    ok: true,
    highRank: highRank,
    allRanks: allRanks
  }
}

/**
 * Тест подключения к OpenDota API
 * Показывает результат в UI
 */
function openDota_testConnection() {
  const ui = SpreadsheetApp.getUi()
  
  try {
    const url = 'https://api.opendota.com/api/heroStats'
    
    console.log('OpenDota: начало теста подключения')
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      muteHttpExceptions: true,
      headers: {
        'Accept': 'application/json'
      }
    })
    
    const responseCode = response.getResponseCode()
    const responseText = response.getContentText()
    
    console.log(`OpenDota: получен ответ ${responseCode}, длина: ${responseText.length}`)
    
    if (responseCode !== HTTP_STATUS.OK) {
      const errorPreview = responseText.substring(0, LIMITS.ERROR_MESSAGE_MAX_LENGTH)
      const errorMessage = `❌ Ошибка HTTP ${responseCode}\n\n${errorPreview}`
      console.error('OpenDota test:', errorMessage)
      ui.alert('Тест OpenDota API', errorMessage, ui.ButtonSet.OK)
      return { ok: false, error: `http_${responseCode}` }
    }
    
    const heroes = JSON.parse(responseText)
    
    if (Array.isArray(heroes) && heroes.length > 0) {
      const successMessage = `✅ Подключение успешно!\n\n` +
                           `Получено героев: ${heroes.length}\n\n` +
                           `Пример героя:\n${heroes[0].localized_name || heroes[0].name} (ID: ${heroes[0].id})\n\n` +
                           `OpenDota API работает и готов к использованию!`
      
      console.log(`OpenDota test: успешно, получено ${heroes.length} героев`)
      ui.alert('Тест OpenDota API', successMessage, ui.ButtonSet.OK)
      return { ok: true, message: successMessage }
    }
    
    const errorMessage = '❌ Неожиданный формат ответа\n\nОтвет не является массивом героев'
    console.error('OpenDota test:', errorMessage)
    ui.alert('Тест OpenDota API', errorMessage, ui.ButtonSet.OK)
    return { ok: false, error: 'invalid_format' }
    
  } catch (e) {
    const errorMessage = `❌ Ошибка: ${e.message}\n\n${e.stack || ''}`
    console.error('OpenDota test: исключение', e)
    ui.alert('Тест OpenDota API', errorMessage, ui.ButtonSet.OK)
    return { ok: false, error: 'exception', details: e.message }
  }
}

