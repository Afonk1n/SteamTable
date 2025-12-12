/**
 * HeroStats - Управление листом статистики героев
 * 
 * Структура:
 * - Фиксированные колонки (A-C): Hero ID, Hero Name, Rank Category
 * - Динамические колонки (начиная с D): даты/время, JSON с данными статистики
 */

const HERO_STATS_CONFIG = {
  COLUMNS: HERO_STATS_COLUMNS
}

/**
 * Форматирование листа HeroStats
 */
function heroStats_formatTable() {
  const sheet = getOrCreateHeroStatsSheet_()
  const lastRow = sheet.getLastRow()

  // Заголовки фиксированных колонок
  const headers = HEADERS.HERO_STATS
  if (!headers || !Array.isArray(headers) || headers.length === 0) {
    console.error('HeroStats: HEADERS.HERO_STATS не определен или пуст')
    SpreadsheetApp.getUi().alert('Ошибка: HEADERS.HERO_STATS не определен в Constants.gs')
    return
  }
  sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers])

  formatHeaderRange_(sheet.getRange(HEADER_ROW, 1, 1, headers.length))

  sheet.setRowHeight(HEADER_ROW, HEADER_HEIGHT)
  if (lastRow > 1) {
    sheet.setRowHeights(DATA_START_ROW, lastRow - 1, ROW_HEIGHT)
  }

  // Ширины колонок
  sheet.setColumnWidth(1, 80)   // A - Hero ID
  sheet.setColumnWidth(2, 150)  // B - Hero Name
  sheet.setColumnWidth(3, 120)  // C - Rank Category

  if (lastRow > 1) {
    // Форматирование данных
    sheet.getRange(DATA_START_ROW, 1, lastRow - 1, 3)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setWrap(false)
    sheet.getRange(`${HERO_STATS_COLUMNS.HERO_NAME}${DATA_START_ROW}:${HERO_STATS_COLUMNS.HERO_NAME}${lastRow}`)
      .setHorizontalAlignment('left')
    
    // Формат числа для Hero ID
    sheet.getRange(`${HERO_STATS_COLUMNS.HERO_ID}${DATA_START_ROW}:${HERO_STATS_COLUMNS.HERO_ID}${lastRow}`)
      .setNumberFormat(NUMBER_FORMATS.INTEGER)
  }

  sheet.setFrozenRows(HEADER_ROW)
  sheet.setFrozenColumns(3) // Замораживаем первые 3 колонки

  // Форматируем все существующие динамические колонки
  heroStats_formatAllDataColumns_(sheet)

  console.log('HeroStats: форматирование завершено')
}

/**
 * Форматирует все динамические колонки с данными (начиная с D)
 * @param {Sheet} sheet - Лист HeroStats
 */
function heroStats_formatAllDataColumns_(sheet) {
  const firstDataCol = HERO_STATS_COLUMNS.FIRST_DATA_COL
  const lastCol = sheet.getLastColumn()
  
  if (lastCol < firstDataCol) {
    return
  }
  
  const lastRow = sheet.getLastRow()
  
  for (let col = firstDataCol; col <= lastCol; col++) {
    // Форматируем заголовок
    const header = sheet.getRange(HEADER_ROW, col)
    header.setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBackground(COLORS.BACKGROUND)
      .setFontWeight('bold')
      .setWrap(true)
    
    sheet.setColumnWidth(col, 150) // Ширина для JSON данных
    
    // Форматируем данные (JSON строки)
    if (lastRow > HEADER_ROW) {
      sheet.getRange(DATA_START_ROW, col, lastRow - HEADER_ROW, 1)
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('left')
        .setWrap(true)
    }
  }
}

/**
 * Добавляет новую колонку с данными статистики для указанного времени
 * @param {Date} dateTime - Дата и время обновления
 * @returns {number} Индекс созданной колонки
 */
function heroStats_addStatsColumn(dateTime) {
  const sheet = getOrCreateHeroStatsSheet_()
  const firstDataCol = HERO_STATS_COLUMNS.FIRST_DATA_COL
  const lastCol = sheet.getLastColumn()
  
  // Формируем заголовок колонки: "DD.MM.YY HH:MM"
  const headerDisplay = Utilities.formatDate(dateTime, Session.getScriptTimeZone(), 'dd.MM.yy HH:mm')
  
  // Проверяем, не существует ли уже колонка с таким заголовком
  if (lastCol >= firstDataCol) {
    const width = lastCol - firstDataCol + 1
    const dateRow = sheet.getRange(HEADER_ROW, firstDataCol, 1, width).getDisplayValues()[0]
    
    for (let i = 0; i < dateRow.length; i++) {
      const header = String(dateRow[i] || '').trim()
      if (header === headerDisplay) {
        const col = firstDataCol + i
        heroStats_formatDataColumn_(sheet, col)
        return col
      }
    }
  }
  
  // Создаём новую колонку справа от последней
  const newCol = Math.max(lastCol + 1, firstDataCol)
  sheet.getRange(HEADER_ROW, newCol).setValue(headerDisplay)
  heroStats_formatDataColumn_(sheet, newCol)
  
  return newCol
}

/**
 * Форматирует одну динамическую колонку с данными
 * @param {Sheet} sheet - Лист HeroStats
 * @param {number} col - Номер колонки
 */
function heroStats_formatDataColumn_(sheet, col) {
  const header = sheet.getRange(HEADER_ROW, col)
  header.setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(COLORS.BACKGROUND)
    .setFontWeight('bold')
    .setWrap(true)
  
  sheet.setColumnWidth(col, 150)
  
  const lastRow = sheet.getLastRow()
  if (lastRow > HEADER_ROW) {
    sheet.getRange(DATA_START_ROW, col, lastRow - HEADER_ROW, 1)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('left')
      .setWrap(true)
  }
}

/**
 * Получает или создает строку для героя и категории ранга
 * @param {number} heroId - ID героя
 * @param {string} heroName - Имя героя
 * @param {string} rankCategory - 'High Rank' или 'All Ranks'
 * @returns {number} Номер строки
 */
function heroStats_getOrCreateRow(heroId, heroName, rankCategory) {
  const sheet = getOrCreateHeroStatsSheet_()
  const lastRow = sheet.getLastRow()
  
  // Ищем существующую строку
  if (lastRow >= DATA_START_ROW) {
    const heroIds = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, 1).getValues()
    const rankCategories = sheet.getRange(DATA_START_ROW, 3, lastRow - HEADER_ROW, 1).getValues()
    
    for (let i = 0; i < heroIds.length; i++) {
      const rowId = Number(heroIds[i][0])
      const rowRank = String(rankCategories[i][0] || '').trim()
      
      if (rowId === heroId && rowRank === rankCategory) {
        return DATA_START_ROW + i
      }
    }
  }
  
  // Создаём новую строку
  const newRow = Math.max(lastRow + 1, DATA_START_ROW)
  sheet.getRange(newRow, 1).setValue(heroId)
  sheet.getRange(newRow, 2).setValue(heroName)
  sheet.getRange(newRow, 3).setValue(rankCategory)
  
  // Форматируем новую строку
  sheet.getRange(newRow, 1, 1, 3)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
  sheet.getRange(newRow, 2).setHorizontalAlignment('left')
  sheet.setRowHeight(newRow, ROW_HEIGHT)
  
  return newRow
}

/**
 * Сохраняет статистику героя в указанную колонку
 * @param {number} row - Номер строки
 * @param {number} col - Номер колонки
 * @param {Object} statsData - Данные статистики {pickRate, winRate, banRate, contestRate, matchCount}
 */
function heroStats_setStatsData(row, col, statsData) {
  const sheet = getOrCreateHeroStatsSheet_()
  
  // Сохраняем как JSON строку
  const jsonString = JSON.stringify(statsData)
  sheet.getRange(row, col).setValue(jsonString)
}

/**
 * Получает статистику героя из указанной колонки
 * @param {number} row - Номер строки
 * @param {number} col - Номер колонки
 * @returns {Object|null} Данные статистики или null
 */
function heroStats_getStatsData(row, col) {
  const sheet = getHeroStatsSheet_()
  if (!sheet) return null
  
  const jsonString = sheet.getRange(row, col).getValue()
  if (!jsonString || typeof jsonString !== 'string') {
    return null
  }
  
  try {
    return JSON.parse(jsonString)
  } catch (e) {
    console.error('HeroStats: ошибка парсинга JSON:', e)
    return null
  }
}

/**
 * Получает статистику героя для последнего обновления
 * @param {number} heroId - ID героя
 * @param {string} rankCategory - 'High Rank' или 'All Ranks'
 * @returns {Object|null} Данные статистики или null
 */
function heroStats_getLatestStats(heroId, rankCategory) {
  const sheet = getHeroStatsSheet_()
  if (!sheet) return null
  
  const row = heroStats_findRow(heroId, rankCategory)
  if (row === -1) return null
  
  const lastCol = sheet.getLastColumn()
  const firstDataCol = HERO_STATS_COLUMNS.FIRST_DATA_COL
  
  if (lastCol < firstDataCol) return null
  
  // Ищем последнюю непустую колонку справа налево
  for (let col = lastCol; col >= firstDataCol; col--) {
    const data = heroStats_getStatsData(row, col)
    if (data) {
      return data
    }
  }
  
  return null
}

/**
 * Получает статистику героя за указанную дату/время
 * @param {number} heroId - ID героя
 * @param {string} rankCategory - 'High Rank' или 'All Ranks'
 * @param {Date} dateTime - Дата и время
 * @returns {Object|null} Данные статистики или null
 */
function heroStats_getStatsForDateTime(heroId, rankCategory, dateTime) {
  const sheet = getHeroStatsSheet_()
  if (!sheet) return null
  
  const row = heroStats_findRow(heroId, rankCategory)
  if (row === -1) return null
  
  const headerDisplay = Utilities.formatDate(dateTime, Session.getScriptTimeZone(), 'dd.MM.yy HH:mm')
  const firstDataCol = HERO_STATS_COLUMNS.FIRST_DATA_COL
  const lastCol = sheet.getLastColumn()
  
  if (lastCol < firstDataCol) return null
  
  // Ищем колонку с нужным заголовком
  const dateRow = sheet.getRange(HEADER_ROW, firstDataCol, 1, lastCol - firstDataCol + 1).getDisplayValues()[0]
  
  for (let i = 0; i < dateRow.length; i++) {
    const header = String(dateRow[i] || '').trim()
    if (header === headerDisplay) {
      const col = firstDataCol + i
      return heroStats_getStatsData(row, col)
    }
  }
  
  return null
}

/**
 * Находит строку для героя и категории ранга
 * @param {number} heroId - ID героя
 * @param {string} rankCategory - 'High Rank' или 'All Ranks'
 * @returns {number} Номер строки или -1 если не найдена
 */
function heroStats_findRow(heroId, rankCategory) {
  const sheet = getHeroStatsSheet_()
  if (!sheet) return -1
  
  const lastRow = sheet.getLastRow()
  if (lastRow < DATA_START_ROW) return -1
  
  const heroIds = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, 1).getValues()
  const rankCategories = sheet.getRange(DATA_START_ROW, 3, lastRow - HEADER_ROW, 1).getValues()
  
  for (let i = 0; i < heroIds.length; i++) {
    const rowId = Number(heroIds[i][0])
    const rowRank = String(rankCategories[i][0] || '').trim()
    
    if (rowId === heroId && rowRank === rankCategory) {
      return DATA_START_ROW + i
    }
  }
  
  return -1
}

/**
 * Получает статистику героя 7 дней назад
 * @param {number} heroId - ID героя
 * @param {string} rankCategory - 'High Rank' или 'All Ranks'
 * @returns {Object|null} Данные статистики или null
 */
function heroStats_getStats7DaysAgo(heroId, rankCategory) {
  const now = new Date()
  const sevenDaysAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000)
  
  // Ищем ближайшую колонку к этой дате
  const sheet = getHeroStatsSheet_()
  if (!sheet) return null
  
  const row = heroStats_findRow(heroId, rankCategory)
  if (row === -1) return null
  
  const firstDataCol = HERO_STATS_COLUMNS.FIRST_DATA_COL
  const lastCol = sheet.getLastColumn()
  
  if (lastCol < firstDataCol) return null
  
  // Парсим все заголовки и ищем ближайший к нужной дате
  const dateRow = sheet.getRange(HEADER_ROW, firstDataCol, 1, lastCol - firstDataCol + 1).getDisplayValues()[0]
  let closestCol = -1
  let closestDiff = Infinity
  
  for (let i = 0; i < dateRow.length; i++) {
    const header = String(dateRow[i] || '').trim()
    if (!header) continue
    
    try {
      // Парсим дату из формата "DD.MM.YY HH:MM"
      const parts = header.split(' ')
      if (parts.length !== 2) continue
      
      const dateStr = parts[0]
      const [day, month, year] = dateStr.split('.')
      const parsedDate = new Date(2000 + parseInt(year), parseInt(month) - 1, parseInt(day))
      
      const diff = Math.abs(parsedDate.getTime() - sevenDaysAgo.getTime())
      if (diff < closestDiff) {
        closestDiff = diff
        closestCol = firstDataCol + i
      }
    } catch (e) {
      // Пропускаем некорректные заголовки
      continue
    }
  }
  
  if (closestCol === -1) return null
  
  // Проверяем, что разница не больше 8 дней (допустимая погрешность)
  if (closestDiff > 8 * 24 * 60 * 60 * 1000) {
    return null
  }
  
  return heroStats_getStatsData(row, closestCol)
}

/**
 * Получает статистику героя 24 часа назад
 * @param {number} heroId - ID героя
 * @param {string} rankCategory - 'High Rank' или 'All Ranks'
 * @returns {Object|null} Данные статистики или null
 */
function heroStats_getStats24HoursAgo(heroId, rankCategory) {
  const now = new Date()
  const twentyFourHoursAgo = new Date(now.getTime() - 24 * 60 * 60 * 1000)
  
  // Ищем ближайшую колонку к этой дате
  const sheet = getHeroStatsSheet_()
  if (!sheet) return null
  
  const row = heroStats_findRow(heroId, rankCategory)
  if (row === -1) return null
  
  const firstDataCol = HERO_STATS_COLUMNS.FIRST_DATA_COL
  const lastCol = sheet.getLastColumn()
  
  if (lastCol < firstDataCol) return null
  
  // Парсим все заголовки и ищем ближайший к нужной дате
  const dateRow = sheet.getRange(HEADER_ROW, firstDataCol, 1, lastCol - firstDataCol + 1).getDisplayValues()[0]
  let closestCol = -1
  let closestDiff = Infinity
  
  for (let i = 0; i < dateRow.length; i++) {
    const header = String(dateRow[i] || '').trim()
    if (!header) continue
    
    try {
      // Парсим дату и время из формата "DD.MM.YY HH:MM"
      const parts = header.split(' ')
      if (parts.length !== 2) continue
      
      const dateStr = parts[0]
      const timeStr = parts[1]
      const [day, month, year] = dateStr.split('.')
      const [hour, minute] = timeStr.split(':')
      const parsedDate = new Date(2000 + parseInt(year), parseInt(month) - 1, parseInt(day), parseInt(hour), parseInt(minute))
      
      const diff = Math.abs(parsedDate.getTime() - twentyFourHoursAgo.getTime())
      if (diff < closestDiff) {
        closestDiff = diff
        closestCol = firstDataCol + i
      }
    } catch (e) {
      // Пропускаем некорректные заголовки
      continue
    }
  }
  
  if (closestCol === -1) return null
  
  // Проверяем, что разница не больше 36 часов (допустимая погрешность для 24h)
  if (closestDiff > 36 * 60 * 60 * 1000) {
    return null
  }
  
  return heroStats_getStatsData(row, closestCol)
}

/**
 * Обновляет статистику всех героев
 * Получает данные через OpenDota API и сохраняет в лист
 */
function heroStats_updateAllStats() {
  console.log('HeroStats: начало обновления статистики')
  
  try {
    // Получаем статистику для обеих категорий через OpenDota API
    const result = openDota_fetchAllHeroStats()
    
    if (!result.ok) {
      console.error('HeroStats: ошибка получения данных:', result.error)
      logAutoAction_('HeroStats', 'Обновление статистики', `Ошибка: ${result.error}`)
      return
    }
    
    const sheet = getOrCreateHeroStatsSheet_()
    const now = new Date()
    const dataCol = heroStats_addStatsColumn(now)
    
    let updatedCount = 0
    let errorCount = 0
    
    // ОПТИМИЗАЦИЯ: Собираем все обновления для batch записи
    const batchUpdates = [] // {row, statsData}
    
    // Подготавливаем статистику High Rank
    if (result.highRank && Array.isArray(result.highRank)) {
      result.highRank.forEach(heroStat => {
        try {
          const row = heroStats_getOrCreateRow(
            heroStat.heroId,
            heroStat.heroName,
            'High Rank'
          )
          
          // Получаем статистику 7 дней назад и 24 часа назад для расчета изменений
          const stats7DaysAgo = heroStats_getStats7DaysAgo(heroStat.heroId, 'High Rank')
          const stats24HoursAgo = heroStats_getStats24HoursAgo(heroStat.heroId, 'High Rank')
          
          // Рассчитываем pickRateChange7d (Immortal за неделю)
          // Используем pickRatePercent вместо contestRatePercent (contestRate был фейком)
          let pickRateChange7d = 0
          if (stats7DaysAgo && stats7DaysAgo.pickRatePercent !== undefined && stats7DaysAgo.pickRatePercent > 0) {
            const currentPickRatePercent = heroStat.pickRatePercent || 0
            pickRateChange7d = ((currentPickRatePercent - stats7DaysAgo.pickRatePercent) / stats7DaysAgo.pickRatePercent) * 100
            
            // ВАЛИДАЦИЯ: Ограничиваем изменения разумными пределами (-1000% до +1000%)
            const MAX_CHANGE_PERCENT = 1000
            const MIN_CHANGE_PERCENT = -1000
            if (pickRateChange7d > MAX_CHANGE_PERCENT || pickRateChange7d < MIN_CHANGE_PERCENT) {
              console.warn(`HeroStats: аномальное изменение pickRate для героя ${heroStat.heroId} (High Rank): ${pickRateChange7d.toFixed(2)}% (ограничено)`)
              pickRateChange7d = pickRateChange7d > 0 ? MAX_CHANGE_PERCENT : MIN_CHANGE_PERCENT
            }
          }
          
          // Рассчитываем pickRateChange24h (Immortal за 24 часа) - для Мета сигнала
          let pickRateChange24h = 0
          if (stats24HoursAgo && stats24HoursAgo.pickRatePercent !== undefined && stats24HoursAgo.pickRatePercent > 0) {
            const currentPickRatePercent = heroStat.pickRatePercent || 0
            pickRateChange24h = ((currentPickRatePercent - stats24HoursAgo.pickRatePercent) / stats24HoursAgo.pickRatePercent) * 100
            
            // ВАЛИДАЦИЯ: Ограничиваем изменения разумными пределами
            const MAX_CHANGE_PERCENT = 1000
            const MIN_CHANGE_PERCENT = -1000
            if (pickRateChange24h > MAX_CHANGE_PERCENT || pickRateChange24h < MIN_CHANGE_PERCENT) {
              console.warn(`HeroStats: аномальное изменение pickRate (24h) для героя ${heroStat.heroId} (High Rank): ${pickRateChange24h.toFixed(2)}% (ограничено)`)
              pickRateChange24h = pickRateChange24h > 0 ? MAX_CHANGE_PERCENT : MIN_CHANGE_PERCENT
            }
          }
          
          // Рассчитываем proContestRateChange7d
          let proContestRateChange7d = 0
          const currentProContestRate = heroStat.proContestRate || 0
          if (stats7DaysAgo && stats7DaysAgo.proContestRate && stats7DaysAgo.proContestRate > 0) {
            proContestRateChange7d = ((currentProContestRate - stats7DaysAgo.proContestRate) / stats7DaysAgo.proContestRate) * 100
            
            // ВАЛИДАЦИЯ: Ограничиваем изменения разумными пределами
            const MAX_CHANGE_PERCENT = 1000
            const MIN_CHANGE_PERCENT = -1000
            if (proContestRateChange7d > MAX_CHANGE_PERCENT || proContestRateChange7d < MIN_CHANGE_PERCENT) {
              console.warn(`HeroStats: аномальное изменение proContestRate для героя ${heroStat.heroId} (High Rank): ${proContestRateChange7d.toFixed(2)}% (ограничено)`)
              proContestRateChange7d = proContestRateChange7d > 0 ? MAX_CHANGE_PERCENT : MIN_CHANGE_PERCENT
            }
          }
          
          const statsData = {
            pickRate: heroStat.pickRate || 0,
            pickRatePercent: heroStat.pickRatePercent || 0, // Процент пиков
            winRate: heroStat.winRate || 0,
            banRate: heroStat.banRate || 0,
            contestRate: heroStat.contestRate || 0,
            contestRatePercent: heroStat.contestRatePercent || 0, // Процент контестов
            matchCount: heroStat.matchCount || 0,
            // Про-статистика
            proPick: heroStat.proPick || 0,
            proBan: heroStat.proBan || 0,
            proContestRate: currentProContestRate,
            // Изменения за 7 дней и 24 часа
            pickRateChange7d: pickRateChange7d,  // Immortal за неделю
            pickRateChange24h: pickRateChange24h,  // Immortal за 24 часа (для Мета сигнала)
            proContestRateChange7d: proContestRateChange7d
            // Убрано: contestRateChange7d (фейк, дублировал pickRateChange7d)
          }
          
          batchUpdates.push({row, statsData})
          updatedCount++
        } catch (e) {
          console.error(`HeroStats: ошибка подготовки High Rank для героя ${heroStat.heroId}:`, e)
          errorCount++
        }
      })
    }
    
    // Подготавливаем статистику All Ranks
    if (result.allRanks && Array.isArray(result.allRanks)) {
      result.allRanks.forEach(heroStat => {
        try {
          const row = heroStats_getOrCreateRow(
            heroStat.heroId,
            heroStat.heroName,
            'All Ranks'
          )
          
          // Получаем статистику 7 дней назад и 24 часа назад для расчета изменений
          const stats7DaysAgo = heroStats_getStats7DaysAgo(heroStat.heroId, 'All Ranks')
          const stats24HoursAgo = heroStats_getStats24HoursAgo(heroStat.heroId, 'All Ranks')
          
          // Рассчитываем pickRateChange7d (All Ranks за неделю)
          let pickRateChange7d = 0
          if (stats7DaysAgo && stats7DaysAgo.pickRatePercent !== undefined && stats7DaysAgo.pickRatePercent > 0) {
            const currentPickRatePercent = heroStat.pickRatePercent || 0
            pickRateChange7d = ((currentPickRatePercent - stats7DaysAgo.pickRatePercent) / stats7DaysAgo.pickRatePercent) * 100
            
            // ВАЛИДАЦИЯ: Ограничиваем изменения разумными пределами (-1000% до +1000%)
            const MAX_CHANGE_PERCENT = 1000
            const MIN_CHANGE_PERCENT = -1000
            if (pickRateChange7d > MAX_CHANGE_PERCENT || pickRateChange7d < MIN_CHANGE_PERCENT) {
              console.warn(`HeroStats: аномальное изменение pickRate для героя ${heroStat.heroId} (All Ranks): ${pickRateChange7d.toFixed(2)}% (ограничено)`)
              pickRateChange7d = pickRateChange7d > 0 ? MAX_CHANGE_PERCENT : MIN_CHANGE_PERCENT
            }
          }
          
          // Рассчитываем pickRateChange24h (All Ranks за 24 часа) - для Мета сигнала
          let pickRateChange24h = 0
          if (stats24HoursAgo && stats24HoursAgo.pickRatePercent !== undefined && stats24HoursAgo.pickRatePercent > 0) {
            const currentPickRatePercent = heroStat.pickRatePercent || 0
            pickRateChange24h = ((currentPickRatePercent - stats24HoursAgo.pickRatePercent) / stats24HoursAgo.pickRatePercent) * 100
            
            // ВАЛИДАЦИЯ: Ограничиваем изменения разумными пределами
            const MAX_CHANGE_PERCENT = 1000
            const MIN_CHANGE_PERCENT = -1000
            if (pickRateChange24h > MAX_CHANGE_PERCENT || pickRateChange24h < MIN_CHANGE_PERCENT) {
              console.warn(`HeroStats: аномальное изменение pickRate (24h) для героя ${heroStat.heroId} (All Ranks): ${pickRateChange24h.toFixed(2)}% (ограничено)`)
              pickRateChange24h = pickRateChange24h > 0 ? MAX_CHANGE_PERCENT : MIN_CHANGE_PERCENT
            }
          }
          
          // Рассчитываем proContestRateChange7d
          let proContestRateChange7d = 0
          const currentProContestRate = heroStat.proContestRate || 0
          if (stats7DaysAgo && stats7DaysAgo.proContestRate && stats7DaysAgo.proContestRate > 0) {
            proContestRateChange7d = ((currentProContestRate - stats7DaysAgo.proContestRate) / stats7DaysAgo.proContestRate) * 100
            
            // ВАЛИДАЦИЯ: Ограничиваем изменения разумными пределами
            const MAX_CHANGE_PERCENT = 1000
            const MIN_CHANGE_PERCENT = -1000
            if (proContestRateChange7d > MAX_CHANGE_PERCENT || proContestRateChange7d < MIN_CHANGE_PERCENT) {
              console.warn(`HeroStats: аномальное изменение proContestRate для героя ${heroStat.heroId} (All Ranks): ${proContestRateChange7d.toFixed(2)}% (ограничено)`)
              proContestRateChange7d = proContestRateChange7d > 0 ? MAX_CHANGE_PERCENT : MIN_CHANGE_PERCENT
            }
          }
          
          const statsData = {
            pickRate: heroStat.pickRate || 0,
            pickRatePercent: heroStat.pickRatePercent || 0, // Процент пиков
            winRate: heroStat.winRate || 0,
            banRate: heroStat.banRate || 0,
            contestRate: heroStat.contestRate || 0,
            contestRatePercent: heroStat.contestRatePercent || 0, // Процент контестов (равен pickRatePercent)
            matchCount: heroStat.matchCount || 0,
            // Про-статистика
            proPick: heroStat.proPick || 0,
            proBan: heroStat.proBan || 0,
            proContestRate: currentProContestRate,
            // Изменения за 7 дней и 24 часа
            pickRateChange7d: pickRateChange7d,  // All Ranks за неделю
            pickRateChange24h: pickRateChange24h,  // All Ranks за 24 часа (для Мета сигнала)
            proContestRateChange7d: proContestRateChange7d
            // Убрано: contestRateChange7d (фейк, дублировал pickRateChange7d)
          }
          
          batchUpdates.push({row, statsData})
          updatedCount++
        } catch (e) {
          console.error(`HeroStats: ошибка подготовки All Ranks для героя ${heroStat.heroId}:`, e)
          errorCount++
        }
      })
    }
    
    // БATCH ЗАПИСЬ: Записываем все обновления одним batch операциями
    if (batchUpdates.length > 0) {
      // Сортируем по строке для последовательной записи
      batchUpdates.sort((a, b) => a.row - b.row)
      
      // Подготавливаем массив значений для batch записи
      const values = batchUpdates.map(update => [JSON.stringify(update.statsData)])
      const rows = batchUpdates.map(update => update.row)
      
      // Batch запись: записываем все строки одним запросом
      // Группируем последовательные строки для эффективной batch записи
      let batchStart = 0
      for (let i = 1; i <= rows.length; i++) {
        // Если строки последовательные или достигли конца массива
        if (i === rows.length || rows[i] !== rows[i-1] + 1) {
          const batchSize = i - batchStart
          const batchRows = rows.slice(batchStart, i)
          const batchValues = values.slice(batchStart, i)
          
          // Записываем batch последовательных строк
          const range = sheet.getRange(batchRows[0], dataCol, batchSize, 1)
          range.setValues(batchValues)
          
          batchStart = i
        }
      }
    }
    
    // Форматируем новую колонку
    heroStats_formatDataColumn_(sheet, dataCol)
    
    console.log(`HeroStats: обновление завершено. Обновлено: ${updatedCount}, ошибок: ${errorCount}`)
    logAutoAction_('HeroStats', 'Обновление статистики', `OK (${updatedCount} записей, ${errorCount} ошибок)`)
    
    // Удаляем старые данные (>90 дней)
    heroStats_archiveOldData()
    
  } catch (e) {
    console.error('HeroStats: критическая ошибка при обновлении:', e)
    logAutoAction_('HeroStats', 'Обновление статистики', `Ошибка: ${e.message}`)
  }
}

/**
 * Удаляет данные старше HERO_STATS_HISTORY_DAYS дней
 */
function heroStats_archiveOldData() {
  const sheet = getHeroStatsSheet_()
  if (!sheet) return
  
  const firstDataCol = HERO_STATS_COLUMNS.FIRST_DATA_COL
  const lastCol = sheet.getLastColumn()
  
  if (lastCol < firstDataCol) return
  
  const now = new Date()
  const cutoffDate = new Date(now.getTime() - HERO_STATS_HISTORY_DAYS * 24 * 60 * 60 * 1000)
  
  const dateRow = sheet.getRange(HEADER_ROW, firstDataCol, 1, lastCol - firstDataCol + 1).getDisplayValues()[0]
  const colsToDelete = []
  
  for (let i = 0; i < dateRow.length; i++) {
    const header = String(dateRow[i] || '').trim()
    if (!header) continue
    
    try {
      // Парсим дату из формата "DD.MM.YY HH:MM"
      const parts = header.split(' ')
      if (parts.length !== 2) continue
      
      const dateStr = parts[0]
      const [day, month, year] = dateStr.split('.')
      const parsedDate = new Date(2000 + parseInt(year), parseInt(month) - 1, parseInt(day))
      
      if (parsedDate < cutoffDate) {
        colsToDelete.push(firstDataCol + i)
      }
    } catch (e) {
      // Пропускаем некорректные заголовки
      continue
    }
  }
  
  // Удаляем колонки справа налево (чтобы индексы не сдвигались)
  if (colsToDelete.length > 0) {
    colsToDelete.sort((a, b) => b - a)
    colsToDelete.forEach(col => {
      sheet.deleteColumn(col)
    })
    
    console.log(`HeroStats: удалено ${colsToDelete.length} старых колонок (старше ${HERO_STATS_HISTORY_DAYS} дней)`)
  }
}

