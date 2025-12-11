/**
 * HeroMapping - Управление маппингом герой → предметы
 * 
 * Связывает предметы из History с героями для анализа статистики
 */

const HERO_MAPPING_CONFIG = {
  COLUMNS: HERO_MAPPING_COLUMNS
}

/**
 * Форматирование листа HeroMapping
 */
function heroMapping_formatTable() {
  const sheet = getOrCreateHeroMappingSheet_()
  const lastRow = sheet.getLastRow()

  // Заголовки
  const headers = HEADERS.HERO_MAPPING
  sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers])

  formatHeaderRange_(sheet.getRange(HEADER_ROW, 1, 1, headers.length))

  sheet.setRowHeight(HEADER_ROW, HEADER_HEIGHT)
  if (lastRow > 1) {
    sheet.setRowHeights(DATA_START_ROW, lastRow - 1, ROW_HEIGHT)
  }

  // Ширины колонок
  sheet.setColumnWidth(1, 250)  // A - Item Name
  sheet.setColumnWidth(2, 150)  // B - Hero Name
  sheet.setColumnWidth(3, 80)   // C - Hero ID
  sheet.setColumnWidth(4, 100)  // D - Auto-detected
  sheet.setColumnWidth(5, 100)  // E - Verified
  sheet.setColumnWidth(6, 120)  // F - Category

  if (lastRow > 1) {
    // Форматирование данных
    sheet.getRange(DATA_START_ROW, 1, lastRow - 1, headers.length)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setWrap(false)
    sheet.getRange(`${HERO_MAPPING_COLUMNS.ITEM_NAME}${DATA_START_ROW}:${HERO_MAPPING_COLUMNS.ITEM_NAME}${lastRow}`)
      .setHorizontalAlignment('left')
    sheet.getRange(`${HERO_MAPPING_COLUMNS.HERO_NAME}${DATA_START_ROW}:${HERO_MAPPING_COLUMNS.HERO_NAME}${lastRow}`)
      .setHorizontalAlignment('left')
    
    // Формат числа для Hero ID
    sheet.getRange(`${HERO_MAPPING_COLUMNS.HERO_ID}${DATA_START_ROW}:${HERO_MAPPING_COLUMNS.HERO_ID}${lastRow}`)
      .setNumberFormat(NUMBER_FORMATS.INTEGER)
    
    // Чекбоксы для Auto-detected и Verified (если еще не установлены)
    const autoDetectedRange = sheet.getRange(`${HERO_MAPPING_COLUMNS.AUTO_DETECTED}${DATA_START_ROW}:${HERO_MAPPING_COLUMNS.AUTO_DETECTED}${lastRow}`)
    const verifiedRange = sheet.getRange(`${HERO_MAPPING_COLUMNS.VERIFIED}${DATA_START_ROW}:${HERO_MAPPING_COLUMNS.VERIFIED}${lastRow}`)
    
    // Проверяем, установлены ли уже чекбоксы (проверяем по данным в первой ячейке)
    try {
      autoDetectedRange.insertCheckboxes()
      verifiedRange.insertCheckboxes()
    } catch (e) {
      // Чекбоксы уже установлены, игнорируем ошибку
      console.log('HeroMapping: чекбоксы уже установлены')
    }
  }

  sheet.setFrozenRows(HEADER_ROW)
  sheet.setFrozenColumns(1) // Замораживаем первую колонку (Item Name)

  try {
    SpreadsheetApp.getUi().alert('Форматирование завершено (HeroMapping)')
  } catch (e) {
    console.log('HeroMapping: невозможно показать UI в данном контексте')
  }
}

/**
 * Получает имя героя для предмета
 * @param {string} itemName - Название предмета
 * @returns {Object|null} {heroName, heroId, category} или null
 */
function heroMapping_getHeroForItem(itemName) {
  const sheet = getHeroMappingSheet_()
  if (!sheet) return null
  
  const lastRow = sheet.getLastRow()
  if (lastRow < DATA_START_ROW) return null
  
  const itemNames = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, 1).getValues()
  
  for (let i = 0; i < itemNames.length; i++) {
    const rowItemName = String(itemNames[i][0] || '').trim()
    if (rowItemName === itemName) {
      const row = DATA_START_ROW + i
      const heroName = sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.HERO_NAME)).getValue()
      const heroId = sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.HERO_ID)).getValue()
      const category = sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.CATEGORY)).getValue()
      
      if (!heroName || String(heroName).trim().length === 0) {
        return null
      }
      
      return {
        heroName: String(heroName).trim(),
        heroId: Number(heroId) || null,
        category: String(category || '').trim()
      }
    }
  }
  
  return null
}

/**
 * Получает все маппинги
 * @returns {Object} Объект {itemName: {heroName, heroId, category}}
 */
function heroMapping_getAllMappings() {
  const sheet = getHeroMappingSheet_()
  if (!sheet) return {}
  
  const lastRow = sheet.getLastRow()
  if (lastRow < DATA_START_ROW) return {}
  
  const mappings = {}
  const count = lastRow - HEADER_ROW
  
  const itemNames = sheet.getRange(DATA_START_ROW, 1, count, 1).getValues()
  const heroNames = sheet.getRange(DATA_START_ROW, 2, count, 1).getValues()
  const heroIds = sheet.getRange(DATA_START_ROW, 3, count, 1).getValues()
  const categories = sheet.getRange(DATA_START_ROW, 6, count, 1).getValues()
  
  for (let i = 0; i < count; i++) {
    const itemName = String(itemNames[i][0] || '').trim()
    const heroName = String(heroNames[i][0] || '').trim()
    const heroId = Number(heroIds[i][0]) || null
    const category = String(categories[i][0] || '').trim()
    
    if (itemName && heroName) {
      mappings[itemName] = {
        heroName: heroName,
        heroId: heroId,
        category: category || 'Hero Item'
      }
    }
  }
  
  return mappings
}

/**
 * Обновляет или создает маппинг для предмета
 * @param {string} itemName - Название предмета
 * @param {string} heroName - Имя героя (null для удаления)
 * @param {boolean} autoDetected - Автоматически определено
 * @returns {number} Номер строки
 */
function heroMapping_updateItem(itemName, heroName, autoDetected = false) {
  const sheet = getOrCreateHeroMappingSheet_()
  const lastRow = sheet.getLastRow()
  
  // Ищем существующую строку
  let row = -1
  if (lastRow >= DATA_START_ROW) {
    const itemNames = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, 1).getValues()
    for (let i = 0; i < itemNames.length; i++) {
      const rowItemName = String(itemNames[i][0] || '').trim()
      if (rowItemName === itemName) {
        row = DATA_START_ROW + i
        break
      }
    }
  }
  
  // Создаём новую строку если не найдена
  if (row === -1) {
    row = Math.max(lastRow + 1, DATA_START_ROW)
    sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.ITEM_NAME)).setValue(itemName)
    
    // Форматируем новую строку
    heroMapping_formatNewRow_(sheet, row)
  }
  
  // Обновляем данные
  if (heroName && heroName.trim().length > 0) {
    sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.HERO_NAME)).setValue(heroName.trim())
    
    // Пытаемся определить Hero ID (можно расширить позже)
    // Пока оставляем пустым, можно заполнить вручную или через справочник
    sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.AUTO_DETECTED)).setValue(autoDetected)
    sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.CATEGORY)).setValue('Hero Item')
  } else {
    // Если герой не указан - это общий предмет
    sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.HERO_NAME)).setValue('')
    sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.CATEGORY)).setValue('Common Item')
  }
  
  return row
}

/**
 * Форматирует новую строку в HeroMapping
 * @param {Sheet} sheet - Лист HeroMapping
 * @param {number} row - Номер строки
 */
function heroMapping_formatNewRow_(sheet, row) {
  sheet.getRange(row, 1, 1, 6)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setWrap(false)
  sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.ITEM_NAME)).setHorizontalAlignment('left') // Item Name
  sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.HERO_NAME)).setHorizontalAlignment('left') // Hero Name
  
  // Формат числа для Hero ID
  sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.HERO_ID))
    .setNumberFormat(NUMBER_FORMATS.INTEGER)
  
  // Чекбоксы (безопасная установка)
  try {
    sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.AUTO_DETECTED)).insertCheckboxes()
    sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.VERIFIED)).insertCheckboxes()
  } catch (e) {
    // Чекбоксы уже установлены
    console.log('HeroMapping: чекбоксы уже установлены для строки', row)
  }
  
  sheet.setRowHeight(row, ROW_HEIGHT)
}

/**
 * Автоматическое определение героев через SteamWebAPI.ru (tag5)
 * Обрабатывает все предметы из History
 */
function heroMapping_autoDetectFromSteamWebAPI() {
  console.log('HeroMapping: начало автоматического определения героев')
  
  const historySheet = getHistorySheet_()
  if (!historySheet) {
    console.error('HeroMapping: лист History не найден')
    return
  }
  
  const lastRow = historySheet.getLastRow()
  if (lastRow < DATA_START_ROW) {
    console.log('HeroMapping: нет предметов в History')
    return
  }
  
  // Получаем все названия предметов из History
  const itemNames = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), lastRow - HEADER_ROW, 1).getValues()
  const uniqueItemNames = []
  const itemNameSet = new Set()
  
  itemNames.forEach(row => {
    const name = String(row[0] || '').trim()
    if (name && !itemNameSet.has(name)) {
      uniqueItemNames.push(name)
      itemNameSet.add(name)
    }
  })
  
  console.log(`HeroMapping: найдено ${uniqueItemNames.length} уникальных предметов`)
  
  if (uniqueItemNames.length === 0) {
    return
  }
  
  // Получаем данные через SteamWebAPI.ru пакетами
  const batchResult = steamWebAPI_fetchItemsBatch(uniqueItemNames, 'dota2')
  
  if (!batchResult.ok && batchResult.items) {
    console.warn('HeroMapping: некоторые предметы не были получены')
  }
  
  let detectedCount = 0
  let skippedCount = 0
  let errorCount = 0
  
  // Обрабатываем каждый предмет
  uniqueItemNames.forEach(itemName => {
    try {
      const itemData = batchResult.items[itemName.toLowerCase()]
      
      if (!itemData) {
        skippedCount++
        return
      }
      
      // Парсим данные
      const parsedData = steamWebAPI_parseItemData(itemData)
      
      // Пытаемся определить героя из tag5
      const heroName = steamWebAPI_getHeroNameFromTags(parsedData)
      
      if (heroName) {
        // Обновляем маппинг
        heroMapping_updateItem(itemName, heroName, true)
        detectedCount++
        console.log(`HeroMapping: определен герой для "${itemName}": ${heroName}`)
      } else {
        // Если герой не найден - помечаем как Common Item
        heroMapping_updateItem(itemName, null, false)
        skippedCount++
      }
    } catch (e) {
      console.error(`HeroMapping: ошибка обработки предмета "${itemName}":`, e)
      errorCount++
    }
  })
  
  console.log(`HeroMapping: автоматическое определение завершено. Определено: ${detectedCount}, пропущено: ${skippedCount}, ошибок: ${errorCount}`)
  
  try {
    logAutoAction_('HeroMapping', 'Автоопределение героев', `OK (определено: ${detectedCount}, пропущено: ${skippedCount}, ошибок: ${errorCount})`)
    SpreadsheetApp.getUi().alert(`Автоопределение завершено!\n\nОпределено: ${detectedCount}\nПропущено: ${skippedCount}\nОшибок: ${errorCount}`)
  } catch (e) {
    console.log('HeroMapping: невозможно показать UI')
  }
}

/**
 * Синхронизирует предметы из History с HeroMapping
 * Добавляет новые предметы, если их еще нет
 */
function heroMapping_syncWithHistory() {
  const historySheet = getHistorySheet_()
  if (!historySheet) return
  
  const mappingSheet = getOrCreateHeroMappingSheet_()
  const historyLastRow = historySheet.getLastRow()
  
  if (historyLastRow < DATA_START_ROW) return
  
  // Получаем все предметы из History
  const itemNames = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), historyLastRow - HEADER_ROW, 1).getValues()
  const itemNameSet = new Set()
  
  itemNames.forEach(row => {
    const name = String(row[0] || '').trim()
    if (name) {
      itemNameSet.add(name)
    }
  })
  
  // Получаем существующие маппинги
  const mappingLastRow = mappingSheet.getLastRow()
  const existingItems = new Set()
  
  if (mappingLastRow >= DATA_START_ROW) {
    const mappingNames = mappingSheet.getRange(DATA_START_ROW, 1, mappingLastRow - HEADER_ROW, 1).getValues()
    mappingNames.forEach(row => {
      const name = String(row[0] || '').trim()
      if (name) {
        existingItems.add(name)
      }
    })
  }
  
  // Добавляем отсутствующие предметы
  let addedCount = 0
  itemNameSet.forEach(itemName => {
    if (!existingItems.has(itemName)) {
      heroMapping_updateItem(itemName, null, false)
      addedCount++
    }
  })
  
  if (addedCount > 0) {
    console.log(`HeroMapping: добавлено ${addedCount} новых предметов из History`)
  }
}

