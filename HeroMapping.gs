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
  sheet.setColumnWidth(2, 100)  // B - Image
  sheet.setColumnWidth(3, 150)  // C - Hero Name
  sheet.setColumnWidth(4, 80)   // D - Hero ID
  sheet.setColumnWidth(5, 100)  // E - Auto-detected
  sheet.setColumnWidth(6, 120)  // F - Category (было G)

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
    
    // Форматирование изображений (колонка B)
    sheet.getRange(`${HERO_MAPPING_COLUMNS.IMAGE}${DATA_START_ROW}:${HERO_MAPPING_COLUMNS.IMAGE}${lastRow}`)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
    
    // Формат числа для Hero ID
    sheet.getRange(`${HERO_MAPPING_COLUMNS.HERO_ID}${DATA_START_ROW}:${HERO_MAPPING_COLUMNS.HERO_ID}${lastRow}`)
      .setNumberFormat(NUMBER_FORMATS.INTEGER)
    
    // Чекбоксы для Auto-detected (если еще не установлены)
    const autoDetectedRange = sheet.getRange(`${HERO_MAPPING_COLUMNS.AUTO_DETECTED}${DATA_START_ROW}:${HERO_MAPPING_COLUMNS.AUTO_DETECTED}${lastRow}`)
    
    // Проверяем, установлены ли уже чекбоксы
    try {
      autoDetectedRange.insertCheckboxes()
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
      const heroName = sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.HERO_NAME)).getValue()  // C
      const heroId = sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.HERO_ID)).getValue()      // D
      const category = sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.CATEGORY)).getValue()   // G
      
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
  
  const itemNames = sheet.getRange(DATA_START_ROW, 1, count, 1).getValues()  // A - Item Name
  const heroNames = sheet.getRange(DATA_START_ROW, 3, count, 1).getValues()  // C - Hero Name (было 2)
  const heroIds = sheet.getRange(DATA_START_ROW, 4, count, 1).getValues()    // D - Hero ID (было 3)
  const categories = sheet.getRange(DATA_START_ROW, 6, count, 1).getValues() // F - Category (было G/7)
  
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
  
  // Пытаемся получить изображение предмета (если доступно)
  // Можно получить из SteamWebAPI, но это требует дополнительного запроса
  // Пока оставляем пустым, будет заполняться при автоопределении или синхронизации
  
  return row
}

/**
 * Форматирует новую строку в HeroMapping
 * @param {Sheet} sheet - Лист HeroMapping
 * @param {number} row - Номер строки
 */
function heroMapping_formatNewRow_(sheet, row) {
  sheet.getRange(row, 1, 1, 6)  // 6 колонок (убрали Verified)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setWrap(false)
  sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.ITEM_NAME)).setHorizontalAlignment('left') // Item Name
  sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.HERO_NAME)).setHorizontalAlignment('left') // Hero Name
  sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.IMAGE))  // Image
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
  
  // Формат числа для Hero ID
  sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.HERO_ID))
    .setNumberFormat(NUMBER_FORMATS.INTEGER)
  
  // Чекбоксы (безопасная установка)
  try {
    sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.AUTO_DETECTED)).insertCheckboxes()
  } catch (e) {
    // Чекбоксы уже установлены
    console.log('HeroMapping: чекбоксы уже установлены для строки', row)
  }
  
  sheet.setRowHeight(row, ROW_HEIGHT)
}

/**
 * Обновляет или создает маппинги для нескольких предметов batch операцией
 * @param {Array<Object>} updates - Массив {itemName, heroName, autoDetected, imageUrl}
 * @private
 */
function heroMapping_updateItemsBatch_(updates) {
  if (!updates || updates.length === 0) return
  
  const sheet = getOrCreateHeroMappingSheet_()
  const lastRow = sheet.getLastRow()
  
  // Читаем существующие предметы batch операцией
  const existingItemsMap = new Map()
  if (lastRow >= DATA_START_ROW) {
    const count = lastRow - HEADER_ROW
    const existingNames = sheet.getRange(DATA_START_ROW, 1, count, 1).getValues()
    for (let i = 0; i < count; i++) {
      const name = String(existingNames[i][0] || '').trim()
      if (name) {
        existingItemsMap.set(name, DATA_START_ROW + i)
      }
    }
  }
  
  // Подготавливаем данные для batch записи
  const newRows = []
  const updateRows = []
  
  updates.forEach(update => {
    const itemName = String(update.itemName || '').trim()
    if (!itemName) return
    
    const existingRow = existingItemsMap.get(itemName)
    const heroName = update.heroName ? String(update.heroName).trim() : ''
    const category = heroName ? 'Hero Item' : 'Common Item'
    const imageUrl = update.imageUrl || ''
    
    if (existingRow) {
      // Обновляем существующую строку
      updateRows.push({
        row: existingRow,
        values: [
          null,  // A - Item Name (не меняем)
          imageUrl,  // B - Image
          heroName,  // C - Hero Name
          null,  // D - Hero ID (пока оставляем пустым)
          update.autoDetected || false,  // E - Auto-detected
          category  // F - Category
        ]
      })
    } else {
      // Создаем новую строку
      newRows.push({
        itemName: itemName,
        values: [
          itemName,  // A - Item Name
          imageUrl,  // B - Image
          heroName,  // C - Hero Name
          null,  // D - Hero ID
          update.autoDetected || false,  // E - Auto-detected
          category  // F - Category (было G)
        ]
      })
    }
  })
  
  // Batch запись обновлений (группируем по колонкам для эффективной записи)
  if (updateRows.length > 0) {
    // Сортируем по строке для последовательной записи
    updateRows.sort((a, b) => a.row - b.row)
    
    // Группируем обновления в последовательные диапазоны для batch записи
    // Записываем каждую колонку отдельно batch операцией
    const imageValues = []
    const heroNameValues = []
    const autoDetectedValues = []
    const categoryValues = []
    const imageRows = []
    const heroNameRows = []
    const autoDetectedRows = []
    const categoryRows = []
    
    updateRows.forEach(u => {
      if (u.values[1] !== null) {  // Image
        imageRows.push(u.row)
        imageValues.push([u.values[1]])
      }
      if (u.values[2] !== null) {  // Hero Name
        heroNameRows.push(u.row)
        heroNameValues.push([u.values[2]])
      }
      if (u.values[4] !== null) {  // Auto-detected
        autoDetectedRows.push(u.row)
        autoDetectedValues.push([u.values[4]])
      }
      if (u.values[5] !== null) {  // Category (теперь позиция 5, была 6)
        categoryRows.push(u.row)
        categoryValues.push([u.values[5]])
      }
    })
    
    // Batch запись каждой колонки
    if (imageValues.length > 0) {
      // Подготавливаем формулы для batch записи
      const imageFormulas = []
      for (let idx = 0; idx < imageValues.length; idx++) {
        const url = imageValues[idx][0]
        if (url && url.trim().length > 0) {
          imageFormulas.push([`=IMAGE("${url}")`])
        } else {
          imageFormulas.push([''])
        }
      }
      // Batch запись формул
      const imageRange = sheet.getRange(imageRows[0], getColumnIndex(HERO_MAPPING_COLUMNS.IMAGE), imageFormulas.length, 1)
      imageRange.setFormulas(imageFormulas)
    }
    if (heroNameValues.length > 0) {
      sheet.getRange(heroNameRows[0], getColumnIndex(HERO_MAPPING_COLUMNS.HERO_NAME), heroNameValues.length, 1).setValues(heroNameValues)
    }
    if (autoDetectedValues.length > 0) {
      sheet.getRange(autoDetectedRows[0], getColumnIndex(HERO_MAPPING_COLUMNS.AUTO_DETECTED), autoDetectedValues.length, 1).setValues(autoDetectedValues)
    }
    if (categoryValues.length > 0) {
      sheet.getRange(categoryRows[0], getColumnIndex(HERO_MAPPING_COLUMNS.CATEGORY), categoryValues.length, 1).setValues(categoryValues)
    }
  }
  
  // Batch запись новых строк
  if (newRows.length > 0) {
    const startRow = Math.max(lastRow + 1, DATA_START_ROW)
    const values = newRows.map(r => r.values)
    
    // Записываем все новые строки за раз
    sheet.getRange(startRow, 1, newRows.length, 6).setValues(values)  // 6 колонок (убрали Verified)
    
    // Добавляем формулы изображений для новых строк batch операцией
    const imageFormulas = newRows.map(r => {
      const imageUrl = r.values[1]
      return imageUrl && imageUrl.trim().length > 0 ? [`=IMAGE("${imageUrl}")`] : ['']
    })
    if (imageFormulas.length > 0) {
      sheet.getRange(startRow, getColumnIndex(HERO_MAPPING_COLUMNS.IMAGE), imageFormulas.length, 1).setFormulas(imageFormulas)
    }
    
    // Форматируем новые строки
    for (let i = 0; i < newRows.length; i++) {
      heroMapping_formatNewRow_(sheet, startRow + i)
    }
  }
}

/**
 * Автоматическое определение героев через SteamWebAPI.ru (tag5)
 * Обрабатывает все предметы из History с batch операциями и fallback на item_by_nameid
 */
function heroMapping_autoDetectFromSteamWebAPI() {
  console.log('HeroMapping: начало автоматического определения героев')
  
  const historySheet = getHistorySheet_()
  if (!historySheet) {
    console.error('HeroMapping: лист History не найден')
    SpreadsheetApp.getUi().alert('Ошибка: лист History не найден')
    return
  }
  
  const lastRow = historySheet.getLastRow()
  if (lastRow < DATA_START_ROW) {
    console.log('HeroMapping: нет предметов в History')
    SpreadsheetApp.getUi().alert('Нет предметов в History для автоопределения')
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
  
  // Диалог подтверждения
  const ui = SpreadsheetApp.getUi()
  const confirmResponse = ui.alert(
    'Автоопределение героев',
    `Будет обработано ${uniqueItemNames.length} предметов.\nЭто может занять несколько минут.\n\nПродолжить?`,
    ui.ButtonSet.YES_NO
  )
  
  if (confirmResponse !== ui.Button.YES) {
    return
  }
  
  // Получаем данные через SteamWebAPI.ru пакетами
  const batchResult = steamWebAPI_fetchItemsBatch(uniqueItemNames, 'dota2')
  
  // Обрабатываем результаты и пробуем fallback для не найденных
  const updates = []
  const notFoundItems = []
  
  uniqueItemNames.forEach(itemName => {
    // Пробуем найти в основном результате
    let itemData = null
    const searchKeys = [
      itemName.toLowerCase().trim(),
      itemName.toLowerCase().trim().replace(/[''""`]/g, ''),
      itemName.toLowerCase().trim().replace(/[''""`]/g, '').replace(/\s+/g, ' ')
    ]
    
    for (const key of searchKeys) {
      if (batchResult.items && batchResult.items[key]) {
        itemData = batchResult.items[key]
        break
      }
    }
    
    // Если не нашли, пробуем fallback на item_by_nameid
    if (!itemData) {
      notFoundItems.push(itemName)
      return // Будем обрабатывать отдельно
    }
    
    // Парсим данные
    const parsedData = steamWebAPI_parseItemData(itemData)
    const heroName = steamWebAPI_getHeroNameFromTags(parsedData)
    const imageUrl = parsedData.imageUrl || ''
    
    updates.push({
      itemName: itemName,
      heroName: heroName || null,
      autoDetected: !!heroName,
      imageUrl: imageUrl
    })
  })
  
  // Обрабатываем не найденные предметы через fallback (item_by_nameid)
  console.log(`HeroMapping: ${notFoundItems.length} предметов не найдено через основной endpoint, пробуем fallback`)
  for (const itemName of notFoundItems) {
    try {
      const fallbackResult = steamWebAPI_fetchItemByNameIdViaName(itemName, 'dota2')
      if (fallbackResult.ok && fallbackResult.item) {
        const parsedData = steamWebAPI_parseItemData(fallbackResult.item)
        const heroName = steamWebAPI_getHeroNameFromTags(parsedData)
        const imageUrl = parsedData.imageUrl || ''
        
        updates.push({
          itemName: itemName,
          heroName: heroName || null,
          autoDetected: !!heroName,
          imageUrl: imageUrl
        })
        console.log(`HeroMapping: предмет "${itemName}" найден через fallback`)
      }
    } catch (e) {
      console.warn(`HeroMapping: ошибка fallback для "${itemName}":`, e.message)
      // Добавляем без героя
      updates.push({
        itemName: itemName,
        heroName: null,
        autoDetected: false,
        imageUrl: ''
      })
    }
    
    // Небольшая задержка между fallback запросами
    Utilities.sleep(LIMITS.BASE_DELAY_MS)
  }
  
  // Batch обновление всех предметов
  heroMapping_updateItemsBatch_(updates)
  
  const detectedCount = updates.filter(u => u.heroName).length
  const skippedCount = updates.filter(u => !u.heroName).length
  
  console.log(`HeroMapping: автоматическое определение завершено. Определено: ${detectedCount}, пропущено: ${skippedCount}`)
  
  try {
    logAutoAction_('HeroMapping', 'Автоопределение героев', `OK (определено: ${detectedCount}, пропущено: ${skippedCount})`)
    SpreadsheetApp.getUi().alert(`Автоопределение завершено!\n\nОпределено: ${detectedCount}\nПропущено: ${skippedCount}`)
  } catch (e) {
    console.log('HeroMapping: невозможно показать UI')
  }
}

/**
 * Синхронизирует предметы из History с HeroMapping
 * Добавляет только новые предметы, если их еще нет (batch операция)
 */
function heroMapping_syncWithHistory() {
  const historySheet = getHistorySheet_()
  if (!historySheet) {
    SpreadsheetApp.getUi().alert('Ошибка: лист History не найден')
    return
  }
  
  const mappingSheet = getOrCreateHeroMappingSheet_()
  const historyLastRow = historySheet.getLastRow()
  
  if (historyLastRow < DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('Нет предметов в History для синхронизации')
    return
  }
  
  // Batch чтение всех предметов из History
  const itemNames = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), historyLastRow - HEADER_ROW, 1).getValues()
  const itemNameSet = new Set()
  
  itemNames.forEach(row => {
    const name = String(row[0] || '').trim()
    if (name) {
      itemNameSet.add(name)
    }
  })
  
  // Batch чтение существующих маппингов
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
  
  // Находим только новые предметы
  const newItems = []
  itemNameSet.forEach(itemName => {
    if (!existingItems.has(itemName)) {
      newItems.push(itemName)
    }
  })
  
  if (newItems.length === 0) {
    SpreadsheetApp.getUi().alert('Все предметы из History уже синхронизированы')
    return
  }
  
  // Batch добавление новых предметов
  const updates = newItems.map(itemName => ({
    itemName: itemName,
    heroName: null,
    autoDetected: false,
    imageUrl: ''
  }))
  
  heroMapping_updateItemsBatch_(updates)
  
  console.log(`HeroMapping: добавлено ${newItems.length} новых предметов из History`)
  
  try {
    logAutoAction_('HeroMapping', 'Синхронизация с History', `Добавлено: ${newItems.length} новых предметов`)
    SpreadsheetApp.getUi().alert(`Синхронизация завершена!\n\nДобавлено новых предметов: ${newItems.length}`)
  } catch (e) {
    console.log('HeroMapping: невозможно показать UI')
  }
}

