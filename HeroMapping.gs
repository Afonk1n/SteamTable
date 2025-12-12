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
  if (!headers || !Array.isArray(headers) || headers.length === 0) {
    console.error('HeroMapping: HEADERS.HERO_MAPPING не определен или пуст')
    SpreadsheetApp.getUi().alert('Ошибка: HEADERS.HERO_MAPPING не определен в Constants.gs')
    return
  }
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
  sheet.setColumnWidth(5, 120)  // E - Category

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
  }

  sheet.setFrozenRows(HEADER_ROW)
  sheet.setFrozenColumns(1) // Замораживаем первую колонку (Item Name)

  console.log('HeroMapping: форматирование завершено')
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
      const category = sheet.getRange(row, getColumnIndex(HERO_MAPPING_COLUMNS.CATEGORY)).getValue()   // E
      
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
  const heroNames = sheet.getRange(DATA_START_ROW, 3, count, 1).getValues()  // C - Hero Name
  const heroIds = sheet.getRange(DATA_START_ROW, 4, count, 1).getValues()    // D - Hero ID
  const categories = sheet.getRange(DATA_START_ROW, 5, count, 1).getValues()  // E - Category
  
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
 * @returns {number} Номер строки
 */
function heroMapping_updateItem(itemName, heroName) {
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
  sheet.getRange(row, 1, 1, 5)  // 5 колонок (убрали AUTO_DETECTED)
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
  
  sheet.setRowHeight(row, ROW_HEIGHT)
}

/**
 * Обновляет или создает маппинги для нескольких предметов batch операцией
 * @param {Array<Object>} updates - Массив {itemName, heroName, imageUrl, category}
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
    // Используем category из update, если указан, иначе определяем по heroName
    const category = update.category || (heroName ? 'Hero Item' : 'Common Item')
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
          category  // E - Category
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
          category  // E - Category
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
    const categoryValues = []
    const imageRows = []
    const heroNameRows = []
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
      if (u.values[4] !== null) {  // Category
        categoryRows.push(u.row)
        categoryValues.push([u.values[4]])
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
    if (categoryValues.length > 0) {
      sheet.getRange(categoryRows[0], getColumnIndex(HERO_MAPPING_COLUMNS.CATEGORY), categoryValues.length, 1).setValues(categoryValues)
    }
  }
  
  // Batch запись новых строк
  if (newRows.length > 0) {
    const startRow = Math.max(lastRow + 1, DATA_START_ROW)
    const values = newRows.map(r => r.values)
    
    // Записываем все новые строки за раз
    sheet.getRange(startRow, 1, newRows.length, 5).setValues(values)  // 5 колонок (убрали AUTO_DETECTED)
    
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
 * Генерирует расширенный список ключей для поиска предмета
 * Учитывает "The", таунты, специальные символы
 * @param {string} itemName - Название предмета
 * @returns {Array<string>} Массив вариантов названия для поиска
 * @private
 */
function heroMapping_generateSearchKeys_(itemName) {
  const base = itemName.toLowerCase().trim()
  const keys = new Set()
  
  // Базовые варианты
  keys.add(base)
  keys.add(base.replace(/[''""`]/g, ''))
  keys.add(base.replace(/[''""`]/g, '').replace(/\s+/g, ' '))
  
  // Обработка "The" в начале
  if (base.startsWith('the ')) {
    const withoutThe = base.substring(4).trim()
    keys.add(withoutThe)
    keys.add(withoutThe.replace(/[''""`]/g, ''))
    keys.add(withoutThe.replace(/[''""`]/g, '').replace(/\s+/g, ' '))
  } else {
    const withThe = 'the ' + base
    keys.add(withThe)
    keys.add(withThe.replace(/[''""`]/g, ''))
    keys.add(withThe.replace(/[''""`]/g, '').replace(/\s+/g, ' '))
  }
  
  // Обработка таунтов (Taunt: *)
  if (base.includes('taunt:')) {
    const withoutTaunt = base.replace(/taunt:\s*/gi, '').trim()
    keys.add(withoutTaunt)
    keys.add(withoutTaunt.replace(/[''""`]/g, ''))
    keys.add(withoutTaunt.replace(/[''""`]/g, '').replace(/\s+/g, ' '))
  }
  
  // Удаление дефисов и специальных символов
  keys.add(base.replace(/[-–—]/g, ' ').replace(/\s+/g, ' ').trim())
  keys.add(base.replace(/[-–—'""`]/g, ' ').replace(/\s+/g, ' ').trim())
  
  return Array.from(keys).filter(k => k.length > 0)
}

/**
 * Улучшенное сопоставление названий предметов
 * Учитывает различные варианты написания, "The", таунты, специальные символы
 * @param {string} searchName - Название для поиска
 * @param {string} apiName - Название из API
 * @returns {boolean} true если названия совпадают
 * @private
 */
function heroMapping_matchItemNames_(searchName, apiName) {
  if (!searchName || !apiName) return false
  
  const normalize = (name) => {
    return name.toLowerCase().trim()
      .replace(/[''""`]/g, '')
      .replace(/\s+/g, ' ')
      .replace(/[-–—]/g, ' ')
      .trim()
  }
  
  const normalizeWithoutThe = (name) => {
    let normalized = normalize(name)
    if (normalized.startsWith('the ')) {
      normalized = normalized.substring(4).trim()
    }
    return normalized
  }
  
  const normalizeWithoutTaunt = (name) => {
    return normalize(name).replace(/taunt:\s*/gi, '').trim()
  }
  
  const searchNorm = normalize(searchName)
  const apiNorm = normalize(apiName)
  
  // Точное совпадение
  if (searchNorm === apiNorm) return true
  
  // Совпадение без "The"
  if (normalizeWithoutThe(searchNorm) === normalizeWithoutThe(apiNorm)) return true
  
  // Совпадение без таунта
  if (normalizeWithoutTaunt(searchNorm) === normalizeWithoutTaunt(apiNorm)) return true
  
  // Совпадение без "The" и таунта
  const searchNoTheNoTaunt = normalizeWithoutTaunt(normalizeWithoutThe(searchNorm))
  const apiNoTheNoTaunt = normalizeWithoutTaunt(normalizeWithoutThe(apiNorm))
  if (searchNoTheNoTaunt === apiNoTheNoTaunt) return true
  
  // Частичное совпадение (если одно название содержит другое)
  if (searchNorm.length > 10 && apiNorm.length > 10) {
    if (searchNorm.includes(apiNorm) || apiNorm.includes(searchNorm)) {
      return true
    }
  }
  
  return false
}

/**
 * Автоматическое определение героев через SteamWebAPI.ru (tag5)
 * Обрабатывает все предметы из History с batch операциями и fallback на item_by_nameid
 * Улучшено для обработки проблемных предметов (The, таунты, специальные символы)
 * @param {boolean} silent - Если true, не показывает UI сообщения и диалоги (для автоматических вызовов)
 * @param {boolean} autoSync - Если true, автоматически синхронизирует предметы из History перед определением героев
 */
function heroMapping_autoDetectFromSteamWebAPI(silent = false, autoSync = false) {
  console.log('HeroMapping: начало автоматического определения героев')
  
  const historySheet = getHistorySheet_()
  if (!historySheet) {
    const message = 'Ошибка: лист History не найден'
    console.error(`HeroMapping: ${message}`)
    if (!silent) {
      try {
        SpreadsheetApp.getUi().alert(message)
      } catch (e) {
        // UI недоступен (вызов из триггера)
      }
    }
    return
  }
  
  const lastRow = historySheet.getLastRow()
  if (lastRow < DATA_START_ROW) {
    const message = 'Нет предметов в History для автоопределения'
    console.log(`HeroMapping: ${message}`)
    if (!silent) {
      try {
        SpreadsheetApp.getUi().alert(message)
      } catch (e) {
        // UI недоступен (вызов из триггера)
      }
    }
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
  
  // Автоматическая синхронизация предметов из History (если запрошено)
  if (autoSync) {
    console.log('HeroMapping: автоматическая синхронизация предметов из History...')
    heroMapping_syncWithHistory(true) // silent = true
    console.log('HeroMapping: предметы синхронизированы')
  }
  
  // Диалог подтверждения (только если не silent режим и не autoSync)
  if (!silent && !autoSync) {
    try {
      const ui = SpreadsheetApp.getUi()
      const confirmResponse = ui.alert(
        'Автоопределение героев',
        `Будет обработано ${uniqueItemNames.length} предметов.\nЭто может занять несколько минут.\n\nПродолжить?`,
        ui.ButtonSet.YES_NO
      )
      
      if (confirmResponse !== ui.Button.YES) {
        return
      }
    } catch (e) {
      // UI недоступен (вызов из триггера), продолжаем выполнение
      console.log('HeroMapping: UI недоступен, продолжаем автоматическое выполнение')
    }
  }
  
  // Получаем данные через SteamWebAPI.ru пакетами (с автоматическим fallback)
  const batchResult = steamWebAPI_fetchItemsBatch(uniqueItemNames, 'dota2', true)
  
  // Обрабатываем результаты (fallback уже выполнен в steamWebAPI_fetchItemsBatch)
  const updates = []
  const notFoundItems = []
  
  uniqueItemNames.forEach(itemName => {
    // Пробуем найти в результате (включая результаты fallback)
    let itemData = null
    
    // Создаем расширенный список вариантов поиска с учетом проблемных названий
    const searchKeys = heroMapping_generateSearchKeys_(itemName)
    
    // Ищем в batchResult.items (объект с ключами из normalizedname/marketname/markethashname)
    if (batchResult.items) {
      for (const key of searchKeys) {
        if (batchResult.items[key]) {
          itemData = batchResult.items[key]
          break
        }
      }
      
      // Если не нашли по ключам, ищем по markethashname в значениях с улучшенным сопоставлением
      if (!itemData) {
        for (const itemKey in batchResult.items) {
          const item = batchResult.items[itemKey]
          const itemNames = [
            item.markethashname,
            item.marketname,
            item.normalizedname
          ].filter(n => n && n.trim().length > 0)
          
          for (const itemNameFromAPI of itemNames) {
            // Используем улучшенное сопоставление
            if (heroMapping_matchItemNames_(itemName, itemNameFromAPI)) {
              itemData = item
              break
            }
          }
          
          if (itemData) break
        }
      }
    }
    
    // Если не нашли, проверяем ручной маппинг для проблемных предметов
    if (!itemData) {
      notFoundItems.push(itemName)
      
      // Проверяем ручной маппинг для проблемных предметов
      const manualMapping = PROBLEMATIC_ITEMS_MANUAL_MAPPING[itemName]
      if (manualMapping) {
        console.log(`HeroMapping: найден ручной маппинг для "${itemName}" -> ${manualMapping.heroName || 'Common Item'}`)
        updates.push({
          itemName: itemName,
          heroName: manualMapping.heroName,
          imageUrl: '',
          category: manualMapping.category
        })
        return
      }
      
      // Если нет ручного маппинга, добавляем без героя
      updates.push({
        itemName: itemName,
        heroName: null,
        imageUrl: ''
      })
      return
    }
    
    // Парсим данные
    const parsedData = steamWebAPI_parseItemData(itemData)
    const heroName = steamWebAPI_getHeroNameFromTags(parsedData)
    const imageUrl = parsedData.imageUrl || ''
    
    updates.push({
      itemName: itemName,
      heroName: heroName || null,
      imageUrl: imageUrl
    })
  })
  
  // Логируем не найденные предметы (fallback уже был выполнен в steamWebAPI_fetchItemsBatch)
  if (notFoundItems.length > 0) {
    console.log(`HeroMapping: ${notFoundItems.length} предметов не найдено даже после fallback: ${notFoundItems.slice(0, 5).join(', ')}${notFoundItems.length > 5 ? '...' : ''}`)
  }
  
  // Batch обновление всех предметов
  heroMapping_updateItemsBatch_(updates)
  
  const detectedCount = updates.filter(u => u.heroName).length
  const skippedCount = updates.filter(u => !u.heroName).length
  
  console.log(`HeroMapping: автоматическое определение завершено. Определено: ${detectedCount}, пропущено: ${skippedCount}`)
  
  try {
    logAutoAction_('HeroMapping', 'Автоопределение героев', `OK (определено: ${detectedCount}, пропущено: ${skippedCount})`)
    if (!silent) {
      try {
        SpreadsheetApp.getUi().alert(`Автоопределение завершено!\n\nОпределено: ${detectedCount}\nПропущено: ${skippedCount}`)
      } catch (e) {
        // UI недоступен (вызов из триггера)
      }
    }
  } catch (e) {
    console.log('HeroMapping: ошибка при логировании')
  }
}

/**
 * Синхронизирует предметы из History с HeroMapping
 * Добавляет только новые предметы, если их еще нет (batch операция)
 * @param {boolean} silent - Если true, не показывает UI сообщения (для автоматических вызовов)
 */
function heroMapping_syncWithHistory(silent = false) {
  const historySheet = getHistorySheet_()
  if (!historySheet) {
    const message = 'Ошибка: лист History не найден'
    console.error(`HeroMapping: ${message}`)
    if (!silent) {
      try {
        SpreadsheetApp.getUi().alert(message)
      } catch (e) {
        // UI недоступен (вызов из триггера)
      }
    }
    return
  }
  
  const mappingSheet = getOrCreateHeroMappingSheet_()
  const historyLastRow = historySheet.getLastRow()
  
  if (historyLastRow < DATA_START_ROW) {
    const message = 'Нет предметов в History для синхронизации'
    console.log(`HeroMapping: ${message}`)
    if (!silent) {
      try {
        SpreadsheetApp.getUi().alert(message)
      } catch (e) {
        // UI недоступен (вызов из триггера)
      }
    }
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
    const message = 'Все предметы из History уже синхронизированы'
    console.log(`HeroMapping: ${message}`)
    if (!silent) {
      try {
        SpreadsheetApp.getUi().alert(message)
      } catch (e) {
        // UI недоступен (вызов из триггера)
      }
    }
    return
  }
  
  // Batch добавление новых предметов
  const updates = newItems.map(itemName => ({
    itemName: itemName,
    heroName: null,
    imageUrl: ''
  }))
  
  heroMapping_updateItemsBatch_(updates)
  
  console.log(`HeroMapping: добавлено ${newItems.length} новых предметов из History`)
  
  try {
    logAutoAction_('HeroMapping', 'Синхронизация с History', `Добавлено: ${newItems.length} новых предметов`)
    if (!silent) {
      try {
        SpreadsheetApp.getUi().alert(`Синхронизация завершена!\n\nДобавлено новых предметов: ${newItems.length}`)
      } catch (e) {
        // UI недоступен (вызов из триггера)
      }
    }
  } catch (e) {
    console.log('HeroMapping: ошибка при логировании')
  }
}

/**
 * Заполняет пустые Hero ID в HeroMapping, используя данные из HeroStats
 * Ищет Hero ID по Hero Name в HeroStats и заполняет в HeroMapping
 * @returns {Object} {filled: number, notFound: number} - количество заполненных и не найденных
 */
function heroMapping_fillMissingHeroIds() {
  const mappingSheet = getHeroMappingSheet_()
  if (!mappingSheet) {
    console.warn('HeroMapping: лист HeroMapping не найден')
    return {filled: 0, notFound: 0}
  }
  
  const statsSheet = getHeroStatsSheet_()
  if (!statsSheet) {
    console.warn('HeroMapping: лист HeroStats не найден, пропускаем заполнение Hero ID')
    return {filled: 0, notFound: 0}
  }
  
  const mappingLastRow = mappingSheet.getLastRow()
  if (mappingLastRow < DATA_START_ROW) {
    console.log('HeroMapping: нет предметов в HeroMapping')
    return {filled: 0, notFound: 0}
  }
  
  const statsLastRow = statsSheet.getLastRow()
  if (statsLastRow < DATA_START_ROW) {
    console.log('HeroMapping: нет данных в HeroStats, пропускаем заполнение Hero ID')
    return {filled: 0, notFound: 0}
  }
  
  // Создаем Map: heroName → heroId из HeroStats
  const heroNameToIdMap = new Map()
  const statsHeroIds = statsSheet.getRange(DATA_START_ROW, getColumnIndex(HERO_STATS_COLUMNS.HERO_ID), statsLastRow - HEADER_ROW, 1).getValues()
  const statsHeroNames = statsSheet.getRange(DATA_START_ROW, getColumnIndex(HERO_STATS_COLUMNS.HERO_NAME), statsLastRow - HEADER_ROW, 1).getValues()
  
  for (let i = 0; i < statsHeroIds.length; i++) {
    const heroId = Number(statsHeroIds[i][0])
    const heroName = String(statsHeroNames[i][0] || '').trim()
    
    if (heroId && heroName && !heroNameToIdMap.has(heroName)) {
      // Берем первый найденный Hero ID (одинаковый для High Rank и All Ranks)
      heroNameToIdMap.set(heroName, heroId)
    }
  }
  
  console.log(`HeroMapping: создан справочник Hero Name → Hero ID (${heroNameToIdMap.size} героев)`)
  
  // Находим строки в HeroMapping с пустым Hero ID, но с заполненным Hero Name
  const mappingItemNames = mappingSheet.getRange(DATA_START_ROW, getColumnIndex(HERO_MAPPING_COLUMNS.ITEM_NAME), mappingLastRow - HEADER_ROW, 1).getValues()
  const mappingHeroNames = mappingSheet.getRange(DATA_START_ROW, getColumnIndex(HERO_MAPPING_COLUMNS.HERO_NAME), mappingLastRow - HEADER_ROW, 1).getValues()
  const mappingHeroIds = mappingSheet.getRange(DATA_START_ROW, getColumnIndex(HERO_MAPPING_COLUMNS.HERO_ID), mappingLastRow - HEADER_ROW, 1).getValues()
  
  const heroIdUpdates = []
  let filledCount = 0
  let notFoundCount = 0
  
  for (let i = 0; i < mappingItemNames.length; i++) {
    const heroName = String(mappingHeroNames[i][0] || '').trim()
    const currentHeroId = mappingHeroIds[i][0]
    
    // Пропускаем, если Hero ID уже заполнен
    if (currentHeroId && Number.isFinite(Number(currentHeroId)) && Number(currentHeroId) > 0) {
      continue
    }
    
    // Пропускаем, если Hero Name пустой
    if (!heroName) {
      continue
    }
    
    // Ищем Hero ID по Hero Name
    const heroId = heroNameToIdMap.get(heroName)
    
    if (heroId) {
      const row = DATA_START_ROW + i
      heroIdUpdates.push({row: row, heroId: heroId})
      filledCount++
    } else {
      notFoundCount++
      console.log(`HeroMapping: не найден Hero ID для "${heroName}"`)
    }
  }
  
  // Batch-запись Hero ID
  if (heroIdUpdates.length > 0) {
    const heroIdCol = getColumnIndex(HERO_MAPPING_COLUMNS.HERO_ID)
    const updates = heroIdUpdates.map(u => [u.heroId])
    const rows = heroIdUpdates.map(u => u.row)
    
    // Группируем по последовательным строкам для оптимизации
    heroIdUpdates.forEach(update => {
      mappingSheet.getRange(update.row, heroIdCol).setValue(update.heroId)
    })
    
    console.log(`HeroMapping: заполнено ${filledCount} Hero ID из HeroStats`)
  }
  
  if (notFoundCount > 0) {
    console.log(`HeroMapping: не найдено Hero ID для ${notFoundCount} героев`)
  }
  
  return {filled: filledCount, notFound: notFoundCount}
}

