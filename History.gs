// History module
// Используем константы из Constants.gs
const HISTORY_CONFIG = {
  STEAM_APPID: STEAM_APP_ID,
  COLUMNS: HISTORY_COLUMNS,
}

// Форматирование листа
function history_formatTable() {
  const sheet = getOrCreateHistorySheet_()
  const lastRow = sheet.getLastRow()

  // Используем константы для заголовков
  const headers = HEADERS.HISTORY
  sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers])

  formatHeaderRange_(sheet.getRange(HEADER_ROW, 1, 1, headers.length))

  sheet.setRowHeight(HEADER_ROW, HEADER_HEIGHT)
  if (lastRow > 1) sheet.setRowHeights(DATA_START_ROW, lastRow - 1, ROW_HEIGHT)

  sheet.setColumnWidth(1, 150) // A - Image
  sheet.setColumnWidth(2, 250) // B - Name
  sheet.setColumnWidth(3, 80)  // C - Status
  sheet.setColumnWidth(4, 100) // D - Link
  sheet.setColumnWidth(5, 100) // E - Buy
  sheet.setColumnWidth(6, 120) // F - Текущая цена
  sheet.setColumnWidth(7, 100) // G - Min
  sheet.setColumnWidth(8, 100) // H - Max
  sheet.setColumnWidth(9, 150)  // I - Trend (объединенный, шире)
  sheet.setColumnWidth(10, 120) // J - Фаза (было K)
  sheet.setColumnWidth(11, 100) // K - Потенциал (было L)
  sheet.setColumnWidth(12, 130) // L - Рекомендация (было M)

  if (lastRow > 1) {
    sheet
      .getRange(2, 1, lastRow - 1, 12)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
    sheet.getRange(`B2:B${lastRow}`).setHorizontalAlignment('left')
    // Форматирование числовых колонок
    sheet.getRange(`F2:H${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY)
    // Форматирование колонки Потенциал (L) как процент с знаком "+"
    const potentialCol = getColumnIndex(HISTORY_COLUMNS.POTENTIAL)
    sheet.getRange(DATA_START_ROW, potentialCol, lastRow - 1, 1)
      .setNumberFormat('+0%;-0%;"—"')
  }

  sheet.setFrozenRows(HEADER_ROW)
  // Дополнительно форматируем все существующие колонки дат (M и далее, было N)
  history_formatAllDateColumns_(sheet)
  // Выделяем минимум и максимум (только визуальное форматирование)
  history_highlightMinMax_(sheet)
  
  // Примечание: Обновление аналитики (тренды, текущая цена, min/max) НЕ вызывается при форматировании,
  // так как это ручная операция только для настройки внешнего вида таблицы.
  // Аналитика обновляется автоматически после сбора цен или вручную через меню "Обновить аналитику".
  // Добавляем колонку кнопки «Купить» в E, если её нет
  const buyHeader = 'Купить'
  const colCount = sheet.getLastColumn()
  let needInsertBuyCol = false
  if (colCount < 5) {
    needInsertBuyCol = true
  } else {
    const e1 = sheet.getRange(1, 5)
    const eVal = String(e1.getValue() || '').trim()
    const eDisp = String(e1.getDisplayValue() || '').trim()
    const looksLikeDate = e1.getValue() instanceof Date || /^\d{2}\.\d{2}\.\d{2}$/.test(eDisp)
    if (eVal === buyHeader) {
      // уже есть — ничего не делаем
    } else if (looksLikeDate || eVal === '') {
      needInsertBuyCol = true
    } else if (eVal !== buyHeader) {
      // другая шапка — тоже вставим buy перед датами
      needInsertBuyCol = true
    }
  }
  if (needInsertBuyCol) sheet.insertColumnBefore(5)
  const buyHeaderCell = sheet.getRange(1, 5)
  buyHeaderCell.setValue(buyHeader)
  formatHeaderRange_(buyHeaderCell)
  const lastRow2 = sheet.getLastRow()
  if (lastRow2 > 1) {
    const rng = sheet.getRange(DATA_START_ROW, 5, lastRow2 - 1, 1)
    rng.insertCheckboxes()
    rng.setHorizontalAlignment('center')
    sheet.setColumnWidth(5, 100)
  }
  try {
    SpreadsheetApp.getUi().alert('Форматирование завершено (History)')
  } catch (e) {
    console.log('History: невозможно показать UI в данном контексте')
  }
  // Безопасная очистка/установка правил условного форматирования при пустом листе
  if (lastRow2 <= 1) {
    sheet.setConditionalFormatRules([])
  } else {
    // Применяем все правила форматирования в одном месте (тренды + аналитика)
    history_applyAllConditionalFormatting_(sheet)
  }
}

// Обновляет всю аналитику History: текущая цена, min/max, тренды, форматирование
function history_updateAllAnalytics_() {
  const sheet = getOrCreateHistorySheet_()
  history_updateCurrentPriceMinMax_(sheet)
  history_updateTrends()
  // ВАЖНО: Сначала применяем условное форматирование (для трендов, фаз, рекомендаций),
  // затем выделение min/max, чтобы оно не перезаписывалось условным форматированием
  history_applyAllConditionalFormatting_(sheet)
  history_highlightMinMax_(sheet)
}

// Гарантировать наличие колонки для текущего периода (ночь/день)
function history_ensurePeriodColumn(period) {
  const sheet = getOrCreateHistorySheet_()
  const now = new Date()
  const hour = now.getHours()
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yy')
  const periodLabel = period === PRICE_COLLECTION_PERIODS.MORNING ? 'ночь' : 'день'
  const headerDisplay = `${todayStr} ${periodLabel}`

  const firstDateCol = HISTORY_COLUMNS.FIRST_DATE_COL
  const lastCol = sheet.getLastColumn()
  
  // Ищем существующую колонку для этого периода
  if (lastCol >= firstDateCol) {
    const width = lastCol - firstDateCol + 1
    const dateRow = sheet.getRange(HEADER_ROW, firstDateCol, 1, width).getDisplayValues()[0]
    
    for (let i = 0; i < dateRow.length; i++) {
      const header = String(dateRow[i] || '').trim()
      if (header === headerDisplay) {
        const col = firstDateCol + i
        history_formatPriceColumn_(sheet, col)
        return col
      }
    }
  }

  // Проверка времени перед созданием колонки
  const minutes = now.getMinutes()
  const minutesStr = minutes < 10 ? '0' + minutes : String(minutes)
  const currentTimeMinutes = hour * 60 + minutes
  const eveningStartMinutes = UPDATE_INTERVALS.EVENING_HOUR * 60 + UPDATE_INTERVALS.EVENING_MINUTE
  const morningStartMinutes = UPDATE_INTERVALS.MORNING_HOUR * 60 + UPDATE_INTERVALS.MORNING_MINUTE
  
  // Проверка времени для колонки "день": должна создаваться только после 12:00
  if (period === PRICE_COLLECTION_PERIODS.EVENING && currentTimeMinutes < eveningStartMinutes) {
    console.log(`History: попытка создать колонку "день" преждевременно (текущее время: ${hour}:${minutesStr}, требуется >= ${UPDATE_INTERVALS.EVENING_HOUR}:${UPDATE_INTERVALS.EVENING_MINUTE.toString().padStart(2, '0')})`)
    throw new Error(`Колонка "день" может быть создана только после ${UPDATE_INTERVALS.EVENING_HOUR}:${UPDATE_INTERVALS.EVENING_MINUTE.toString().padStart(2, '0')}. Текущее время: ${hour}:${minutesStr}`)
  }
  
  // Проверка времени для колонки "ночь": должна создаваться с 00:00 до 12:00
  // Примечание: триггер срабатывает в 00:00, но проверка времени была строгой (00:10)
  // Изменено: разрешаем создание с 00:00, так как триггер не может быть настроен на точные минуты
  if (period === PRICE_COLLECTION_PERIODS.MORNING) {
    // Разрешаем создание колонки "ночь" с 00:00 до 12:00 (вместо строго 00:10)
    // Это предотвращает ошибки, когда триггер срабатывает в 00:00
    const isWithinMorningPeriod = (hour === 0 && minutes >= 0) || (hour > 0 && hour < UPDATE_INTERVALS.EVENING_HOUR) || (hour === UPDATE_INTERVALS.EVENING_HOUR && minutes === 0)
    if (!isWithinMorningPeriod) {
      console.log(`History: попытка создать колонку "ночь" вне допустимого времени (текущее время: ${hour}:${minutesStr}, требуется 00:00-12:00)`)
      throw new Error(`Колонка "ночь" может быть создана только с 00:00 до 12:00. Текущее время: ${hour}:${minutesStr}`)
    }
  }

  // Если создаем колонку "день", сначала проверяем наличие "ночи" за сегодня
  // Порядок должен быть: сначала ночь, потом день
  if (period === PRICE_COLLECTION_PERIODS.EVENING) {
    const nightLabel = `${todayStr} ночь`
    let nightColExists = false
    let nightColIndex = -1
    
    if (lastCol >= firstDateCol) {
      const width = lastCol - firstDateCol + 1
      const dateRow = sheet.getRange(HEADER_ROW, firstDateCol, 1, width).getDisplayValues()[0]
      
      for (let i = 0; i < dateRow.length; i++) {
        const header = String(dateRow[i] || '').trim()
        if (header === nightLabel) {
          nightColExists = true
          nightColIndex = firstDateCol + i
          break
        }
      }
    }
    
    // Если "ночи" нет - создаем её первой
    if (!nightColExists) {
      const newNightCol = Math.max(lastCol + 1, firstDateCol)
      sheet.getRange(HEADER_ROW, newNightCol).setValue(nightLabel)
      history_formatPriceColumn_(sheet, newNightCol)
      
      const nightHeader = sheet.getRange(HEADER_ROW, newNightCol)
      nightHeader.setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBackground(COLORS.BACKGROUND)
        .setFontWeight('bold')
        .setWrap(true)
      sheet.setColumnWidth(newNightCol, 100)
      
      // Обновляем lastCol после создания колонки "ночь"
      lastCol = sheet.getLastColumn()
    }
  }

  // Создаём новую колонку для текущего периода справа от последней
  const newCol = Math.max(lastCol + 1, firstDateCol)
  sheet.getRange(HEADER_ROW, newCol).setValue(headerDisplay)
  history_formatPriceColumn_(sheet, newCol)
  
  // Форматируем заголовок
  const header = sheet.getRange(HEADER_ROW, newCol)
  header.setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(COLORS.BACKGROUND)
    .setFontWeight('bold')
    .setWrap(true)
  sheet.setColumnWidth(newCol, 100)
  
  return newCol
}

// Создание столбца для текущего периода (обёртка для меню)
// Обратная совместимость: старое название сохранено
function history_ensureTodayColumn() {
  const period = getCurrentPricePeriod()
  return history_ensurePeriodColumn(period)
}

// Обновление цен для указанного периода
function history_updatePricesForPeriod(period) {
  const sheet = getOrCreateHistorySheet_()
  
  // Пытаемся получить или создать колонку периода
  // Если колонка не может быть создана (например, из-за проверки времени) - возвращаем false
  let periodCol
  try {
    periodCol = history_ensurePeriodColumn(period)
  } catch (e) {
    // Если колонка не может быть создана (например, преждевременная попытка) - просто возвращаем false
    // Это нормально для unified_priceUpdate, который может сработать до создания колонки
    console.log(`History: не удалось получить колонку периода ${period}: ${e.message}`)
    return false
  }

  const lastRow = sheet.getLastRow()
  if (lastRow <= 1 || !periodCol || periodCol < HISTORY_COLUMNS.FIRST_DATE_COL) return false

  const count = lastRow - 1
  const names = sheet.getRange(DATA_START_ROW, 2, count, 1).getValues()
  const periodVals = sheet.getRange(DATA_START_ROW, periodCol, count, 1).getValues()
  const statusVals = sheet.getRange(DATA_START_ROW, 3, count, 1).getValues()

  let updatedCount = 0
  let errorCount = 0
  let timeoutCount = 0
  let skippedCount = 0
  const startedAt = Date.now()

  for (let i = 0; i < count; i++) {
    if (shouldStopByTimeBudget_(startedAt, 330000)) {
      timeoutCount = count - i
      if (i > 0) {
        sheet.getRange(DATA_START_ROW, periodCol, i, 1).setValues(periodVals.slice(0, i))
        sheet.getRange(DATA_START_ROW, 3, i, 1).setValues(statusVals.slice(0, i))
      }
      break
    }
    const name = String(names[i][0] || '').trim()
    const hasValue = periodVals[i][0]
    if (!name || hasValue) {
      skippedCount++
      continue
    }

    try {
      const res = fetchLowestPriceWithBackoff_(HISTORY_CONFIG.STEAM_APPID, name, {
        attempts: 2,
        baseDelayMs: 200,
        betweenItemsMs: 150,
        timeBudgetMs: 330000,
        startedAt,
      })
      if (res && res.ok) {
        periodVals[i][0] = res.price
        statusVals[i][0] = STATUS.OK
        updatedCount++
      } else {
        statusVals[i][0] = STATUS.WARNING
        errorCount++
      }
    } catch (e) {
      console.error(`History: ошибка при обновлении ${name}:`, e)
      statusVals[i][0] = STATUS.WARNING
      errorCount++
    }
  }

  sheet.getRange(DATA_START_ROW, periodCol, count, 1).setValues(periodVals)
  sheet.getRange(DATA_START_ROW, 3, count, 1).setValues(statusVals)

  // Период считаем завершённым только если:
  // 1) Обработаны все строки (учитывая обновлённые и пропущенные),
  // 2) Не было таймаута,
  // 3) Все строки имеют значение (либо обновлено, либо уже было) - ошибки без значений считаются незавершенными
  //    Это гарантирует, что все строки имеют цену перед завершением периода.
  const allProcessed = (updatedCount + skippedCount + errorCount + timeoutCount) === count
  
  // Проверяем, что все строки действительно имеют значения (не только обновлены, но и не пустые)
  let allHaveValues = true
  let emptyCount = 0
  if (allProcessed && timeoutCount === 0) {
    const finalVals = sheet.getRange(DATA_START_ROW, periodCol, count, 1).getValues()
    for (let i = 0; i < finalVals.length; i++) {
      const val = finalVals[i][0]
      if (!val || (typeof val !== 'number') || isNaN(val) || val <= 0) {
        allHaveValues = false
        emptyCount++
      }
    }
  }
  
  // Если были пустые значения, логируем для диагностики
  if (emptyCount > 0) {
    console.log(`History: обнаружено ${emptyCount} пустых значений из ${count} строк. Период не завершён.`)
  }
  
  // Дополнительная защита от преждевременного завершения:
  // - Должна быть сделана хотя бы одна успешная запись
  // - Доля пустых значений равна нулю
  // - Не было таймаута
  const completed = (updatedCount > 0) && allProcessed && timeoutCount === 0 && allHaveValues

  console.log(`History: обновление завершено (период: ${period}). Обновлено: ${updatedCount}, Ошибок: ${errorCount}, Пропущено: ${skippedCount}, Прервано: ${timeoutCount}, Завершено: ${completed}`)

  return completed
}

// Функция для ручного обновления (через меню)
function history_updateAllPrices(isManualRun = false) {
  const period = getCurrentPricePeriod()
  return history_updatePricesForPeriod(period)
}

function history_updateSinglePrice(row, col) {
  const sheet = getOrCreateHistorySheet_()
  const name = sheet.getRange(`${HISTORY_CONFIG.COLUMNS.NAME}${row}`).getValue()

  try {
    const res = fetchLowestPrice_(HISTORY_CONFIG.STEAM_APPID, name)
    if (res.ok) {
      sheet.getRange(row, col).setValue(res.price)
      sheet.getRange(`${HISTORY_CONFIG.COLUMNS.STATUS}${row}`).setValue('✓')
      return 'updated'
    }
    sheet.getRange(`${HISTORY_CONFIG.COLUMNS.STATUS}${row}`).setValue('❌')
    return 'error'
  } catch (e) {
    console.error('History error:', e)
    sheet.getRange(`${HISTORY_CONFIG.COLUMNS.STATUS}${row}`).setValue('⚠️')
    return 'error'
  }
}

function history_updateImagesAndLinks() {
  const ui = SpreadsheetApp.getUi()
  const response = ui.alert(
    'Обновить изображения и ссылки (History)',
    'Как вы хотите обновить изображения и ссылки?\n\n' +
      'Да - Обновить все изображения и ссылки\n' +
      'Нет - Обновить только отсутствующие\n' +
      'Отмена - Отменить обновление',
    ui.ButtonSet.YES_NO_CANCEL
  )
  if (response === ui.Button.CANCEL) return

  const sheet = getOrCreateHistorySheet_()
  const lastRow = sheet.getLastRow()
  let updatedCount = 0
  let errorCount = 0

  if (lastRow > 1) {
    const count = lastRow - 1
    const names = sheet.getRange(DATA_START_ROW, 2, count, 1).getValues() // B
    const imageFormulas = sheet.getRange(DATA_START_ROW, 1, count, 1).getFormulas() // A
    const linkFormulas = sheet.getRange(DATA_START_ROW, 4, count, 1).getFormulas() // D

    for (let i = 0; i < count; i++) {
      const name = String(names[i][0] || '').trim()
      const curImage = imageFormulas[i][0]
      const curLink = linkFormulas[i][0]
      const needsUpdate = (response === ui.Button.YES) || !curImage || !curLink
      if (!name || !needsUpdate) continue
      try {
        const built = buildImageAndLinkFormula_(HISTORY_CONFIG.STEAM_APPID, name)
        imageFormulas[i][0] = built.image
        linkFormulas[i][0] = built.link
        updatedCount++
        Utilities.sleep(100) // Оптимизировано: было 1000мс, стало 100мс (как в Invest/Sales)
      } catch (e) {
        console.error('History: ошибка при подготовке формул', i + 2, e)
        errorCount++
      }
    }

    sheet.getRange(DATA_START_ROW, 1, count, 1).setFormulas(imageFormulas)
    sheet.getRange(DATA_START_ROW, 4, count, 1).setFormulas(linkFormulas)
  }

  try {
    ui.alert(
      'History — результат обновления',
      `Обновлено строк: ${updatedCount}\nОшибок: ${errorCount}`,
      ui.ButtonSet.OK
    )
  } catch (e) {
    console.log('History: невозможно показать UI в данном контексте')
  }
}


function history_findDuplicates() {
  const sheet = getOrCreateHistorySheet_()
  const res = highlightDuplicatesByName_(sheet, 2, '#e3f2fd')
  SpreadsheetApp.getUi().alert(res.duplicates ? `Найдено повторов: ${res.duplicates}` : 'Повторов не найдено')
}

function history_highlightMinMax() {
  const sheet = getOrCreateHistorySheet_()
  try {
    const lastRow = sheet.getLastRow()
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('Нет данных для обработки')
      return
    }
    
    history_highlightMinMax_(sheet)
    SpreadsheetApp.getUi().alert('Выделение Min/Max обновлено')
  } catch (e) {
    console.error('History: ошибка при выделении min/max:', e)
    SpreadsheetApp.getUi().alert('Ошибка при выделении Min/Max')
  }
}


// Функция getOrCreateHistorySheet_ перенесена в SheetService.gs

// Применяет формат цен для одной колонки-датой (все строки ниже заголовка)
function history_formatPriceColumn_(sheet, colIndex) {
  const lastRow = sheet.getLastRow()
  if (lastRow <= 1 || colIndex < 8) return
  sheet.getRange(DATA_START_ROW, colIndex, lastRow - 1, 1).setNumberFormat('#,##0.00 ₽')
  sheet.getRange(DATA_START_ROW, colIndex, lastRow - 1, 1).setHorizontalAlignment('center')
  sheet.getRange(DATA_START_ROW, colIndex, lastRow - 1, 1).setVerticalAlignment('middle')
}

// Выделяет минимум и максимум в каждой строке History
// Гарантирует выделение только одной ячейки min и одной max (самое правое значение при совпадениях)
function history_highlightMinMax_(sheet) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  const firstDateCol = HISTORY_COLUMNS.FIRST_DATE_COL // 14 (колонка N)
  
  if (lastRow <= 1 || lastCol < firstDateCol) return

  // Сбрасываем все фоны в диапазоне цен целиком (более надежно)
  // Важно: сбрасываем только в диапазоне колонок с датами, не затрагивая другие колонки
  const priceDataRange = sheet.getRange(DATA_START_ROW, firstDateCol, lastRow - 1, lastCol - firstDateCol + 1)
  priceDataRange.setBackground(null)
  
  // Дополнительно убеждаемся, что нет фона в других колонках строк (на случай если где-то применялось форматирование строк)
  // Но не трогаем колонки с данными (A-M), так как там может быть другое форматирование

  let processedRows = 0
  let skippedRows = 0
  const highlights = [] // Массив для batch обновления

  // Получаем все данные за один раз
  const names = sheet.getRange(DATA_START_ROW, 2, lastRow - 1, 1).getValues()
  const priceDataWidth = lastCol - firstDateCol + 1
  const priceData = sheet.getRange(DATA_START_ROW, firstDateCol, lastRow - 1, priceDataWidth).getValues()

  // Обрабатываем каждую строку
  for (let i = 0; i < lastRow - 1; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue

    // Собираем все валидные цены с их позициями
    const priceCells = []
    
    for (let j = 0; j < priceData[i].length; j++) {
      const value = priceData[i][j]
      // Учитываем только валидные числовые значения > 0
      if (typeof value === 'number' && !isNaN(value) && value > 0) {
        priceCells.push({ row: i + DATA_START_ROW, col: j + firstDateCol, value })
      }
    }

    // Если меньше 2 цен, пропускаем строку
    if (priceCells.length < 2) {
      skippedRows++
      continue
    }

    // Находим минимум и максимум из собранных цен
    const prices = priceCells.map(cell => cell.value)
    const minPrice = Math.min(...prices)
    const maxPrice = Math.max(...prices)

    // Если минимум и максимум одинаковые, не выделяем
    if (minPrice === maxPrice) {
      skippedRows++
      continue
    }

    // Находим самые правые (актуальные) позиции для min и max
    // При совпадениях выбираем самую правую ячейку
    let rightmostMinCol = -1
    let rightmostMaxCol = -1
    let rightmostMinRow = -1
    let rightmostMaxRow = -1
    
    // Идем справа налево, находим первое (самое правое) вхождение
    for (let k = priceCells.length - 1; k >= 0; k--) {
      const cell = priceCells[k]
      
      // Для минимального значения
      if (cell.value === minPrice && rightmostMinCol === -1) {
        rightmostMinCol = cell.col
        rightmostMinRow = cell.row
      }
      
      // Для максимального значения
      if (cell.value === maxPrice && rightmostMaxCol === -1) {
        rightmostMaxCol = cell.col
        rightmostMaxRow = cell.row
      }
      
      // Если нашли оба значения, можно прервать
      if (rightmostMinCol !== -1 && rightmostMaxCol !== -1) {
        break
      }
    }

    // Добавляем только одно выделение min и одно max
    if (rightmostMinCol !== -1 && rightmostMinRow !== -1) {
      highlights.push({ row: rightmostMinRow, col: rightmostMinCol, color: '#ffcdd2' })
    }
    if (rightmostMaxCol !== -1 && rightmostMaxRow !== -1) {
      highlights.push({ row: rightmostMaxRow, col: rightmostMaxCol, color: '#c8e6c9' })
    }
    
    processedRows++
  }

  // Batch обновление всех выделений - группируем по цвету для оптимизации
  // Группируем ячейки по цвету, чтобы минимизировать количество запросов
  const highlightsByColor = {}
  highlights.forEach(highlight => {
    if (!highlightsByColor[highlight.color]) {
      highlightsByColor[highlight.color] = []
    }
    highlightsByColor[highlight.color].push({ row: highlight.row, col: highlight.col })
  })
  
  // Применяем цвета группами (для каждого цвета - один batch запрос)
  // Примечание: к сожалению, Google Sheets API не поддерживает установку разных цветов
  // для разных ячеек в одном диапазоне, поэтому оставляем цикл, но он уже оптимизирован
  // тем, что мы собираем все ячейки заранее
  highlights.forEach(highlight => {
    sheet.getRange(highlight.row, highlight.col).setBackground(highlight.color)
  })

  console.log(`History: выделение Min/Max завершено. Обработано строк: ${processedRows}, пропущено: ${skippedRows}, выделений: ${highlights.length}`)
}

// Обновляет столбцы Текущая цена, Min, Max в History
function history_updateCurrentPriceMinMax_(sheet = null) {
  if (!sheet) sheet = getOrCreateHistorySheet_()
  const lastRow = sheet.getLastRow()
  if (lastRow <= 1) return
  
  const lastCol = sheet.getLastColumn()
  const firstDateCol = HISTORY_COLUMNS.FIRST_DATE_COL // 14
  const currentPriceCol = getColumnIndex(HISTORY_COLUMNS.CURRENT_PRICE)
  const minPriceCol = getColumnIndex(HISTORY_COLUMNS.MIN_PRICE)
  const maxPriceCol = getColumnIndex(HISTORY_COLUMNS.MAX_PRICE)
  
  const count = lastRow - 1
  const names = sheet.getRange(DATA_START_ROW, 2, count, 1).getValues() // B
  const currentPrices = sheet.getRange(DATA_START_ROW, currentPriceCol, count, 1).getValues()
  const minPrices = sheet.getRange(DATA_START_ROW, minPriceCol, count, 1).getValues()
  const maxPrices = sheet.getRange(DATA_START_ROW, maxPriceCol, count, 1).getValues()
  
  // Получаем все цены за один раз для оптимизации
  const priceDataWidth = lastCol >= firstDateCol ? lastCol - firstDateCol + 1 : 0
  const priceData = priceDataWidth > 0 
    ? sheet.getRange(DATA_START_ROW, firstDateCol, count, priceDataWidth).getValues()
    : []
  
  const period = getCurrentPricePeriod()
  
  // ОПТИМИЗАЦИЯ: Читаем заголовки колонок дат один раз вместо вызова getHistoryPriceForPeriod_ в цикле
  const dateHeaders = priceDataWidth > 0
    ? sheet.getRange(HEADER_ROW, firstDateCol, 1, priceDataWidth).getDisplayValues()[0]
    : []
  
  const now = new Date()
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yy')
  const periodLabel = period === PRICE_COLLECTION_PERIODS.MORNING ? 'ночь' : 'день'
  const targetHeader = `${todayStr} ${periodLabel}`
  
  // Находим индекс колонки с текущим периодом
  let currentPeriodColIndex = -1
  for (let j = dateHeaders.length - 1; j >= 0; j--) {
    if (String(dateHeaders[j] || '').trim() === targetHeader) {
      currentPeriodColIndex = j
      break
    }
  }
  
  let updatedCount = 0
  
  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue
    
    // ОПТИМИЗАЦИЯ: Получаем текущую цену напрямую из уже прочитанных данных
    if (currentPeriodColIndex >= 0 && priceDataWidth > 0 && priceData[i] && priceData[i][currentPeriodColIndex]) {
      const price = priceData[i][currentPeriodColIndex]
      if (typeof price === 'number' && !isNaN(price) && price > 0) {
        currentPrices[i][0] = price
      } else {
        // Если цена за текущий период отсутствует, ищем последнюю заполненную цену
        let foundPrice = null
        for (let j = priceData[i].length - 1; j >= 0; j--) {
          const value = priceData[i][j]
          if (typeof value === 'number' && !isNaN(value) && value > 0) {
            foundPrice = value
            break
          }
        }
        currentPrices[i][0] = foundPrice
      }
    } else {
      currentPrices[i][0] = null
    }
    
    // Вычисляем Min и Max из всех цен
    const prices = []
    if (priceDataWidth > 0 && priceData[i]) {
      for (let j = 0; j < priceData[i].length; j++) {
        const value = priceData[i][j]
        if (typeof value === 'number' && !isNaN(value) && value > 0) {
          prices.push(value)
        }
      }
    }
    
    if (prices.length > 0) {
      minPrices[i][0] = Math.min(...prices)
      maxPrices[i][0] = Math.max(...prices)
    } else {
      minPrices[i][0] = null
      maxPrices[i][0] = null
    }
    
    updatedCount++
  }
  
  // Batch запись всех значений
  sheet.getRange(DATA_START_ROW, currentPriceCol, count, 1).setValues(currentPrices)
  sheet.getRange(DATA_START_ROW, minPriceCol, count, 1).setValues(minPrices)
  sheet.getRange(DATA_START_ROW, maxPriceCol, count, 1).setValues(maxPrices)
  
  console.log(`History: обновлено Текущая цена/Min/Max для ${updatedCount} строк`)
}

// Форматирует все существующие колонки дат (N и далее) для всех строк
// ОПТИМИЗИРОВАНО: использует batch-операции вместо цикла по колонкам
function history_formatAllDateColumns_(sheet) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  const firstDateCol = HISTORY_COLUMNS.FIRST_DATE_COL
  if (lastRow <= 1 || lastCol < firstDateCol) return
  
  const dateColsCount = lastCol - firstDateCol + 1
  const dataRowsCount = lastRow - 1
  
  // Batch-операции для всех колонок дат сразу
  if (dataRowsCount > 0) {
    // Форматирование всех колонок дат одним запросом (формат, выравнивание)
    const dateDataRange = sheet.getRange(DATA_START_ROW, firstDateCol, dataRowsCount, dateColsCount)
    dateDataRange.setNumberFormat('#,##0.00 ₽')
    dateDataRange.setHorizontalAlignment('center')
    dateDataRange.setVerticalAlignment('middle')
  }
  
  // Форматирование заголовков колонок дат
  if (dateColsCount > 0) {
    const headerRange = sheet.getRange(HEADER_ROW, firstDateCol, 1, dateColsCount)
    headerRange.setHorizontalAlignment('center')
    headerRange.setVerticalAlignment('middle')
    formatHeaderRange_(headerRange)
  }
  
  // Установка ширины колонок (можно сделать batch, но setColumnWidth принимает только одну колонку)
  // Оставляем цикл только для ширины, так как это быстрая операция
  for (let col = firstDateCol; col <= lastCol; col++) {
    sheet.setColumnWidth(col, 100)
  }
}


// ===== СИСТЕМА АНАЛИЗА ТРЕНДОВ =====

// Основная функция анализа тренда для строки
function history_analyzeTrend(row) {
  const sheet = getOrCreateHistorySheet_()
  const lastCol = sheet.getLastColumn()
  const firstDateCol = HISTORY_COLUMNS.FIRST_DATE_COL // 14
  
  if (lastCol < firstDateCol) return { trend: '🟪', daysChange: 0 }

  // Собираем цены для анализа, группируя по дате (игнорируя период ночь/день)
  // Используем Map для группировки: ключ - дата (строка dd.MM.yy), значение - последняя цена за этот день
  const pricesByDate = new Map()
  const dateHeaders = []
  
  // Собираем все цены с их датами и колонками
  const priceEntries = []
  for (let col = firstDateCol; col <= lastCol; col++) {
    const value = sheet.getRange(row, col).getValue()
    const headerDisplay = sheet.getRange(HEADER_ROW, col).getDisplayValue()
    if (typeof value === 'number' && !isNaN(value) && value > 0 && headerDisplay) {
      // Извлекаем дату из заголовка (формат: "dd.MM.yy ночь" или "dd.MM.yy день" или просто "dd.MM.yy")
      const headerStr = String(headerDisplay).trim()
      const dateMatch = headerStr.match(/^(\d{2}\.\d{2}\.\d{2})/)
      if (dateMatch) {
        const dateKey = dateMatch[1] // Ключ даты без периода
        // Определяем период (ночь или день) для правильной сортировки
        const isDay = headerStr.includes('день')
        const isNight = headerStr.includes('ночь')
        // Сохраняем все записи для последующей группировки
        priceEntries.push({
          dateKey,
          value,
          col,
          isDay,
          isNight,
          isAfter: isDay || (!isNight && !isDay) // день идет после ночи
        })
        if (!dateHeaders.includes(dateKey)) {
          dateHeaders.push(dateKey)
        }
      }
    }
  }
  
  // Группируем по дате, беря последнюю цену за день
  // Если колонки идут слева направо (старые -> новые), последняя колонка для даты = последняя цена
  // Сортируем записи по колонке (позиции), чтобы последняя колонка для каждой даты была последней
  priceEntries.sort((a, b) => a.col - b.col)
  
  // Теперь для каждой даты последняя запись в массиве = последняя цена за день
  for (const entry of priceEntries) {
    pricesByDate.set(entry.dateKey, entry.value)
  }

  // Преобразуем Map в массивы цен и дат (в хронологическом порядке)
  // ВАЖНО: Сортируем даты в хронологическом порядке перед созданием массивов
  const sortedDateKeys = dateHeaders.sort((a, b) => {
    // Сравниваем даты в формате dd.MM.yy
    const partsA = a.split('.')
    const partsB = b.split('.')
    if (partsA.length !== 3 || partsB.length !== 3) return 0
    
    const yearA = 2000 + parseInt(partsA[2], 10)
    const yearB = 2000 + parseInt(partsB[2], 10)
    if (yearA !== yearB) return yearA - yearB
    
    const monthA = parseInt(partsA[1], 10)
    const monthB = parseInt(partsB[1], 10)
    if (monthA !== monthB) return monthA - monthB
    
    const dayA = parseInt(partsA[0], 10)
    const dayB = parseInt(partsB[0], 10)
    return dayA - dayB
  })
  
  const prices = []
  const dates = []
  for (const dateKey of sortedDateKeys) {
    const price = pricesByDate.get(dateKey)
    if (price) {
      prices.push(price)
      // Преобразуем строку даты в объект Date для расчета разницы
      const dateParts = dateKey.split('.')
      if (dateParts.length === 3) {
        const day = parseInt(dateParts[0], 10)
        const month = parseInt(dateParts[1], 10) - 1 // месяцы в JS начинаются с 0
        const year = 2000 + parseInt(dateParts[2], 10) // предполагаем формат yy -> 20yy
        dates.push(new Date(year, month, day))
      } else {
        dates.push(new Date()) // fallback
      }
    }
  }

  if (prices.length < 2) return { trend: '🟪', daysChange: 0 }

  // Анализируем тренд с помощью 4 методов
  const methods = [
    history_simpleComparison_(prices),
    history_movingAverages_(prices),
    history_linearRegression_(prices),
    history_momentumAnalysis_(prices)
  ]

  // Голосование между методами
  const votes = { '🟩': 0, '🟥': 0, '🟨': 0, '🟪': 0 }
  methods.forEach(method => {
    if (votes.hasOwnProperty(method)) votes[method]++
  })

  // Определяем победителя с приоритетами при равенстве
  // Приоритет: падение > рост > боковик > неопределенность
  // Логика: защита капитала важнее упущенной прибыли
  let trend = '🟪'
  let maxVotes = 0
  const priorityOrder = ['🟥', '🟩', '🟨', '🟪']
  
  for (const trendType of priorityOrder) {
    if (votes[trendType] > maxVotes) {
      maxVotes = votes[trendType]
      trend = trendType
    }
  }

  // Вычисляем количество дней с последней смены тренда
  const daysChange = history_calculateDaysChange_(prices, dates, trend)

  return { trend, daysChange }
}

// Метод 1: Простое сравнение последних значений с адаптивным порогом
function history_simpleComparison_(prices) {
  if (prices.length < 2) return '🟪'
  
  const config = TREND_ANALYSIS_CONFIG.SIMPLE_COMPARISON
  const recent = prices.slice(-3) // последние 3 значения
  const first = recent[0]
  const last = recent[recent.length - 1]
  
  // Вычисляем волатильность для адаптивного порога
  const volatility = history_calculateVolatility_(recent)
  const adaptiveThreshold = config.BASE_THRESHOLD + (volatility * config.VOLATILITY_MULTIPLIER)
  
  const change = Math.abs((last - first) / first)
  
  if (change < adaptiveThreshold * config.SIDEWAYS_FACTOR) return '🟨' // Боковик
  return change > adaptiveThreshold ? (last > first ? '🟩' : '🟥') : '🟨'
}

// Вспомогательная функция для расчета волатильности
function history_calculateVolatility_(prices) {
  if (prices.length < 2) return 0
  let sumSquaredChanges = 0
  for (let i = 1; i < prices.length; i++) {
    const change = Math.abs((prices[i] - prices[i-1]) / prices[i-1])
    sumSquaredChanges += change * change
  }
  return Math.sqrt(sumSquaredChanges / (prices.length - 1))
}

// Метод 2: Скользящие средние с адаптивными порогами
function history_movingAverages_(prices) {
  if (prices.length < 4) return '🟪'
  
  const config = TREND_ANALYSIS_CONFIG.MOVING_AVERAGES
  const shortWindow = Math.min(config.SHORT_WINDOW, Math.floor(prices.length / 2))
  const longWindow = Math.min(config.LONG_WINDOW, prices.length)
  
  const shortMA = history_calculateMA_(prices, shortWindow)
  const longMA = history_calculateMA_(prices, longWindow)
  
  if (longMA === 0) return '🟪'
  
  const diff = (shortMA - longMA) / longMA
  const volatility = history_calculateVolatility_(prices)
  const adaptiveThreshold = config.BASE_THRESHOLD + (volatility * config.VOLATILITY_MULTIPLIER)
  
  if (diff > adaptiveThreshold) return '🟩'
  if (diff < -adaptiveThreshold) return '🟥'
  return '🟨'
}

// Метод 3: Линейная регрессия со скользящим окном
function history_linearRegression_(prices) {
  if (prices.length < 3) return '🟪'
  
  const config = TREND_ANALYSIS_CONFIG.LINEAR_REGRESSION
  // Используем последние N значений или всю выборку, если меньше
  const window = Math.min(config.WINDOW, prices.length)
  const recentPrices = prices.slice(-window)
  
  const n = recentPrices.length
  const x = Array.from({length: n}, (_, i) => i)
  const y = recentPrices
  
  const sumX = x.reduce((a, b) => a + b, 0)
  const sumY = y.reduce((a, b) => a + b, 0)
  const sumXY = x.reduce((sum, xi, i) => sum + xi * y[i], 0)
  const sumXX = x.reduce((sum, xi) => sum + xi * xi, 0)
  
  const slope = (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX)
  const avgPrice = sumY / n
  
  // Адаптивный порог на основе средней цены
  const relativeSlope = slope / avgPrice
  
  if (relativeSlope > config.GROWTH_THRESHOLD) return '🟩'
  if (relativeSlope < config.FALL_THRESHOLD) return '🟥'
  return '🟨'
}

// Метод 4: Momentum анализ с защитой от всплесков
function history_momentumAnalysis_(prices) {
  if (prices.length < 3) return '🟪'
  
  const config = TREND_ANALYSIS_CONFIG.MOMENTUM
  const recent = prices.slice(-Math.min(config.WINDOW, prices.length))
  const momentum = recent[recent.length - 1] - recent[0]
  const avgPrice = recent.reduce((a, b) => a + b, 0) / recent.length
  
  if (avgPrice === 0) return '🟪'
  
  const momentumPercent = momentum / avgPrice
  const volatility = history_calculateVolatility_(recent)
  
  // Адаптивный порог: чем выше волатильность, тем выше порог
  const adaptiveThreshold = config.BASE_THRESHOLD + (volatility * config.VOLATILITY_MULTIPLIER)
  
  if (momentumPercent > adaptiveThreshold) return '🟩'
  if (momentumPercent < -adaptiveThreshold) return '🟥'
  return '🟨'
}

// Вспомогательная функция для расчета скользящего среднего
function history_calculateMA_(prices, window) {
  if (prices.length < window) return prices.reduce((a, b) => a + b, 0) / prices.length
  
  const recent = prices.slice(-window)
  return recent.reduce((a, b) => a + b, 0) / recent.length
}

// Расчет количества дней с последней смены тренда
function history_calculateDaysChange_(prices, dates, currentTrend) {
  if (prices.length < 3) return 0
  if (dates.length < 2) return 0
  
  // Анализируем тренды для каждого периода, начиная с предпоследнего
  for (let i = prices.length - 2; i >= 1; i--) {
    const periodPrices = prices.slice(0, i + 1)
    const periodTrend = history_simpleComparison_(periodPrices)
    
    if (periodTrend !== currentTrend) {
      // Нашли смену тренда
      const changeDate = dates[i]
      const currentDate = dates[dates.length - 1]
      
      if (changeDate instanceof Date && currentDate instanceof Date) {
        const diffTime = Math.abs(currentDate - changeDate)
        const daysDiff = Math.ceil(diffTime / (1000 * 60 * 60 * 24))
        return daysDiff > 0 ? daysDiff : 1 // Минимум 1 день
      }
    }
  }
  
  // Если тренд не менялся, вычисляем разницу между первой и последней датой
  if (dates.length >= 2) {
    const firstDate = dates[0]
    const lastDate = dates[dates.length - 1]
    if (firstDate instanceof Date && lastDate instanceof Date) {
      const diffTime = Math.abs(lastDate - firstDate)
      const daysDiff = Math.ceil(diffTime / (1000 * 60 * 60 * 24))
      return daysDiff > 0 ? daysDiff : 1 // Минимум 1 день
    }
  }
  
  return 0 // Если не удалось вычислить
}

// Форматирует отображение дней смены с описанием тренда
// Примеры: "🟥 Падает 35 дн.", "🟩 Растет 12 дн.", "🟨 Боковик 5 дн."
function history_formatDaysChange_(trend, daysChange) {
  if (!daysChange || daysChange === 0) {
    return '—'
  }
  
  const trendLabels = {
    '🟥': 'Падает',
    '🟩': 'Растет',
    '🟨': 'Боковик',
    '🟪': 'Нет данных'
  }
  
  const label = trendLabels[trend] || 'Тренд'
  
  return `${trend} ${label} ${daysChange} д.`
}

// ОПТИМИЗИРОВАННАЯ версия анализа тренда - работает с уже прочитанными данными (без запросов к таблице)
function history_analyzeTrendFromPrices_(prices, dates) {
  if (prices.length < 2) return { trend: '🟪', daysChange: 0 }

  // Анализируем тренд с помощью 4 методов
  const methods = [
    history_simpleComparison_(prices),
    history_movingAverages_(prices),
    history_linearRegression_(prices),
    history_momentumAnalysis_(prices)
  ]

  // Подсчитываем голоса
  const votes = { '🟩': 0, '🟥': 0, '🟨': 0, '🟪': 0 }
  methods.forEach(trend => {
    if (votes.hasOwnProperty(trend)) {
      votes[trend]++
    }
  })

  // Определяем итоговый тренд
  let finalTrend = '🟪'
  const maxVotes = Math.max(votes['🟩'], votes['🟥'], votes['🟨'], votes['🟪'])
  
  // Приоритет при равенстве: 🟥 > 🟩 > 🟨 > 🟪
  if (votes['🟥'] === maxVotes) {
    finalTrend = '🟥'
  } else if (votes['🟩'] === maxVotes) {
    finalTrend = '🟩'
  } else if (votes['🟨'] === maxVotes) {
    finalTrend = '🟨'
  }

  // Расчет дней смены тренда
  const daysChange = history_calculateDaysChange_(prices, dates, finalTrend)

  return { trend: finalTrend, daysChange }
}

// Обновление трендов для всех строк (оптимизированная версия)
function history_updateTrends() {
  const sheet = getOrCreateHistorySheet_()
  const lastRow = sheet.getLastRow()
  if (lastRow <= 1) {
    try {
      SpreadsheetApp.getUi().alert('Нет данных для анализа трендов')
    } catch (e) {
      console.log('History: невозможно показать UI в данном контексте')
    }
    return
  }

  // Убедимся, что колонки для расширенной аналитики существуют
  history_ensureExtendedAnalyticsColumns_()

  const count = lastRow - 1
  const lastCol = sheet.getLastColumn()
  const firstDateCol = HISTORY_COLUMNS.FIRST_DATE_COL // 14
  
  // Определяем номера колонок для расширенной аналитики (K, L, M)
  const phaseCol = getColumnIndex(HISTORY_COLUMNS.PHASE)           // K
  const potentialCol = getColumnIndex(HISTORY_COLUMNS.POTENTIAL)   // L
  const recommendationCol = getColumnIndex(HISTORY_COLUMNS.RECOMMENDATION) // M
  
  // ОПТИМИЗАЦИЯ: Читаем все данные одним batch-запросом
  const names = sheet.getRange(DATA_START_ROW, 2, count, 1).getValues() // B
  const trends = sheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.TREND), count, 1).getValues() // I (теперь содержит объединенный тренд+дни)
  const phases = sheet.getRange(DATA_START_ROW, phaseCol, count, 1).getValues()
  const potentials = sheet.getRange(DATA_START_ROW, potentialCol, count, 1).getValues()
  const recommendations = sheet.getRange(DATA_START_ROW, recommendationCol, count, 1).getValues()
  
  // ОПТИМИЗАЦИЯ: Читаем все цены и заголовки колонок одним batch-запросом
  const priceDataWidth = lastCol >= firstDateCol ? lastCol - firstDateCol + 1 : 0
  const allPriceData = priceDataWidth > 0 
    ? sheet.getRange(DATA_START_ROW, firstDateCol, count, priceDataWidth).getValues()
    : []
  const allHeaders = priceDataWidth > 0
    ? sheet.getRange(HEADER_ROW, firstDateCol, 1, priceDataWidth).getDisplayValues()[0]
    : []

  let updatedCount = 0
  
  // Анализируем тренды для всех строк
  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue
    
    const row = i + 2
    
    // ОПТИМИЗАЦИЯ: Используем уже прочитанные данные вместо отдельных запросов
    const pricesByDate = new Map()
    const dateHeaders = []
    const priceEntries = []
    
    // Обрабатываем данные из уже прочитанного массива
    if (priceDataWidth > 0 && allPriceData[i]) {
      for (let j = 0; j < priceDataWidth; j++) {
        const value = allPriceData[i][j]
        const headerDisplay = allHeaders[j]
        if (typeof value === 'number' && !isNaN(value) && value > 0 && headerDisplay) {
          const headerStr = String(headerDisplay).trim()
          const dateMatch = headerStr.match(/^(\d{2}\.\d{2}\.\d{2})/)
          if (dateMatch) {
            const dateKey = dateMatch[1]
            const col = firstDateCol + j
            priceEntries.push({
              dateKey,
              value,
              col
            })
            if (!dateHeaders.includes(dateKey)) {
              dateHeaders.push(dateKey)
            }
          }
        }
      }
    }
    
    // Группируем по дате, беря последнюю цену за день
    priceEntries.sort((a, b) => a.col - b.col)
    for (const entry of priceEntries) {
      pricesByDate.set(entry.dateKey, entry.value)
    }
    
    // Сортируем даты в хронологическом порядке
    const sortedDateKeys = dateHeaders.sort((a, b) => {
      const partsA = a.split('.')
      const partsB = b.split('.')
      if (partsA.length !== 3 || partsB.length !== 3) return 0
      const yearA = 2000 + parseInt(partsA[2], 10)
      const yearB = 2000 + parseInt(partsB[2], 10)
      if (yearA !== yearB) return yearA - yearB
      const monthA = parseInt(partsA[1], 10)
      const monthB = parseInt(partsB[1], 10)
      if (monthA !== monthB) return monthA - monthB
      const dayA = parseInt(partsA[0], 10)
      const dayB = parseInt(partsB[0], 10)
      return dayA - dayB
    })
    
    // Создаем массив цен в хронологическом порядке (по одной цене на день)
    const prices = []
    const dates = []
    for (const dateKey of sortedDateKeys) {
      const price = pricesByDate.get(dateKey)
      if (price) {
        prices.push(price)
        // Преобразуем строку даты в объект Date для расчета разницы
        const dateParts = dateKey.split('.')
        if (dateParts.length === 3) {
          const day = parseInt(dateParts[0], 10)
          const month = parseInt(dateParts[1], 10) - 1 // месяцы в JS начинаются с 0
          const year = 2000 + parseInt(dateParts[2], 10) // предполагаем формат yy -> 20yy
          dates.push(new Date(year, month, day))
        } else {
          dates.push(new Date()) // fallback
        }
      }
    }
    
    // ОПТИМИЗАЦИЯ: Анализируем тренд используя уже собранные данные (без дополнительных запросов к таблице)
    const analysis = history_analyzeTrendFromPrices_(prices, dates)
    
    // Базовый анализ тренда - объединяем тренд и дни смены в одну колонку
    trends[i][0] = history_formatDaysChange_(analysis.trend, analysis.daysChange)
    
    // Расширенный анализ
    if (prices.length >= 7) {
      phases[i][0] = history_determineCyclePhase_(prices)
      const potential = history_calculateGrowthPotential_(prices)
      // Храним числовое значение для сортировки (в процентах, например 14 для +14%)
      potentials[i][0] = potential ? potential.to85th / 100 : null
      // Извлекаем чистый тренд и дни смены из объединенного формата для рекомендаций
      const trendStr = String(trends[i][0] || '')
      const trendMatch = trendStr.match(/^([🟥🟩🟨🟪])/)
      const daysMatch = trendStr.match(/(\d+)\s+дн?\.?/)
      const cleanTrend = trendMatch ? trendMatch[1] : '🟪'
      const daysChange = daysMatch ? parseInt(daysMatch[1], 10) : 0
      
      recommendations[i][0] = history_generateRecommendation_(
        phases[i][0],
        cleanTrend,
        potential,
        daysChange
      )
    } else {
      phases[i][0] = '❓'
      potentials[i][0] = null
      recommendations[i][0] = '👀 НАБЛЮДАТЬ'
    }
    
    updatedCount++
  }
  
  // Batch обновление всех данных за одну операцию
  const trendCol = getColumnIndex(HISTORY_COLUMNS.TREND)
  sheet.getRange(DATA_START_ROW, trendCol, count, 1).setValues(trends)
  sheet.getRange(DATA_START_ROW, phaseCol, count, 1).setValues(phases)
  sheet.getRange(DATA_START_ROW, potentialCol, count, 1).setValues(potentials)
  sheet.getRange(DATA_START_ROW, recommendationCol, count, 1).setValues(recommendations)
  
  // Применяем все правила условного форматирования (тренды + аналитика)
  history_applyAllConditionalFormatting_(sheet)
  
  console.log(`History: обновлено трендов: ${updatedCount}`)
  
  try {
    SpreadsheetApp.getUi().alert(
      'History — анализ трендов завершен',
      `Обновлено трендов: ${updatedCount}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    )
  } catch (e) {
    console.log('History: невозможно показать UI в данном контексте')
  }
}

// Убедиться что колонки для расширенной аналитики существуют (K, L, M)
function history_ensureExtendedAnalyticsColumns_() {
  const sheet = getOrCreateHistorySheet_()
  const lastRow = sheet.getLastRow()
  
  // Колонки K, L, M зарезервированы для Фаза, Потенциал, Рекомендация
  // Проверяем заголовки
  const phaseCol = getColumnIndex(HISTORY_COLUMNS.PHASE)
  const potentialCol = getColumnIndex(HISTORY_COLUMNS.POTENTIAL)
  const recommendationCol = getColumnIndex(HISTORY_COLUMNS.RECOMMENDATION)
  
  const phaseHeader = sheet.getRange(HEADER_ROW, phaseCol).getValue()
  const potentialHeader = sheet.getRange(HEADER_ROW, potentialCol).getValue()
  const recommendationHeader = sheet.getRange(HEADER_ROW, recommendationCol).getValue()
  
  // Если заголовки пустые или неверные - устанавливаем правильные
  if (!phaseHeader || phaseHeader !== 'Фаза') {
    sheet.getRange(HEADER_ROW, phaseCol).setValue('Фаза')
  }
  if (!potentialHeader || potentialHeader !== 'Потенциал (P85)') {
    sheet.getRange(HEADER_ROW, potentialCol).setValue('Потенциал (P85)')
  }
  if (!recommendationHeader || recommendationHeader !== 'Рекомендация') {
    sheet.getRange(HEADER_ROW, recommendationCol).setValue('Рекомендация')
  }
  
  // Форматируем заголовки
  const headerRange = sheet.getRange(HEADER_ROW, phaseCol, 1, 3)
  formatHeaderRange_(headerRange)
  
  // Ширины колонок
  sheet.setColumnWidth(phaseCol, 120)  // Фаза
  sheet.setColumnWidth(potentialCol, 100)  // Потенциал
  sheet.setColumnWidth(recommendationCol, 130) // Рекомендация
  
  // Центрирование данных и форматирование
  if (lastRow > 1) {
    sheet.getRange(DATA_START_ROW, phaseCol, lastRow - 1, 3)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
    // Форматируем колонку Потенциал как процент с знаком "+" для положительных значений
    sheet.getRange(DATA_START_ROW, potentialCol, lastRow - 1, 1)
      .setNumberFormat('+0%;-0%;"—"')
  }
}

// Применение всех правил условного форматирования (тренды + расширенная аналитика)
function history_applyAllConditionalFormatting_(sheet) {
  const lastRow = sheet.getLastRow()
  if (lastRow <= 1) return
  
  // Диапазоны для форматирования
  const trendCol = getColumnIndex(HISTORY_COLUMNS.TREND)
  const phaseCol = getColumnIndex(HISTORY_COLUMNS.PHASE)
  const recommendationCol = getColumnIndex(HISTORY_COLUMNS.RECOMMENDATION)
  
  applyAnalyticsFormatting_(sheet, {
    trendCol,
    phaseCol,
    recommendationCol
  }, lastRow)
}

// ===== РАСШИРЕННАЯ АНАЛИТИКА ДЛЯ DOTA 2 СТРАТЕГИИ =====

// Определение фазы цикла предмета
function history_determineCyclePhase_(prices) {
  const config = TREND_ANALYSIS_CONFIG.CYCLE_PHASE
  if (prices.length < config.MIN_DATA_POINTS) return '❓'
  
  const currentPrice = prices[prices.length - 1]
  const minPrice = Math.min(...prices)
  const maxPrice = Math.max(...prices)
  const priceRange = maxPrice - minPrice
  
  if (priceRange === 0) return '➡️ СЕРЕДИНА'
  
  // Процент от диапазона min-max
  const positionInRange = (currentPrice - minPrice) / priceRange
  
  // Анализ тренда последних 7 дней
  const recentPrices = prices.slice(-Math.min(7, prices.length))
  const recentTrend = history_linearRegression_(recentPrices)
  
  // Определяем фазу на основе позиции и тренда
  if (positionInRange < config.BOTTOM_THRESHOLD) {
    return '🟩 ДНО'  // В нижних 25%
  } else if (positionInRange < 0.5 && recentTrend === '🟩') {
    return '↗️ РОСТ'  // Средний диапазон + растет
  } else if (positionInRange > config.TOP_THRESHOLD) {
    return '🔥 ПИК'  // В верхних 25%
  } else if (positionInRange > 0.5 && recentTrend === '🟥') {
    return '↘️ КОРРЕКЦИЯ'  // Средний диапазон + падает
  }
  
  return '➡️ СЕРЕДИНА'  // Средний диапазон без явного тренда
}

// Расчет потенциала роста
function history_calculateGrowthPotential_(prices) {
  const config = TREND_ANALYSIS_CONFIG.GROWTH_POTENTIAL
  if (prices.length < config.MIN_DATA_POINTS) return null
  
  const currentPrice = prices[prices.length - 1]
  const maxPrice = Math.max(...prices)
  
  // Потенциал до исторического максимума
  const potentialToMax = ((maxPrice - currentPrice) / currentPrice) * 100
  
  // Реалистичный пик (85-й перцентиль цен)
  const sortedPrices = [...prices].sort((a, b) => a - b)
  const percentile85Index = Math.floor(sortedPrices.length * config.PERCENTILE_TARGET)
  const percentile85 = sortedPrices[percentile85Index]
  const potentialTo85th = ((percentile85 - currentPrice) / currentPrice) * 100
  
  return {
    toMax: Math.round(potentialToMax),
    to85th: Math.round(potentialTo85th)
  }
}

// Генерация рекомендации на основе анализа
function history_generateRecommendation_(phase, trend, potential, daysChange) {
  if (!potential) return '👀 НАБЛЮДАТЬ'
  
  // Используем to85th для более реалистичных рекомендаций
  const p85 = potential.to85th
  
  // Покупка на дне с растущим/боковым трендом и высоким потенциалом
  if (phase === '🟩 ДНО' && (trend === '🟩' || trend === '🟨') && p85 > 50) {
    return '🟩 КУПИТЬ'
  }
  
  // Держать при росте с хорошим потенциалом
  if (phase === '↗️ РОСТ' && trend === '🟩' && p85 > 30) {
    return '🟨 ДЕРЖАТЬ'
  }
  
  // Продавать на пике или при падении с малым потенциалом
  if (phase === '🔥 ПИК' || (trend === '🟥' && daysChange > 3 && p85 < 20)) {
    return '🟥 ПРОДАТЬ'
  }
  
  // Коррекция - ждать дна для покупки
  if (phase === '↘️ КОРРЕКЦИЯ' && trend === '🟥') {
    return '⏳ ЖДАТЬ ДНА'
  }
  
  // Средний диапазон без явного сигнала
  if (phase === '➡️ СЕРЕДИНА') {
    if (trend === '🟩' && p85 > 40) return '🟨 ДЕРЖАТЬ'
    if (trend === '🟥' && p85 < 15) return '🟥 ПРОДАТЬ'
  }
  
  return '👀 НАБЛЮДАТЬ'
}