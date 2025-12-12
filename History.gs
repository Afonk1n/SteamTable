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
  if (!headers || !Array.isArray(headers) || headers.length === 0) {
    console.error('History: HEADERS.HISTORY не определен или пуст')
    SpreadsheetApp.getUi().alert('Ошибка: HEADERS.HISTORY не определен в Constants.gs')
    return
  }
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
  sheet.setColumnWidth(9, 130) // I - Investment Score
  sheet.setColumnWidth(10, 130) // J - Рекомендация
  sheet.setColumnWidth(11, 120) // K - Фаза
  sheet.setColumnWidth(12, 100) // L - Потенциал
  sheet.setColumnWidth(13, 150) // M - Тренд (объединенный формат: "🟨 Боковик 39 д.", убрали колонку Дней смены)
  sheet.setColumnWidth(14, 100) // N - Hero Trend (перемещено из O)
  sheet.setColumnWidth(15, 120) // O - Contest Rate Change (7d) (перемещено из P)
  sheet.setColumnWidth(16, 120) // P - Contest Rate (current) (перемещено из Q)
  sheet.setColumnWidth(17, 100) // Q - Pick Rate (current) (перемещено из R)
  sheet.setColumnWidth(18, 100) // R - Win Rate (current) (перемещено из S)
  sheet.setColumnWidth(19, 150) // S - Hero Name (перемещено из T)

  if (lastRow > 1) {
    const dataCols = 19 // Количество колонок с данными (до дат, было 20)
    sheet
      .getRange(2, 1, lastRow - 1, dataCols)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
    sheet.getRange(`B2:B${lastRow}`).setHorizontalAlignment('left')
    sheet.getRange(`T2:T${lastRow}`).setHorizontalAlignment('left') // Hero Name - выравнивание влево
    // Форматирование числовых колонок
    sheet.getRange(`F2:H${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY)
    // Форматирование колонки Потенциал (L) как процент с знаком "+"
    const potentialCol = getColumnIndex(HISTORY_COLUMNS.POTENTIAL)
    sheet.getRange(DATA_START_ROW, potentialCol, lastRow - 1, 1)
      .setNumberFormat('+0%;-0%;"—"')
    // Форматирование колонок статистики героя:
    // O (Pro Contest Rate current) - процент контест-рейта про-сцены
    const proContestRateCol = getColumnIndex(HISTORY_COLUMNS.PRO_CONTEST_RATE_CURRENT)
    sheet.getRange(DATA_START_ROW, proContestRateCol, lastRow - 1, 1)
      .setNumberFormat(NUMBER_FORMATS.PERCENT)
    // P (Pro Contest Rate Change 7d) - процент изменения
    const proContestRateChangeCol = getColumnIndex(HISTORY_COLUMNS.PRO_CONTEST_RATE_CHANGE_7D)
    sheet.getRange(DATA_START_ROW, proContestRateChangeCol, lastRow - 1, 1)
      .setNumberFormat(NUMBER_FORMATS.PERCENT)
    // Q (Pick Rate Change Immortal 7d) - процент изменения за неделю
    const pickRateChange7dCol = getColumnIndex(HISTORY_COLUMNS.PICK_RATE_CHANGE_IMMORTAL_7D)
    sheet.getRange(DATA_START_ROW, pickRateChange7dCol, lastRow - 1, 1)
      .setNumberFormat(NUMBER_FORMATS.PERCENT)
    // R (Pick Rate Change Immortal 24h) - процент изменения за 24ч
    const pickRateChange24hCol = getColumnIndex(HISTORY_COLUMNS.PICK_RATE_CHANGE_IMMORTAL_24H)
    sheet.getRange(DATA_START_ROW, pickRateChange24hCol, lastRow - 1, 1)
      .setNumberFormat(NUMBER_FORMATS.PERCENT)
    // S (Pick Rate Immortal) - процент пиков Immortal
    const pickRateCol = getColumnIndex(HISTORY_COLUMNS.PICK_RATE_IMMORTAL)
    sheet.getRange(DATA_START_ROW, pickRateCol, lastRow - 1, 1)
      .setNumberFormat(NUMBER_FORMATS.PERCENT)
    // T (Win Rate current) - процент
    const winRateCol = getColumnIndex(HISTORY_COLUMNS.WIN_RATE_CURRENT)
    sheet.getRange(DATA_START_ROW, winRateCol, lastRow - 1, 1)
      .setNumberFormat(NUMBER_FORMATS.PERCENT)
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
  console.log('History: форматирование завершено')
  // Безопасная очистка/установка правил условного форматирования при пустом листе
  if (lastRow2 <= 1) {
    sheet.setConditionalFormatRules([])
  } else {
    // Применяем все правила форматирования в одном месте (тренды + аналитика)
    history_applyAllConditionalFormatting_(sheet)
  }
}

// Обновляет всю аналитику History: текущая цена, min/max, тренды, форматирование
// @param {boolean} skipHeroStats - Если true, пропускает синхронизацию статистики героев (для оптимизации setup)
function history_updateAllAnalytics_(skipHeroStats = false) {
  const sheet = getOrCreateHistorySheet_()
  history_updateCurrentPriceMinMax_(sheet)
  history_updateTrends()
  // Синхронизация статистики героев (колонки O-T) - пропускаем если уже выполнена в setup
  if (!skipHeroStats) {
    history_syncHeroStats()
  }
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
      
      // ВАЖНО: Сбрасываем фон для всех ячеек новой колонки, чтобы они не наследовали форматирование
      const lastRow = sheet.getLastRow()
      if (lastRow > HEADER_ROW) {
        sheet.getRange(DATA_START_ROW, newNightCol, lastRow - HEADER_ROW, 1).setBackground(null)
      }
      
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
  
  // ВАЖНО: Сбрасываем фон для всех ячеек новой колонки, чтобы они не наследовали форматирование от предыдущей колонки
  // Это предотвращает наследование окраски min/max от предыдущих колонок
  const lastRow = sheet.getLastRow()
  if (lastRow > HEADER_ROW) {
    sheet.getRange(DATA_START_ROW, newCol, lastRow - HEADER_ROW, 1).setBackground(null)
  }
  
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
      if (res && res.ok && res.price !== undefined) {
        // ВАЛИДАЦИЯ: Проверяем цену перед записью
        const validation = validatePrice_(res.price, name)
        if (validation.valid) {
          periodVals[i][0] = validation.price
          statusVals[i][0] = STATUS.OK
          updatedCount++
        } else {
          // Цена не прошла валидацию - не записываем и отмечаем как ошибку
          console.error(`History: цена не прошла валидацию для "${name}": ${validation.error}, цена: ${res.price}`)
          statusVals[i][0] = STATUS.WARNING
          errorCount++
        }
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
    if (res.ok && res.price !== undefined) {
      // ВАЛИДАЦИЯ: Проверяем цену перед записью
      const validation = validatePrice_(res.price, name)
      if (validation.valid) {
        sheet.getRange(row, col).setValue(validation.price)
        sheet.getRange(`${HISTORY_CONFIG.COLUMNS.STATUS}${row}`).setValue('✓')
        return 'updated'
      } else {
        console.error(`History: цена не прошла валидацию для "${name}": ${validation.error}, цена: ${res.price}`)
        sheet.getRange(`${HISTORY_CONFIG.COLUMNS.STATUS}${row}`).setValue('⚠️')
        return 'error'
      }
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
        Utilities.sleep(LIMITS.HISTORY_UPDATE_DELAY_MS)
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
    
    // Ищем последнюю заполненную цену из всех колонок (включая старые)
    let lastFoundPrice = null
    if (priceDataWidth > 0 && priceData[i]) {
      for (let j = priceData[i].length - 1; j >= 0; j--) {
        const value = priceData[i][j]
        if (typeof value === 'number' && !isNaN(value) && value > 0) {
          lastFoundPrice = value
          break
        }
      }
    }
    
    // Проверяем, есть ли цена за текущий период
    let currentPeriodPrice = null
    if (currentPeriodColIndex >= 0 && priceDataWidth > 0 && priceData[i] && priceData[i][currentPeriodColIndex]) {
      const price = priceData[i][currentPeriodColIndex]
      if (typeof price === 'number' && !isNaN(price) && price > 0) {
        currentPeriodPrice = price
      }
    }
    
    // Если есть цена за текущий период - используем её
    // Если нет, но есть последняя заполненная цена - используем её (будет окрашена в желтый)
    // Если нет цен в колонках, но текущая цена уже заполнена - сохраняем её (не перезаписываем на null)
    // Только если вообще нет цен и текущая цена пустая - ставим null
    const existingCurrentPrice = currentPrices[i][0]
    const hasExistingCurrentPrice = existingCurrentPrice !== null && 
                                    existingCurrentPrice !== '' && 
                                    Number.isFinite(Number(existingCurrentPrice)) && 
                                    Number(existingCurrentPrice) > 0
    
    if (currentPeriodPrice || lastFoundPrice) {
      // Есть цена в колонках - используем её
      currentPrices[i][0] = currentPeriodPrice || lastFoundPrice
    } else if (hasExistingCurrentPrice) {
      // Нет цен в колонках, но текущая цена уже заполнена - сохраняем её (не перезаписываем на null)
      currentPrices[i][0] = existingCurrentPrice
    } else {
      // Нет цен вообще - ставим null (не используем среднее из Min/Max, это некорректные данные)
      currentPrices[i][0] = null
    }
    
    // ЛОГИКА Min/Max:
    // Min/Max получаются из SteamWebAPI один раз при первоначальной настройке
    // При обновлении цен проверяем: если новая цена выходит за границы Min/Max - обновляем
    const currentMin = minPrices[i][0]
    const currentMax = maxPrices[i][0]
    const hasCurrentMin = currentMin !== null && currentMin !== '' && Number.isFinite(Number(currentMin)) && Number(currentMin) > 0
    const hasCurrentMax = currentMax !== null && currentMax !== '' && Number.isFinite(Number(currentMax)) && Number(currentMax) > 0
    
    // Если есть новая цена за текущий период - проверяем, нужно ли обновить Min/Max
    if (currentPeriodPrice && Number.isFinite(currentPeriodPrice) && currentPeriodPrice > 0) {
      let newMin = hasCurrentMin ? Number(currentMin) : currentPeriodPrice
      let newMax = hasCurrentMax ? Number(currentMax) : currentPeriodPrice
      
      // Если новая цена меньше текущего Min - обновляем Min
      if (currentPeriodPrice < newMin) {
        newMin = currentPeriodPrice
      }
      
      // Если новая цена больше текущего Max - обновляем Max
      if (currentPeriodPrice > newMax) {
        newMax = currentPeriodPrice
      }
      
      // Сохраняем обновленные значения
      minPrices[i][0] = newMin
      maxPrices[i][0] = newMax
      
    } else {
      // Если нет новой цены за текущий период - сохраняем существующие значения Min/Max
      // Это важно: не перезаписываем Min/Max из SteamWebAPI
      // Fallback: если Min/Max не установлены, но есть цены в колонках - вычисляем из них
      if (!hasCurrentMin || !hasCurrentMax) {
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
          // Устанавливаем Min/Max из колонок только если они еще не были установлены из SteamWebAPI
          if (!hasCurrentMin) {
            minPrices[i][0] = Math.min(...prices)
          } else {
            minPrices[i][0] = currentMin
          }
          
          if (!hasCurrentMax) {
            maxPrices[i][0] = Math.max(...prices)
          } else {
            maxPrices[i][0] = currentMax
          }
        } else {
          // Нет цен - оставляем как есть (или null)
          minPrices[i][0] = hasCurrentMin ? currentMin : null
          maxPrices[i][0] = hasCurrentMax ? currentMax : null
        }
      } else {
        // Min/Max уже установлены - сохраняем их без изменений
        minPrices[i][0] = currentMin
        maxPrices[i][0] = currentMax
      }
    }
    
    updatedCount++
  }
  
  // Batch запись всех значений
  sheet.getRange(DATA_START_ROW, currentPriceCol, count, 1).setValues(currentPrices)
  sheet.getRange(DATA_START_ROW, minPriceCol, count, 1).setValues(minPrices)
  sheet.getRange(DATA_START_ROW, maxPriceCol, count, 1).setValues(maxPrices)
  
  // Окрашиваем устаревшие цены в желтый (STABLE)
  // Если цена не за текущий период - она устарела и должна быть желтой
  const backgroundsToSet = []
  for (let i = 0; i < count; i++) {
    const currentPrice = currentPrices[i][0]
    if (currentPrice != null && currentPrice !== '') {
      // Проверяем, есть ли цена за текущий период
      const hasCurrentPeriodPrice = currentPeriodColIndex >= 0 && 
                                    priceDataWidth > 0 && 
                                    priceData[i] && 
                                    priceData[i][currentPeriodColIndex] &&
                                    typeof priceData[i][currentPeriodColIndex] === 'number' &&
                                    !isNaN(priceData[i][currentPeriodColIndex]) &&
                                    priceData[i][currentPeriodColIndex] > 0
      
      // Если нет цены за текущий период - цена устарела, окрашиваем в желтый
      if (!hasCurrentPeriodPrice) {
        backgroundsToSet.push({ row: i + DATA_START_ROW, col: currentPriceCol, color: COLORS.STABLE })
      } else {
        // Если есть цена за текущий период - сбрасываем фон (белый)
        backgroundsToSet.push({ row: i + DATA_START_ROW, col: currentPriceCol, color: null })
      }
    }
  }
  
  // Batch-применение фонов
  backgroundsToSet.forEach(bg => {
    sheet.getRange(bg.row, bg.col).setBackground(bg.color)
  })
  
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
  // Сортируем по колонке, чтобы последняя колонка для каждой даты была последней
  priceEntries.sort((a, b) => a.col - b.col)
  
  // Теперь для каждой даты последняя запись в массиве = последняя цена за день
  // Map автоматически перезаписывает предыдущее значение, поэтому последняя цена за день будет сохранена
  for (const entry of priceEntries) {
    pricesByDate.set(entry.dateKey, entry.value)
  }

  // Создаем уникальный список дат из Map (уже без дубликатов)
  const uniqueDateKeys = Array.from(pricesByDate.keys())

  // Преобразуем Map в массивы цен и дат (в хронологическом порядке)
  // ВАЖНО: Сортируем даты в хронологическом порядке перед созданием массивов
  const sortedDateKeys = uniqueDateKeys.sort((a, b) => {
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
      // Используем полдень (12:00) для избежания проблем с часовыми поясами
      const dateParts = dateKey.split('.')
      if (dateParts.length === 3) {
        const day = parseInt(dateParts[0], 10)
        const month = parseInt(dateParts[1], 10) - 1 // месяцы в JS начинаются с 0
        const year = 2000 + parseInt(dateParts[2], 10) // предполагаем формат yy -> 20yy
        dates.push(new Date(year, month, day, 12, 0, 0)) // 12:00:00 для точности
      } else {
        dates.push(new Date()) // fallback
      }
    }
  }
  
  // Проверка: количество цен и дат должно совпадать (одна запись на день)
  if (prices.length !== dates.length) {
    console.warn(`History: несоответствие количества цен (${prices.length}) и дат (${dates.length}) для строки ${row}`)
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
// ВАЖНО: Использует разницу между реальными датами, а не количество записей
function history_calculateDaysChange_(prices, dates, currentTrend) {
  if (prices.length < 3) return 0
  if (dates.length < 2) return 0
  
  // ВАЖНО: dates уже должны быть сгруппированы по дате (одна запись на день)
  // Проверяем, что dates.length соответствует количеству уникальных дат
  // Если нет - значит группировка не работает и нужно исправить вызывающий код
  
  // Анализируем тренды для каждого периода, начиная с предпоследнего
  for (let i = prices.length - 2; i >= 1; i--) {
    const periodPrices = prices.slice(0, i + 1)
    const periodTrend = history_simpleComparison_(periodPrices)
    
    if (periodTrend !== currentTrend) {
      // Нашли смену тренда
      const changeDate = dates[i]
      const currentDate = dates[dates.length - 1]
      
      if (changeDate instanceof Date && currentDate instanceof Date) {
        // Вычисляем разницу в днях между датами
        // Используем Math.floor для точного подсчета календарных дней
        const diffTime = Math.abs(currentDate - changeDate)
        const daysDiff = Math.floor(diffTime / (1000 * 60 * 60 * 24))
        return daysDiff > 0 ? daysDiff : 1 // Минимум 1 день
      }
    }
  }
  
  // Если тренд не менялся, вычисляем разницу между первой и последней датой
  if (dates.length >= 2) {
    const firstDate = dates[0]
    const lastDate = dates[dates.length - 1]
    if (firstDate instanceof Date && lastDate instanceof Date) {
      // Вычисляем разницу в днях между датами
      // Используем Math.floor для точного подсчета календарных дней
      const diffTime = Math.abs(lastDate - firstDate)
      const daysDiff = Math.floor(diffTime / (1000 * 60 * 60 * 24))
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
  
  // Определяем номера колонок для расширенной аналитики (K, L, J)
  const phaseCol = getColumnIndex(HISTORY_COLUMNS.PHASE)           // K
  const potentialCol = getColumnIndex(HISTORY_COLUMNS.POTENTIAL)   // L
  const recommendationCol = getColumnIndex(HISTORY_COLUMNS.RECOMMENDATION) // J
  
  // ОПТИМИЗАЦИЯ: Читаем все данные одним batch-запросом
  const names = sheet.getRange(DATA_START_ROW, 2, count, 1).getValues() // B
  const trends = sheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.TREND), count, 1).getValues() // M (теперь содержит объединенный тренд+дни)
  const phases = sheet.getRange(DATA_START_ROW, phaseCol, count, 1).getValues()
  const potentials = sheet.getRange(DATA_START_ROW, potentialCol, count, 1).getValues()
  const recommendations = sheet.getRange(DATA_START_ROW, recommendationCol, count, 1).getValues()
  const investmentScores = sheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.INVESTMENT_SCORE), count, 1).getValues() // I
  
  // ОПТИМИЗАЦИЯ: Читаем все цены и заголовки колонок одним batch-запросом
  const priceDataWidth = lastCol >= firstDateCol ? lastCol - firstDateCol + 1 : 0
  const allPriceData = priceDataWidth > 0 
    ? sheet.getRange(DATA_START_ROW, firstDateCol, count, priceDataWidth).getValues()
    : []
  const allHeaders = priceDataWidth > 0
    ? sheet.getRange(HEADER_ROW, firstDateCol, 1, priceDataWidth).getDisplayValues()[0]
    : []

  let updatedCount = 0
  let skippedCount = 0
  let errorCount = 0
  const startedAt = Date.now()
  const MAX_EXECUTION_TIME_MS = 300000 // 5 минут
  const MAX_ROW_PROCESSING_TIME_MS = 5000 // Максимум 5 секунд на одну строку
  
  // Анализируем тренды для всех строк
  for (let i = 0; i < count; i++) {
    // Проверка общего таймаута
    if (Date.now() - startedAt > MAX_EXECUTION_TIME_MS) {
      console.warn(`History: превышено время выполнения updateTrends (${MAX_EXECUTION_TIME_MS}ms), прервано на строке ${i + 1}`)
      break
    }
    
    const rowStartTime = Date.now()
    const name = String(names[i][0] || '').trim()
    if (!name) {
      skippedCount++
      continue
    }
    
    const row = i + 2
    
    // Обрабатываем каждую строку в try-catch для защиты от зависания на проблемных данных
    try {
    
    // ОПТИМИЗАЦИЯ: Используем уже прочитанные данные вместо отдельных запросов
    const pricesByDate = new Map()
    const dateHeaders = []
    const priceEntries = []
    
    // Обрабатываем данные из уже прочитанного массива
    if (priceDataWidth > 0 && allPriceData[i]) {
      for (let j = 0; j < priceDataWidth; j++) {
        const value = allPriceData[i][j]
        const headerDisplay = allHeaders[j]
        // ВАЛИДАЦИЯ: Проверяем цену перед добавлением
        if (typeof value === 'number' && !isNaN(value) && value > 0 && headerDisplay) {
          const priceValidation = validatePrice_(value, `${name} (колонка ${firstDateCol + j})`)
          if (!priceValidation.valid) {
            console.warn(`History: пропущена некорректная цена для "${name}" в колонке ${firstDateCol + j}: ${value}`)
            continue // Пропускаем эту цену
          }
          
          const headerStr = String(headerDisplay).trim()
          const dateMatch = headerStr.match(/^(\d{2}\.\d{2}\.\d{2})/)
          if (dateMatch) {
            const dateKey = dateMatch[1]
            const col = firstDateCol + j
            priceEntries.push({
              dateKey,
              value: priceValidation.price, // Используем валидированную цену
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
    // Сортируем по колонке, чтобы последняя колонка для каждой даты была последней
    priceEntries.sort((a, b) => a.col - b.col)
    for (const entry of priceEntries) {
      // Map автоматически перезаписывает предыдущее значение, поэтому последняя цена за день будет сохранена
      pricesByDate.set(entry.dateKey, entry.value)
    }
    
    // Создаем уникальный список дат из Map (уже без дубликатов)
    const uniqueDateKeys = Array.from(pricesByDate.keys())
    
    // Сортируем даты в хронологическом порядке
    const sortedDateKeys = uniqueDateKeys.sort((a, b) => {
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
    // ВАЖНО: prices и dates должны иметь одинаковую длину - одна запись на день
    const prices = []
    const dates = []
    for (const dateKey of sortedDateKeys) {
      const price = pricesByDate.get(dateKey)
      if (price) {
        // Дополнительная валидация перед добавлением (на случай, если валидация была пропущена ранее)
        const priceValidation = validatePrice_(price, `${name} (${dateKey})`)
        if (!priceValidation.valid) {
          console.warn(`History: пропущена некорректная цена для "${name}" за ${dateKey}: ${price}`)
          continue
        }
        
        prices.push(priceValidation.price)
        // Преобразуем строку даты в объект Date для расчета разницы
        // Используем полдень (12:00) для избежания проблем с часовыми поясами
        const dateParts = dateKey.split('.')
        if (dateParts.length === 3) {
          const day = parseInt(dateParts[0], 10)
          const month = parseInt(dateParts[1], 10) - 1 // месяцы в JS начинаются с 0
          const year = 2000 + parseInt(dateParts[2], 10) // предполагаем формат yy -> 20yy
          dates.push(new Date(year, month, day, 12, 0, 0)) // 12:00:00 для точности
        } else {
          dates.push(new Date()) // fallback
        }
      }
    }
    
    // Проверка: количество цен и дат должно совпадать
    if (prices.length !== dates.length) {
      console.warn(`History: несоответствие количества цен (${prices.length}) и дат (${dates.length}) для строки ${row}`)
    }
    
    // ВАЛИДАЦИЯ: Проверяем достаточность данных для анализа
    if (prices.length < 2) {
      console.warn(`History: недостаточно данных для анализа тренда для "${name}" (${prices.length} цен)`)
      trends[i][0] = '❓'
      phases[i][0] = '❓'
      potentials[i][0] = null
      recommendations[i][0] = '👀 НАБЛЮДАТЬ'
      skippedCount++
      continue
    }
    
    // Проверка таймаута обработки строки
    if (Date.now() - rowStartTime > MAX_ROW_PROCESSING_TIME_MS) {
      console.warn(`History: превышено время обработки строки ${row} ("${name}") - ${MAX_ROW_PROCESSING_TIME_MS}ms, пропускаем`)
      trends[i][0] = '❓'
      phases[i][0] = '❓'
      potentials[i][0] = null
      recommendations[i][0] = '❓ ТАЙМАУТ'
      skippedCount++
      continue
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
      
      // Получаем Investment Score из колонки I (если доступен)
      const investmentScoreStr = String(investmentScores[i][0] || '').trim()
      let investmentScore = null
      if (investmentScoreStr && investmentScoreStr !== '—') {
        // Парсим число из формата "🟩 0.93"
        const scoreMatch = investmentScoreStr.match(/(\d+\.?\d*)/)
        if (scoreMatch) {
          investmentScore = parseFloat(scoreMatch[1])
        }
      }
      
      recommendations[i][0] = history_generateRecommendation_(
        phases[i][0],
        cleanTrend,
        potential,
        daysChange,
        investmentScore
      )
    } else {
      phases[i][0] = '❓'
      potentials[i][0] = null
      recommendations[i][0] = '👀 НАБЛЮДАТЬ'
    }
    
    updatedCount++
    } catch (e) {
      // Обработка ошибок для строк с пропущенными значениями (SteamWebAPI не обрабатывает некоторые новые предметы)
      console.error(`History: ошибка при обработке строки ${row} (${name}):`, e.message)
      errorCount++
      // Устанавливаем значения по умолчанию для проблемных строк
      trends[i][0] = '❓'
      phases[i][0] = '❓'
      potentials[i][0] = null
      recommendations[i][0] = '❓ ОШИБКА'
      // Продолжаем обработку следующих строк
    }
  }
  
  // Batch обновление всех данных за одну операцию
  const trendCol = getColumnIndex(HISTORY_COLUMNS.TREND)
  sheet.getRange(DATA_START_ROW, trendCol, count, 1).setValues(trends)
  sheet.getRange(DATA_START_ROW, phaseCol, count, 1).setValues(phases)
  sheet.getRange(DATA_START_ROW, potentialCol, count, 1).setValues(potentials)
  sheet.getRange(DATA_START_ROW, recommendationCol, count, 1).setValues(recommendations)
  
  // Применяем все правила условного форматирования (тренды + аналитика)
  history_applyAllConditionalFormatting_(sheet)
  
  console.log(`History: обновлено трендов: ${updatedCount}, пропущено: ${skippedCount}, ошибок: ${errorCount}`)
  
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

/**
 * Нормализует формат цен во всех колонках с датами
 * Конвертирует текстовые значения вида "39,99 ₽" в числа и применяет единый формат
 */
function history_normalizePriceFormats() {
  const sheet = getOrCreateHistorySheet_()
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  const firstDateCol = HISTORY_COLUMNS.FIRST_DATE_COL
  
  if (lastRow <= 1 || lastCol < firstDateCol) {
    SpreadsheetApp.getUi().alert('Нет данных для нормализации')
    return
  }
  
  const dateColsCount = lastCol - firstDateCol + 1
  const dataRowsCount = lastRow - 1
  
  let convertedCount = 0
  let errorCount = 0
  
  // Обрабатываем каждую колонку с датами
  for (let col = firstDateCol; col <= lastCol; col++) {
    const values = sheet.getRange(DATA_START_ROW, col, dataRowsCount, 1).getValues()
    const displayValues = sheet.getRange(DATA_START_ROW, col, dataRowsCount, 1).getDisplayValues()
    const newValues = []
    let hasChanges = false
    
    for (let i = 0; i < values.length; i++) {
      const value = values[i][0]
      const displayValue = String(displayValues[i][0] || '').trim()
      
      // Если значение пустое - оставляем пустым
      if (!value && !displayValue) {
        newValues.push([''])
        continue
      }
      
      // Проверяем, нужно ли конвертировать значение
      // Конвертируем если: displayValue содержит запятую с цифрами после (десятичный разделитель) или знак рубля
      // ИЛИ value - строка, которую можно распарсить
      const hasDecimalComma = displayValue.match(/,\d{1,2}(\s*₽)?$/) // Запятая с 1-2 цифрами после (десятичный разделитель)
      const hasRuble = displayValue.includes('₽')
      const isStringValue = typeof value === 'string' && value.trim().length > 0
      
      // Если значение уже число и displayValue не содержит признаков текстового формата - оставляем как есть
      if (typeof value === 'number' && !isNaN(value) && value > 0 && !hasDecimalComma && !hasRuble) {
        newValues.push([value])
        continue
      }
      
      // Если значение требует конвертации (текстовый формат "39,99 ₽" или строка) - конвертируем
      if (hasDecimalComma || hasRuble || isStringValue) {
        try {
          // Используем displayValue для парсинга, так как он содержит реальное отображение
          let cleanValue = displayValue
          
          // Убираем знак рубля и пробелы
          cleanValue = cleanValue.replace(/₽/g, '').replace(/\s+/g, '').trim()
          
          // Заменяем запятую на точку для десятичных (российский формат "39,99" -> "39.99")
          cleanValue = cleanValue.replace(',', '.')
          
          // Убираем все нечисловые символы кроме точки и минуса
          cleanValue = cleanValue.replace(/[^\d.-]/g, '')
          
          const numValue = parseFloat(cleanValue)
          
          // ВАЛИДАЦИЯ: Проверяем конвертированное значение
          if (!isNaN(numValue) && isFinite(numValue)) {
            const validation = validatePrice_(numValue, `колонка ${col}, строка ${i + DATA_START_ROW}`)
            if (validation.valid) {
              // Убеждаемся, что записываем именно число, а не строку
              newValues.push([Number(validation.price)])
              hasChanges = true
              convertedCount++
            } else {
              // Цена не прошла валидацию - очищаем ячейку или оставляем пустой
              console.warn(`History: цена не прошла валидацию при нормализации: ${validation.error}, значение: ${numValue}`)
              newValues.push([''])
              hasChanges = true
              errorCount++
            }
          } else {
            // Не удалось конвертировать - оставляем как есть
            newValues.push([value])
          }
        } catch (e) {
          console.error(`History: ошибка конвертации значения в колонке ${col}, строка ${i + DATA_START_ROW}:`, e)
          newValues.push([value])
          errorCount++
        }
      } else {
        // Значение не требует конвертации - оставляем как есть
        newValues.push([value])
      }
    }
    
    // Записываем конвертированные значения, если были изменения
    if (hasChanges) {
      const range = sheet.getRange(DATA_START_ROW, col, dataRowsCount, 1)
      // Сначала применяем формат, чтобы Google Sheets правильно интерпретировал числа
      range.setNumberFormat('#,##0.00 ₽')
      // Затем записываем значения как числа
      range.setValues(newValues)
      // Применяем выравнивание
      range.setHorizontalAlignment('center')
      range.setVerticalAlignment('middle')
    } else {
      // Даже если изменений не было, убеждаемся, что формат применен
      history_formatPriceColumn_(sheet, col)
    }
  }
  
  // Применяем форматирование ко всем колонкам дат одним batch-запросом для единообразия
  // Это гарантирует, что все ячейки имеют правильный формат, даже если они не были изменены
  history_formatAllDateColumns_(sheet)
  
  // Проверяем результат: берем несколько примеров для демонстрации
  let examples = []
  let exampleCount = 0
  const maxExamples = 3
  
  if (convertedCount > 0) {
    // Берем первую колонку с датами и проверяем первые несколько строк
    const firstDateCol = HISTORY_COLUMNS.FIRST_DATE_COL
    const checkRows = Math.min(20, dataRowsCount)
    const sampleValues = sheet.getRange(DATA_START_ROW, firstDateCol, checkRows, 1).getValues()
    const sampleDisplay = sheet.getRange(DATA_START_ROW, firstDateCol, checkRows, 1).getDisplayValues()
    
    for (let i = 0; i < checkRows && exampleCount < maxExamples; i++) {
      const val = sampleValues[i][0]
      const display = String(sampleDisplay[i][0] || '').trim()
      if (val && typeof val === 'number' && val > 0 && display) {
        examples.push(`"${display}" → число ${val.toFixed(2)}`)
        exampleCount++
      }
    }
  }
  
  let message = `Нормализация завершена:\n` +
    `• Конвертировано значений: ${convertedCount}\n` +
    `• Ошибок: ${errorCount}\n` +
    `• Колонок обработано: ${dateColsCount}`
  
  if (examples.length > 0) {
    message += `\n\nПримеры конвертации:\n${examples.join('\n')}`
    message += `\n\n✅ ВАЖНО: Визуально значения выглядят так же, но теперь это ЧИСЛА, а не текст!`
    message += `\n\nДо: текст "39,99 ₽" (нельзя использовать в формулах)`
    message += `\nПосле: число 39.99 с форматом валюты (можно использовать в формулах)`
    message += `\n\nТеперь все цены в том же формате, что и от SteamWebAPI (числа).`
  } else if (convertedCount > 0) {
    message += `\n\n✅ Значения конвертированы из текста в числа.`
    message += `\nТеперь они в том же формате, что и цены от SteamWebAPI (числа, а не текст).`
    message += `\nВизуально они выглядят так же, но теперь их можно использовать в формулах.`
  } else {
    message += `\n\nℹ️ Все значения уже были в правильном формате (числа).`
  }
  
  console.log(`History: ${message}`)
  SpreadsheetApp.getUi().alert('Нормализация формата цен', message, SpreadsheetApp.getUi().ButtonSet.OK)
}

// Убедиться что колонки для расширенной аналитики существуют (K, L, J)
function history_ensureExtendedAnalyticsColumns_() {
  const sheet = getOrCreateHistorySheet_()
  const lastRow = sheet.getLastRow()
  
  // Колонки K, L, J зарезервированы для Фаза, Потенциал, Рекомендация
  // Проверяем заголовки
  const phaseCol = getColumnIndex(HISTORY_COLUMNS.PHASE)           // K
  const potentialCol = getColumnIndex(HISTORY_COLUMNS.POTENTIAL)   // L
  const recommendationCol = getColumnIndex(HISTORY_COLUMNS.RECOMMENDATION) // J
  
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
function history_generateRecommendation_(phase, trend, potential, daysChange, investmentScore = null, heroTrend = null) {
  // Если есть Investment Score, используем его (0-100 шкала)
  if (investmentScore !== null && typeof investmentScore === 'number' && !isNaN(investmentScore)) {
    if (investmentScore >= 75) {
      return '🟩 КУПИТЬ'
    }
    if (investmentScore >= 60) {
      return '🟨 ДЕРЖАТЬ'
    }
    if (investmentScore < 40) {
      return '🟥 ПРОДАТЬ'
    }
    return '👀 НАБЛЮДАТЬ'
  }
  
  // Fallback на старую логику, если Investment Score не рассчитан
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

// ===== СИСТЕМА ИНВЕСТИЦИОННЫХ РЕКОМЕНДАЦИЙ =====

/**
 * Синхронизация статистики героев из HeroStats в History
 * Обновляет колонки O-T (Hero Trend, Contest Rate Change, Contest Rate, Pick Rate, Win Rate, Hero Name)
 * Оптимизировано: кэширование статистики и Hero Trend Score по heroId
 */
function history_syncHeroStats() {
  const sheet = getOrCreateHistorySheet_()
  const lastRow = sheet.getLastRow()
  if (lastRow < DATA_START_ROW) return
  
  const startTime = Date.now()
  const TIME_BUDGET_MS = 300000 // 5 минут (оставляем запас до лимита 6 минут)
  
  const mappings = heroMapping_getAllMappings()
  const itemNames = sheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), lastRow - HEADER_ROW, 1).getValues()
  
  // ОПТИМИЗАЦИЯ: Собираем уникальные heroId и кэшируем статистику для каждого
  const uniqueHeroIds = new Set()
  for (const itemName of Object.keys(mappings)) {
    const mapping = mappings[itemName]
    if (mapping.heroId && mapping.category === 'Hero Item') {
      uniqueHeroIds.add(mapping.heroId)
    }
  }
  
  // ОПТИМИЗАЦИЯ: Кэшируем статистику для каждого уникального heroId (вызываем getLatestStats только один раз)
  const heroDataMap = {}
  const heroTrendScoreCache = {} // Кэш для Hero Trend Score по heroId
  
  for (const heroId of uniqueHeroIds) {
    // Проверка времени выполнения
    if (Date.now() - startTime > TIME_BUDGET_MS) {
      console.warn(`History: history_syncHeroStats прервано по таймауту (обработано ${Object.keys(heroDataMap).length} из ${uniqueHeroIds.size} героев)`)
      break
    }
    
    // Находим mapping для получения heroName (берем первый попавшийся предмет с этим heroId)
    let heroName = null
    let rankCategory = null
    for (const itemName of Object.keys(mappings)) {
      const mapping = mappings[itemName]
      if (mapping.heroId === heroId && mapping.category === 'Hero Item') {
        heroName = mapping.heroName
        break
      }
    }
    
    // Приоритет: High Rank > All Ranks (вызываем только один раз для каждого)
    const highRankStats = heroStats_getLatestStats(heroId, 'High Rank')
    const allRanksStats = heroStats_getLatestStats(heroId, 'All Ranks')
    const stats = highRankStats || allRanksStats
    rankCategory = highRankStats ? 'High Rank' : 'All Ranks'
    
    if (stats) {
      heroDataMap[heroId] = {
        stats: stats,
        rankCategory: rankCategory,
        heroName: heroName
      }
      
      // ОПТИМИЗАЦИЯ: Кэшируем Hero Trend Score для этого heroId (вычисляем один раз)
      try {
        const heroStatsObj = {[rankCategory]: stats}
        const heroTrendScore = analytics_calculateHeroTrendScore(heroId, rankCategory, heroStatsObj)
        heroTrendScoreCache[heroId] = analytics_formatScore(heroTrendScore)
      } catch (e) {
        console.error(`History: ошибка расчета Hero Trend Score для heroId ${heroId}:`, e)
        heroTrendScoreCache[heroId] = ''
      }
    }
  }
  
  // ОПТИМИЗАЦИЯ: Подготавливаем данные для batch записи
  const heroNames = []
  const heroTrends = []
  const proContestRates = []
  const proContestRateChanges = []
  const pickRateChanges7d = []
  const pickRateChanges24h = []
  const pickRates = []
  const winRates = []
  const metaSignals = []
  const updateRows = []
  
  const heroNameCol = getColumnIndex(HISTORY_COLUMNS.HERO_NAME)
  const heroTrendCol = getColumnIndex(HISTORY_COLUMNS.HERO_TREND)
  const proContestRateCol = getColumnIndex(HISTORY_COLUMNS.PRO_CONTEST_RATE_CURRENT)
  const proContestRateChangeCol = getColumnIndex(HISTORY_COLUMNS.PRO_CONTEST_RATE_CHANGE_7D)
  const pickRateChange7dCol = getColumnIndex(HISTORY_COLUMNS.PICK_RATE_CHANGE_IMMORTAL_7D)
  const pickRateChange24hCol = getColumnIndex(HISTORY_COLUMNS.PICK_RATE_CHANGE_IMMORTAL_24H)
  const pickRateCol = getColumnIndex(HISTORY_COLUMNS.PICK_RATE_IMMORTAL)
  const winRateCol = getColumnIndex(HISTORY_COLUMNS.WIN_RATE_CURRENT)
  const metaSignalCol = getColumnIndex(HISTORY_COLUMNS.META_SIGNAL)
  
  // Подготавливаем все значения
  for (let i = 0; i < itemNames.length; i++) {
    const itemName = String(itemNames[i][0] || '').trim()
    if (!itemName) {
      heroNames.push([''])
      heroTrends.push([''])
      proContestRates.push([''])
      proContestRateChanges.push([''])
      pickRateChanges7d.push([''])
      pickRateChanges24h.push([''])
      pickRates.push([''])
      winRates.push([''])
      metaSignals.push([''])
      continue
    }
    
    const row = DATA_START_ROW + i
    updateRows.push(row)
    const mapping = mappings[itemName]
    
    if (mapping && mapping.heroId && mapping.category === 'Hero Item') {
      const heroData = heroDataMap[mapping.heroId]
      if (heroData && heroData.stats) {
        try {
          const stats = typeof heroData.stats === 'string' ? JSON.parse(heroData.stats) : heroData.stats
          
          // ОПТИМИЗАЦИЯ: Используем кэшированный Hero Trend Score вместо пересчета
          const heroTrendScore = heroTrendScoreCache[mapping.heroId] || ''
          
          heroNames.push([heroData.heroName || ''])
          heroTrends.push([heroTrendScore])
          
          // Pro Contest Rate (текущий) - в процентах (45.2 = 45.2%), формат процента умножает на 100, поэтому делим на 100
          proContestRates.push([stats.proContestRate ? stats.proContestRate / 100 : 0])
          
          // Pro Contest Rate Change (7d) - в процентах (15 = 15%), формат процента умножает на 100, поэтому делим на 100
          proContestRateChanges.push([stats.proContestRateChange7d ? stats.proContestRateChange7d / 100 : 0])
          
          // Pick Rate Change Immortal (7d) - в процентах (10 = 10%), формат процента умножает на 100, поэтому делим на 100
          pickRateChanges7d.push([stats.pickRateChange7d ? stats.pickRateChange7d / 100 : 0])
          
          // Pick Rate Change Immortal (24h) - в процентах (25 = 25%), формат процента умножает на 100, поэтому делим на 100
          pickRateChanges24h.push([stats.pickRateChange24h ? stats.pickRateChange24h / 100 : 0])
          
          // Pick Rate Immortal - в процентах (1.4 = 1.4%), формат процента умножает на 100, поэтому делим на 100
          pickRates.push([stats.pickRatePercent !== undefined ? stats.pickRatePercent / 100 : 0])
          
          // Win Rate - в процентах (52.02 = 52.02%), формат процента умножает на 100, поэтому делим на 100
          winRates.push([stats.winRate ? stats.winRate / 100 : 0])
          
          // Мета сигнал - рассчитываем отдельно
          let metaSignal = ''
          try {
            // Используем те же данные, что и для Hero Trend Score
            const rankCategoryForMeta = heroData.rankCategory || (mapping && mapping.heroId ? 'High Rank' : null)
            if (rankCategoryForMeta && heroData.stats) {
              const heroStatsObjForMeta = {[rankCategoryForMeta]: heroData.stats}
              const metaSignalScore = analytics_calculateMetaSignal(mapping.heroId, rankCategoryForMeta, heroStatsObjForMeta)
              metaSignal = analytics_formatMetaSignal(metaSignalScore)
            }
          } catch (e) {
            console.error(`History: ошибка расчета Мета сигнала для heroId ${mapping.heroId}:`, e)
          }
          metaSignals.push([metaSignal])
        } catch (e) {
          console.log(`Ошибка при подготовке статистики для ${itemName}: ${e.message}`)
          heroNames.push([''])
          heroTrends.push([''])
          proContestRates.push([''])
          proContestRateChanges.push([''])
          pickRateChanges7d.push([''])
          pickRateChanges24h.push([''])
          pickRates.push([''])
          winRates.push([''])
          metaSignals.push([''])
        }
      } else {
        // Нет данных о герое
        heroNames.push([''])
        heroTrends.push([''])
        proContestRates.push([''])
        proContestRateChanges.push([''])
        pickRateChanges7d.push([''])
        pickRateChanges24h.push([''])
        pickRates.push([''])
        winRates.push([''])
        metaSignals.push([''])
      }
    } else {
      // Общий предмет - очищаем колонки
      heroNames.push([''])
      heroTrends.push([''])
      proContestRates.push([''])
      proContestRateChanges.push([''])
      pickRateChanges7d.push([''])
      pickRateChanges24h.push([''])
      pickRates.push([''])
      winRates.push([''])
      metaSignals.push([''])
    }
  }
  
  // BATCH ЗАПИСЬ: Записываем все колонки одним batch операциями
  if (updateRows.length > 0 || heroNames.length > 0) {
    const count = heroNames.length
    sheet.getRange(DATA_START_ROW, heroNameCol, count, 1).setValues(heroNames)
    sheet.getRange(DATA_START_ROW, heroTrendCol, count, 1).setValues(heroTrends)
    sheet.getRange(DATA_START_ROW, proContestRateCol, count, 1).setValues(proContestRates)
    sheet.getRange(DATA_START_ROW, proContestRateChangeCol, count, 1).setValues(proContestRateChanges)
    sheet.getRange(DATA_START_ROW, pickRateChange7dCol, count, 1).setValues(pickRateChanges7d)
    sheet.getRange(DATA_START_ROW, pickRateChange24hCol, count, 1).setValues(pickRateChanges24h)
    sheet.getRange(DATA_START_ROW, pickRateCol, count, 1).setValues(pickRates)
    sheet.getRange(DATA_START_ROW, winRateCol, count, 1).setValues(winRates)
    sheet.getRange(DATA_START_ROW, metaSignalCol, count, 1).setValues(metaSignals)
  }
  
  // Проверяем изменения Hero Trend Score (важные уведомления)
  try {
    telegram_checkHeroTrendChanges_()
  } catch (e) {
    console.error('History: ошибка при проверке изменений Hero Trend Score:', e)
    // Не прерываем выполнение, просто логируем ошибку
  }
  
  // Проверяем Мета сигнал (горячие уведомления о патч-имбах)
  try {
    telegram_checkMetaSignalOpportunities_()
  } catch (e) {
    console.error('History: ошибка при проверке Мета сигнала:', e)
    // Не прерываем выполнение, просто логируем ошибку
  }
}

/**
 * Обновление колонок статистики героев (M-R)
 * Алиас для history_syncHeroStats()
 */
function history_updateHeroStatsColumns() {
  history_syncHeroStats()
}

/**
 * Расчет Investment Score для всех предметов в History
 */
function history_updateInvestmentScores() {
  const sheet = getOrCreateHistorySheet_()
  const lastRow = sheet.getLastRow()
  if (lastRow < DATA_START_ROW) return
  
  const mappings = heroMapping_getAllMappings()
  const itemNames = sheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), lastRow - HEADER_ROW, 1).getValues()
  
  // Получаем данные из SteamWebAPI для всех предметов (batch)
  const itemNamesList = itemNames.map(row => String(row[0] || '').trim()).filter(name => name)
  const itemsData = {}
  
  // Batch запросы (до 50 предметов за раз)
  const batchSize = API_CONFIG.STEAM_WEB_API.MAX_ITEMS_PER_REQUEST
  for (let i = 0; i < itemNamesList.length; i += batchSize) {
    const batch = itemNamesList.slice(i, i + batchSize)
    const result = steamWebAPI_fetchItems(batch, 'dota2')
    if (result.ok && result.items) {
      result.items.forEach(item => {
        if (item.marketname) {
          itemsData[item.marketname] = item
        }
      })
    }
    // Задержка между batch запросами
    if (i + batchSize < itemNamesList.length) {
      Utilities.sleep(LIMITS.METRICS_UPDATE_DELAY_MS)
    }
  }
  
  // Подготовка данных для batch-операций
  const investmentScores = []
  
  // Рассчитываем Investment Score для всех строк
  for (let i = 0; i < itemNames.length; i++) {
    const itemName = String(itemNames[i][0] || '').trim()
    if (!itemName) {
      investmentScores.push([null])
      continue
    }
    
    const mapping = mappings[itemName]
    const itemData = itemsData[itemName]
    
    if (!itemData) {
      investmentScores.push([null])
      continue
    }
    
    const row = DATA_START_ROW + i
    
    // Получаем историю цен для предмета
    const historyData = history_getPriceHistoryForItem_(sheet, row)
    
    // Определяем категорию и heroId
    const category = mapping ? mapping.category : 'Common Item'
    const heroId = mapping && mapping.heroId ? mapping.heroId : null
    const rankCategory = mapping && mapping.heroId ? 'High Rank' : null // Приоритет High Rank
    
    // Получаем статистику героя
    let heroStats = null
    if (heroId && rankCategory) {
      const latestStats = heroStats_getLatestStats(heroId, rankCategory)
      if (latestStats) {
        heroStats = {[rankCategory]: latestStats}
      }
    }
    
    // Рассчитываем Investment Score
    const investmentScore = analytics_calculateInvestmentScore(
      itemData,
      heroStats,
      historyData,
      category,
      heroId,
      rankCategory
    )
    
    investmentScores.push([analytics_formatScore(investmentScore)])
  }
  
  // Batch-запись Investment Scores
  const count = investmentScores.length
  if (count > 0) {
    sheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.INVESTMENT_SCORE), count, 1).setValues(investmentScores)
  }
  
  // Проверяем возможности для покупки (критические уведомления)
  try {
    telegram_checkHistoryInvestmentOpportunities_()
  } catch (e) {
    console.error('History: ошибка при проверке возможностей для покупки:', e)
    // Не прерываем выполнение, просто логируем ошибку
  }
}

/**
 * Получение истории цен для предмета из History
 * @param {Sheet} sheet - Лист History
 * @param {number} row - Номер строки
 * @returns {Object} {prices: Array<number>, dates: Array<Date>}
 */
function history_getPriceHistoryForItem_(sheet, row) {
  const prices = []
  const dates = []
  
  const firstDateCol = HISTORY_COLUMNS.FIRST_DATE_COL
  const lastCol = sheet.getLastColumn()
  
  for (let col = firstDateCol; col <= lastCol; col++) {
    const header = sheet.getRange(HEADER_ROW, col).getValue()
    const price = sheet.getRange(row, col).getValue()
    
    if (price && typeof price === 'number' && price > 0) {
      prices.push(price)
      
      // Парсим дату из заголовка
      const dateMatch = String(header).match(/(\d{2})\.(\d{2})\.(\d{2})/)
      if (dateMatch) {
        const day = parseInt(dateMatch[1])
        const month = parseInt(dateMatch[2]) - 1
        const year = 2000 + parseInt(dateMatch[3])
        dates.push(new Date(year, month, day, 12, 0, 0)) // Полдень для точности
      }
    }
  }
  
  return { prices, dates }
}