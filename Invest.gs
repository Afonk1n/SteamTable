// Invest module
// Используем константы из Constants.gs
const INVEST_CONFIG = {
  STEAM_APPID: STEAM_APP_ID,
  STEAM_FEE: STEAM_FEE,
  COLUMNS: INVEST_COLUMNS,
}

const INVEST_SHEET_NAME = SHEET_NAMES.INVEST

// Форматирование новой строки Invest (при добавлении из History)
function invest_formatNewRow_(sheet, row) {
  if (row <= HEADER_ROW) return
  const name = sheet.getRange(`B${row}`).getValue()
  if (!name) return
  
  // Базовое форматирование строки (A-R = 18 колонок после удаления H, K и DAYS_CHANGE)
  const numCols = getColumnIndex(INVEST_COLUMNS.RECOMMENDATION)
  sheet.getRange(row, 1, 1, numCols).setVerticalAlignment('middle').setHorizontalAlignment('center')
  sheet.getRange(`B${row}`).setHorizontalAlignment('left')
  
  // Форматы чисел (используем константы)
  // УДАЛЕНЫ: H (Текущая стоимость без комиссии), K (Прибыль % без комиссии)
  sheet.getRange(`C${row}`).setNumberFormat(NUMBER_FORMATS.INTEGER) // Количество
  sheet.getRange(`D${row}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Цена покупки
  sheet.getRange(`E${row}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Текущая цена
  sheet.getRange(`F${row}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Цель
  sheet.getRange(`G${row}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Общие вложения
  sheet.getRange(`H${row}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Текущая стоимость с комиссией (было I)
  sheet.getRange(`I${row}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Профит (было J)
  sheet.getRange(`J${row}`).setNumberFormat(NUMBER_FORMATS.PERCENT) // Прибыль % с комиссией (было L)
  sheet.getRange(`L${row}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Min цена (было N)
  sheet.getRange(`M${row}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Max цена (было O)
  // O: Фаза (было R, убрали DAYS_CHANGE)
  // P: Потенциал (было S)
  // Q: Рекомендация (было T)
  
  // Устанавливаем высоту строки
  sheet.setRowHeight(row, ROW_HEIGHT)
}

// Функции getInvestSheet_ и getOrCreateInvestSheet_ перенесены в SheetService.gs

function invest_applySale(row, qtySold, sellPricePerUnit) {
  const sheet = getInvestSheet_()
  if (!sheet) return
  const name = sheet.getRange(`${INVEST_CONFIG.COLUMNS.NAME}${row}`).getValue()
  const qtyAvailable = Number(sheet.getRange(`${INVEST_CONFIG.COLUMNS.QUANTITY}${row}`).getValue())
  if (!name || !Number.isFinite(qtyAvailable) || qtyAvailable <= 0) return

  const remaining = qtyAvailable - qtySold
  if (remaining > 0) {
    sheet.getRange(`${INVEST_CONFIG.COLUMNS.QUANTITY}${row}`).setValue(remaining)
    const currentPrice = Number(sheet.getRange(`${INVEST_CONFIG.COLUMNS.CURRENT_PRICE}${row}`).getValue()) || 0
    invest_calculateSingle_(sheet, row, currentPrice)
  } else {
    sheet.deleteRow(row)
  }

  // Логирование продажи
  try {
    logOperation_('SELL', name, qtySold, sellPricePerUnit, qtySold * sellPricePerUnit, 'Invest')
  } catch (e) {
    console.error('Invest: ошибка при логировании продажи:', e)
  }

  // Синхронизация с Sales - используем максимальную цену продажи
  const sales = SpreadsheetApp.getActive().getSheetByName('Sales') || null
  if (sales) {
    const sRow = findRowByName_(sales, name, getColumnIndex(SALES_COLUMNS.NAME))
    const currentSell = sRow > 1 ? Number(sales.getRange(sRow, getColumnIndex(SALES_COLUMNS.SELL_PRICE)).getValue()) : null
    const newSell = sellPricePerUnit
    
    if (sRow === -1) {
      // Предмета нет в Sales - создаём новую строку
      const target = Math.max(sales.getLastRow() + 1, DATA_START_ROW)
      const nameCol = getColumnIndex(SALES_COLUMNS.NAME)
      const sellPriceCol = getColumnIndex(SALES_COLUMNS.SELL_PRICE)
      const currentPriceCol = getColumnIndex(SALES_COLUMNS.CURRENT_PRICE)
      
      sales.getRange(target, nameCol).setValue(name)
      sales.getRange(target, sellPriceCol).setValue(newSell)
      
      // Полное форматирование новой строки Sales
      sales_formatNewRow_(sales, target)
      
      const historySheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.HISTORY)
      if (historySheet) {
        const period = getCurrentPricePeriod()
        const priceResult = getHistoryPriceForPeriod_(historySheet, name, period)
        if (priceResult && priceResult.found && priceResult.price > 0) {
          sales.getRange(target, currentPriceCol).setValue(priceResult.price)
          sales_calculateSingle_(sales, target, priceResult.price)
        }
      }
      
      sales_syncMinMaxFromHistory(false)
      sales_syncTrendDaysFromHistory(false)
      sales_syncExtendedAnalyticsFromHistory(false)
    } else {
      // Предмет есть - используем максимальную цену продажи
      const sellPriceCol = getColumnIndex(SALES_COLUMNS.SELL_PRICE)
      if (!Number.isFinite(currentSell) || newSell > currentSell) {
        // Обновляем только если новая цена больше текущей
        sales.getRange(sRow, sellPriceCol).setValue(newSell)
        sales.getRange(sRow, sellPriceCol).setNumberFormat(NUMBER_FORMATS.CURRENCY)
        
        // Пересчитываем просадку с текущей ценой продажи
        const currentPriceCol = getColumnIndex(SALES_COLUMNS.CURRENT_PRICE)
        const currentPrice = Number(sales.getRange(sRow, currentPriceCol).getValue()) || 0
        if (currentPrice > 0) {
          sales_calculateSingle_(sales, sRow, currentPrice)
        }
      }
    }
  }
}

// Очистка цен и пересчетов + синхронизация Min/Max из History (ежедневно)
function invest_dailyReset() {
  const sheet = getInvestSheet_()
  if (!sheet) return
  const lastRow = sheet.getLastRow()
  if (lastRow <= 1) return

  const rangesToClear = [
    `${INVEST_CONFIG.COLUMNS.CURRENT_PRICE}2:${INVEST_CONFIG.COLUMNS.CURRENT_PRICE}${lastRow}`,
    `${INVEST_CONFIG.COLUMNS.TOTAL_INVESTMENT}2:${INVEST_CONFIG.COLUMNS.PROFIT_AFTER_FEE}${lastRow}`,
  ]
  rangesToClear.forEach(range => sheet.getRange(range).clearContent())

  invest_formatGoalColumn_(DATA_START_ROW, lastRow)

  invest_syncMinMaxFromHistory()
  invest_syncTrendDaysFromHistory()
  invest_syncExtendedAnalyticsFromHistory()
  
  // Обновляем аналитику портфеля
  try {
    portfolioStats_update()
    // Отправляем ежедневный отчет в Telegram
    telegram_sendDailyReport()
  } catch (e) {
    console.error('Invest: ошибка при обновлении аналитики портфеля:', e)
  }

  try {
    logAutoAction_('Invest', 'Ежедневный сброс', 'OK')
  } catch (e) {
    console.error('Invest: ошибка при логировании ежедневного сброса:', e)
  }
}

function invest_updateSinglePrice(row) {
  const sheet = getInvestSheet_()
  if (!sheet) return 'error'
  const historySheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.HISTORY)
  if (!historySheet) return 'error'
  
  const itemName = sheet.getRange(`${INVEST_CONFIG.COLUMNS.NAME}${row}`).getValue()
  if (!itemName) return 'error'

  const priceResult = getHistoryPriceForPeriod_(historySheet, itemName, getCurrentPricePeriod())
  
  if (!priceResult.found) {
    return 'error'
  }

  const priceColIndex = getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE)
  invest_calculateSingle_(sheet, row, priceResult.price)
  
  if (priceResult.isOutdated) {
    sheet.getRange(row, priceColIndex).setBackground(COLORS.STABLE)
  } else {
    sheet.getRange(row, priceColIndex).setBackground(null)
  }
  
  return 'updated'
}

function invest_updateImagesAndLinks() {
  updateImagesAndLinksMenu_(INVEST_CONFIG, getInvestSheet_, 'Invest')
}

// Форматирование колонки "Цель" на основе сравнения текущей цены (E) с целью (F)
// Применяется напрямую через код (batch), без условного форматирования
function invest_formatGoalColumn_(startRow = DATA_START_ROW, endRow = null) {
  const sheet = getInvestSheet_()
  if (!sheet) return
  
  if (!endRow) {
    endRow = sheet.getLastRow()
  }
  if (endRow < startRow) return
  
  const rowCount = endRow - startRow + 1
  if (rowCount <= 0) return
  
  // Batch-чтение: читаем колонки текущей цены и цели одним запросом
  const currentPriceCol = getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE)
  const goalCol = getColumnIndex(INVEST_COLUMNS.GOAL)
  const currentPrices = sheet.getRange(startRow, currentPriceCol, rowCount, 1).getValues()
  const goals = sheet.getRange(startRow, goalCol, rowCount, 1).getValues()
  const goalRange = sheet.getRange(startRow, goalCol, rowCount, 1)
  
  // Подготавливаем массивы для batch-форматирования
  const backgrounds = []
  
  for (let i = 0; i < rowCount; i++) {
    const currentPrice = Number(currentPrices[i][0])
    const goal = Number(goals[i][0])
    
    // Проверяем, что обе ячейки содержат валидные положительные числа
    if (Number.isFinite(currentPrice) && currentPrice > 0 && 
        Number.isFinite(goal) && goal > 0) {
      const ratio = currentPrice / goal
      
      if (ratio <= 0.5) {
        // Красный: текущая цена <= 50% цели
        backgrounds.push([COLORS.LOSS])
      } else if (ratio >= 0.8) {
        // Зелёный: текущая цена >= 80% цели
        backgrounds.push([COLORS.PROFIT])
      } else {
        // Белый (без форматирования) для промежуточных значений
        backgrounds.push([null])
      }
    } else {
      // Белый для пустых или невалидных ячеек
      backgrounds.push([null])
    }
  }
  
  // Batch-применение форматирования одним запросом
  goalRange.setBackgrounds(backgrounds)
}

function invest_updateCalculations(row, currentPrice) {
  const sheet = getInvestSheet_()
  if (!sheet) return
  invest_calculateSingle_(sheet, row, currentPrice)
}

// Добавить или обновить позицию в Invest (усреднение цены)
function invest_addOrUpdatePosition_(name, qtyToAdd, buyPricePerUnit) {
  const sheet = getInvestSheet_()
  if (!sheet) return
  const row = findRowByName_(sheet, name, 2)
  if (row === -1) {
    // Новый предмет - создаём строку
    const target = Math.max(sheet.getLastRow() + 1, DATA_START_ROW)
    sheet.getRange(target, getColumnIndex(INVEST_COLUMNS.NAME)).setValue(name)
    sheet.getRange(target, getColumnIndex(INVEST_COLUMNS.QUANTITY)).setValue(qtyToAdd)
    sheet.getRange(target, getColumnIndex(INVEST_COLUMNS.BUY_PRICE)).setValue(buyPricePerUnit)
    
    // Пытаемся получить текущую цену из History
    let currentPrice = 0
    const historySheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.HISTORY)
    if (historySheet) {
      const period = getCurrentPricePeriod()
      const priceResult = getHistoryPriceForPeriod_(historySheet, name, period)
      if (priceResult && priceResult.found && priceResult.price > 0) {
        currentPrice = priceResult.price
      }
    }
    
    // Устанавливаем текущую цену и вычисляем расчёты
    sheet.getRange(target, getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE)).setValue(currentPrice)
    invest_calculateSingle_(sheet, target, currentPrice)
    
    // Ссылки/картинки
    setImageAndLink_(sheet, target, INVEST_CONFIG.STEAM_APPID, name, INVEST_CONFIG.COLUMNS)
    
    // Применяем форматирование только для новой строки
    invest_formatNewRow_(sheet, target)
    
    return { created: true, row: target }
  }
  // есть позиция — усредняем цену покупки
  const quantityCol = getColumnIndex(INVEST_COLUMNS.QUANTITY)
  const buyPriceCol = getColumnIndex(INVEST_COLUMNS.BUY_PRICE)
  const currentQty = Number(sheet.getRange(row, quantityCol).getValue()) || 0
  const currentBuy = Number(sheet.getRange(row, buyPriceCol).getValue()) || 0
  const newQty = currentQty + qtyToAdd
  const newAvg = newQty > 0 ? (currentQty * currentBuy + qtyToAdd * buyPricePerUnit) / newQty : buyPricePerUnit
  sheet.getRange(row, quantityCol).setValue(newQty)
  sheet.getRange(row, buyPriceCol).setValue(newAvg)
  const currentPriceCol = getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE)
  invest_calculateSingle_(sheet, row, Number(sheet.getRange(row, currentPriceCol).getValue()) || 0)
  return { created: false, row }
}

function invest_formatTable() {
  const sheet = getOrCreateInvestSheet_()
  const lastRow = sheet.getLastRow()

  // Используем константы для заголовков
  const headers = HEADERS.INVEST // 18 колонок (убрали H, K и DAYS_CHANGE)

  sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers])
  
  // Проверяем и исправляем заголовок "Потенциал" на "Потенциал (P85)" если нужно
  const potentialColIndex = getColumnIndex(INVEST_COLUMNS.POTENTIAL)
  const currentPotentialHeader = sheet.getRange(HEADER_ROW, potentialColIndex).getValue()
  if (currentPotentialHeader && currentPotentialHeader !== 'Потенциал (P85)') {
    sheet.getRange(HEADER_ROW, potentialColIndex).setValue('Потенциал (P85)')
  }

  formatHeaderRange_(sheet.getRange(HEADER_ROW, 1, 1, headers.length))

  sheet.setRowHeight(HEADER_ROW, HEADER_HEIGHT)
  if (lastRow > 1) sheet.setRowHeights(DATA_START_ROW, lastRow - 1, ROW_HEIGHT)

  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.IMAGE), COLUMN_WIDTHS.IMAGE)
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.NAME), COLUMN_WIDTHS.NAME)
  sheet.setColumnWidths(3, 11, COLUMN_WIDTHS.WIDE) // C-M (11 колонок после удаления H и K)
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.TREND), COLUMN_WIDTHS.MEDIUM) // Объединенный формат, шире
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.PHASE), COLUMN_WIDTHS.WIDE)
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.POTENTIAL), COLUMN_WIDTHS.MEDIUM)
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.RECOMMENDATION), COLUMN_WIDTHS.EXTRA_WIDE)

  if (lastRow > 1) {
    sheet.getRange(`D2:I${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // D-G, H (с комиссией), I (Профит)
    sheet.getRange(`F2:F${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Цель - явное форматирование для гарантии
    sheet.getRange(`J2:J${lastRow}`).setNumberFormat(NUMBER_FORMATS.PERCENT) // Прибыль % с комиссией (было L)
    sheet.getRange(`L2:M${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Min, Max (было O-P)
    // Форматирование колонки Потенциал (Q) как процент с знаком "+"
    const potentialCol = getColumnIndex(INVEST_COLUMNS.POTENTIAL)
    sheet.getRange(DATA_START_ROW, potentialCol, lastRow - 1, 1).setNumberFormat('+0%;-0%;"—"')

    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - 1, headers.length)
    dataRange.setVerticalAlignment('middle').setWrap(true)

    sheet.getRange(`A2:A${lastRow}`).setHorizontalAlignment('center')
    sheet.getRange(`B2:B${lastRow}`).setHorizontalAlignment('left')
    sheet.getRange(`C2:R${lastRow}`).setHorizontalAlignment('center') // До R (было S, убрали DAYS_CHANGE)
  }

  if (lastRow > 1) {
    const profitRanges = sheet.getRange(`I2:J${lastRow}`) // Профит и Прибыль % с комиссией (было J-L)
    const trendCol = getColumnIndex(INVEST_COLUMNS.TREND)
    const phaseCol = getColumnIndex(INVEST_COLUMNS.PHASE)
    const recommendationCol = getColumnIndex(INVEST_COLUMNS.RECOMMENDATION)
    
    applyAnalyticsFormatting_(sheet, {
      trendCol,
      phaseCol,
      recommendationCol,
      profitRange: profitRanges
    }, lastRow)
    
    // Форматируем колонку цели напрямую через код (batch)
    // Это надёжнее, чем условное форматирование с формулами
    invest_formatGoalColumn_(DATA_START_ROW, lastRow)
  } else {
    sheet.setConditionalFormatRules([])
  }

  sheet.setFrozenRows(HEADER_ROW)
  // Добавляем колонку кнопки «Продать» если отсутствует
  const lastCol = sheet.getLastColumn()
  const sellHeader = 'Продать'
  let sellCol = null
  for (let c = 1; c <= lastCol; c++) {
    if (sheet.getRange(1, c).getValue() === sellHeader) {
      sellCol = c; break
    }
  }
  if (!sellCol) {
    sellCol = lastCol + 1
    sheet.getRange(1, sellCol).setValue(sellHeader)
  }
  if (lastRow > 1) {
    const rng = sheet.getRange(DATA_START_ROW, sellCol, lastRow - 1, 1)
    rng.insertCheckboxes()
    rng.setHorizontalAlignment('center')
  }
  // Стиль шапки «Продать» как и остальные
  formatHeaderRange_(sheet.getRange(HEADER_ROW, sellCol, 1, 1))
  SpreadsheetApp.getUi().alert('Форматирование завершено (Invest)')
}


function invest_findDuplicates() {
  const sheet = getInvestSheet_()
  if (!sheet) return
  const res = highlightDuplicatesByName_(sheet, DATA_START_ROW, COLORS.DUPLICATE)
  SpreadsheetApp.getUi().alert(res.duplicates ? `Найдено повторов: ${res.duplicates}` : 'Повторов не найдено')
}

// Синхронизация Min/Max из листа History по названию (использует универсальную функцию)
function invest_syncMinMaxFromHistory(updateAll = true) {
  const sheet = getInvestSheet_()
  if (!sheet) return

  // INVEST_COLUMNS.MIN_PRICE = 'N', INVEST_COLUMNS.MAX_PRICE = 'O'
  const minColIndex = getColumnIndex(INVEST_COLUMNS.MIN_PRICE)
  const maxColIndex = getColumnIndex(INVEST_COLUMNS.MAX_PRICE)
  
  return syncMinMaxFromHistoryUniversal_(sheet, minColIndex, maxColIndex, updateAll)
}

// Синхронизация Тренд из листа History по названию (теперь объединенный формат)
function invest_syncTrendDaysFromHistory(updateAll = true) {
  const sheet = getInvestSheet_()
  if (!sheet) return

  // INVEST_COLUMNS.TREND = 'N' (теперь содержит объединенный формат "🟥 Падает 35 дн.")
  const trendColIndex = getColumnIndex(INVEST_COLUMNS.TREND)
  
  return syncTrendFromHistoryUniversal_(sheet, trendColIndex, updateAll)
}

// Синхронизация расширенной аналитики (Фаза/Потенциал/Рекомендация) из History
function invest_syncExtendedAnalyticsFromHistory(updateAll = true) {
  const sheet = getInvestSheet_()
  if (!sheet) return

  // INVEST_COLUMNS: PHASE = 'R', POTENTIAL = 'S', RECOMMENDATION = 'T'
  const phaseColIndex = getColumnIndex(INVEST_COLUMNS.PHASE)
  const potentialColIndex = getColumnIndex(INVEST_COLUMNS.POTENTIAL)
  const recommendationColIndex = getColumnIndex(INVEST_COLUMNS.RECOMMENDATION)
  
  return syncExtendedAnalyticsFromHistoryUniversal_(sheet, phaseColIndex, potentialColIndex, recommendationColIndex, updateAll)
}

/**
 * Комплексное обновление всей аналитики (Min/Max + Тренд/Дней смены + Фаза/Потенциал/Рекомендация)
 */
function invest_updateAllAnalytics() {
  updateAllAnalyticsManual_(
    'Invest',
    invest_syncMinMaxFromHistory,
    invest_syncTrendDaysFromHistory,
    invest_syncExtendedAnalyticsFromHistory
  )
}

