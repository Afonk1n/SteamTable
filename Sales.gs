// Sales module
// Используем константы из Constants.gs
const SALES_CONFIG = {
  STEAM_APPID: STEAM_APP_ID,
  UPDATE_INTERVAL_MINUTES: UPDATE_INTERVALS.PRICES_MINUTES,
  COLUMNS: SALES_COLUMNS,
}

const SALES_SHEET_NAME = SHEET_NAMES.SALES

// Форматирование новой строки Sales (при добавлении из Invest)
function sales_formatNewRow_(sheet, row) {
  if (row <= HEADER_ROW) return
  const name = sheet.getRange(`B${row}`).getValue()
  if (!name) return
  
  // Базовое форматирование строки (A-M = 13 колонок без статуса)
  const numCols = getColumnIndex(SALES_COLUMNS.RECOMMENDATION)
  sheet.getRange(row, 1, 1, numCols).setVerticalAlignment('middle').setHorizontalAlignment('center')
  sheet.getRange(`B${row}`).setHorizontalAlignment('left')
  
  // Форматы чисел (используем константы)
  sheet.getRange(`C${row}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Цена продажи
  sheet.getRange(`D${row}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Текущая цена
  sheet.getRange(`E${row}`).setNumberFormat(NUMBER_FORMATS.PERCENT) // Процент просадки
  sheet.getRange(`G${row}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Min цена (было H)
  sheet.getRange(`H${row}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Max цена (было I)
  // Колонка Дней смены (была J) удалена, ее данные теперь в Тренд (I)
  // Форматирование колонки Потенциал (K) как процент с знаком "+"
  const potentialCol = getColumnIndex(SALES_COLUMNS.POTENTIAL)
  sheet.getRange(row, potentialCol).setNumberFormat('+0%;-0%;"—"')
  
  // Добавляем изображение и ссылку
  setImageAndLink_(sheet, row, SALES_CONFIG.STEAM_APPID, name, SALES_CONFIG.COLUMNS)
  
  // Устанавливаем высоту строки
  sheet.setRowHeight(row, ROW_HEIGHT)
}

// Функции getSalesSheet_ и getOrCreateSalesSheet_ перенесены в SheetService.gs

function sales_dailyReset() {
  const sheet = getSalesSheet_()
  if (!sheet) return
  const lastRow = sheet.getLastRow()
  if (lastRow <= 1) return

  const rangesToClear = [
    `${SALES_CONFIG.COLUMNS.CURRENT_PRICE}2:${SALES_CONFIG.COLUMNS.PRICE_DROP}${lastRow}`,
  ]
  rangesToClear.forEach(range => sheet.getRange(range).clearContent())

  sales_syncMinMaxFromHistory()
  sales_syncTrendDaysFromHistory()
  sales_syncExtendedAnalyticsFromHistory()

  try {
    logAutoAction_('Sales', 'Ежедневный сброс', 'OK')
  } catch (e) {
    console.error('Sales: ошибка при логировании ежедневного сброса:', e)
  }
}

function sales_updateSinglePrice(row) {
  const sheet = getSalesSheet_()
  if (!sheet) return 'error'
  const historySheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.HISTORY)
  if (!historySheet) return 'error'
  
  const itemName = sheet.getRange(`${SALES_CONFIG.COLUMNS.NAME}${row}`).getValue()
  if (!itemName) return 'error'

  const priceResult = getHistoryPriceForPeriod_(historySheet, itemName, getCurrentPricePeriod())
  
  if (!priceResult.found) {
    return 'error'
  }

  const priceColIndex = getColumnIndex(SALES_COLUMNS.CURRENT_PRICE)
  sales_calculateSingle_(sheet, row, priceResult.price)
  
  if (priceResult.isOutdated) {
    sheet.getRange(row, priceColIndex).setBackground(COLORS.STABLE)
  } else {
    sheet.getRange(row, priceColIndex).setBackground(null)
  }
  
  return 'updated'
}

function sales_updateCalculations(row, currentPrice) {
  const sheet = getSalesSheet_()
  if (!sheet) return
  sales_calculateSingle_(sheet, row, currentPrice)
}

function sales_updateImagesAndLinks() {
  updateImagesAndLinksMenu_(SALES_CONFIG, getSalesSheet_, 'Sales')
}

function sales_formatTable() {
  const sheet = getOrCreateSalesSheet_()
  const lastRow = sheet.getLastRow()

  // Используем константы для заголовков
  const headers = HEADERS.SALES
  
  // Проверка на undefined
  if (!headers || !Array.isArray(headers) || headers.length === 0) {
    console.error('Sales: HEADERS.SALES не определен или пуст')
    SpreadsheetApp.getUi().alert('Ошибка: заголовки Sales не определены')
    return
  }

  sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers])
  
  // Проверяем и исправляем заголовок "Потенциал" на "Потенциал (P85)" если нужно
  const potentialColIndex = getColumnIndex(SALES_COLUMNS.POTENTIAL)
  const currentPotentialHeader = sheet.getRange(HEADER_ROW, potentialColIndex).getValue()
  if (currentPotentialHeader && currentPotentialHeader !== 'Потенциал (P85)') {
    sheet.getRange(HEADER_ROW, potentialColIndex).setValue('Потенциал (P85)')
  }

  formatHeaderRange_(sheet.getRange(HEADER_ROW, 1, 1, headers.length))

  sheet.setRowHeight(HEADER_ROW, HEADER_HEIGHT)
  if (lastRow > 1) sheet.setRowHeights(DATA_START_ROW, lastRow - 1, ROW_HEIGHT)

  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.IMAGE), COLUMN_WIDTHS.IMAGE)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.NAME), COLUMN_WIDTHS.NAME)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.SELL_PRICE), COLUMN_WIDTHS.WIDE)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.CURRENT_PRICE), COLUMN_WIDTHS.WIDE)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.PRICE_DROP), COLUMN_WIDTHS.WIDE)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.LINK), COLUMN_WIDTHS.NARROW)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.MIN_PRICE), COLUMN_WIDTHS.MEDIUM)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.MAX_PRICE), COLUMN_WIDTHS.WIDE)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.TREND), COLUMN_WIDTHS.WIDE) // Объединенный формат, шире для текста
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.PHASE), COLUMN_WIDTHS.WIDE)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.POTENTIAL), COLUMN_WIDTHS.MEDIUM)
  sheet.setColumnWidth(getColumnIndex(SALES_COLUMNS.RECOMMENDATION), COLUMN_WIDTHS.EXTRA_WIDE)

  if (lastRow > 1) {
    sheet.getRange(`C2:D${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY)
    sheet.getRange(`E2:E${lastRow}`).setNumberFormat(NUMBER_FORMATS.PERCENT)
    sheet.getRange(`G2:H${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Min, Max (было H-I)
    // Колонка Дней смены (была J) удалена, ее данные теперь в Тренд (I)
    // Форматирование колонки Потенциал (K) как процент с знаком "+"
    const potentialCol = getColumnIndex(SALES_COLUMNS.POTENTIAL)
    sheet.getRange(DATA_START_ROW, potentialCol, lastRow - 1, 1).setNumberFormat('+0%;-0%;"—"')

    sheet
      .getRange(DATA_START_ROW, 1, lastRow - 1, headers.length)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')

    sheet.getRange(`B2:B${lastRow}`).setHorizontalAlignment('left')

    const dropRange = sheet.getRange(`E2:E${lastRow}`)
    const trendCol = getColumnIndex(SALES_COLUMNS.TREND)
    const phaseCol = getColumnIndex(SALES_COLUMNS.PHASE)
    const recommendationCol = getColumnIndex(SALES_COLUMNS.RECOMMENDATION)
    
    applyAnalyticsFormatting_(sheet, {
      trendCol,
      phaseCol,
      recommendationCol,
      dropRange
    }, lastRow)
  }
  else {
    sheet.setConditionalFormatRules([])
  }

  sheet.setFrozenRows(HEADER_ROW)
  SpreadsheetApp.getUi().alert('Форматирование завершено (Sales)')
}


function sales_findDuplicates() {
  const sheet = getSalesSheet_()
  if (!sheet) return
  const res = highlightDuplicatesByName_(sheet, DATA_START_ROW, COLORS.DUPLICATE)
  SpreadsheetApp.getUi().alert(res.duplicates ? `Найдено повторов: ${res.duplicates}` : 'Повторов не найдено')
}

function sales_syncMinMaxFromHistory(updateAll = true) {
  const sheet = getSalesSheet_()
  if (!sheet) return

  // SALES_COLUMNS.MIN_PRICE = 'G', SALES_COLUMNS.MAX_PRICE = 'H'
  const minColIndex = getColumnIndex(SALES_COLUMNS.MIN_PRICE)
  const maxColIndex = getColumnIndex(SALES_COLUMNS.MAX_PRICE)
  
  return syncMinMaxFromHistoryUniversal_(sheet, minColIndex, maxColIndex, updateAll)
}

// Синхронизация Тренд/Дней смены из листа History по названию
function sales_syncTrendDaysFromHistory(updateAll = true) {
  const sheet = getSalesSheet_()
  if (!sheet) return

  // SALES_COLUMNS.TREND = 'I' (теперь содержит объединенный формат "🟥 Падает 35 дн.")
  const trendColIndex = getColumnIndex(SALES_COLUMNS.TREND)
  
  return syncTrendFromHistoryUniversal_(sheet, trendColIndex, updateAll)
}

// Синхронизация расширенной аналитики (Фаза/Потенциал/Рекомендация) из History
function sales_syncExtendedAnalyticsFromHistory(updateAll = true) {
  const sheet = getSalesSheet_()
  if (!sheet) return

  // SALES_COLUMNS: PHASE = 'K', POTENTIAL = 'L', RECOMMENDATION = 'M'
  const phaseColIndex = getColumnIndex(SALES_COLUMNS.PHASE)
  const potentialColIndex = getColumnIndex(SALES_COLUMNS.POTENTIAL)
  const recommendationColIndex = getColumnIndex(SALES_COLUMNS.RECOMMENDATION)
  
  return syncExtendedAnalyticsFromHistoryUniversal_(sheet, phaseColIndex, potentialColIndex, recommendationColIndex, updateAll)
}

/**
 * Комплексное обновление всей аналитики (Min/Max + Тренд/Дней смены + Фаза/Потенциал/Рекомендация)
 */
function sales_updateAllAnalytics() {
  updateAllAnalyticsManual_(
    'Sales',
    sales_syncMinMaxFromHistory,
    sales_syncTrendDaysFromHistory,
    sales_syncExtendedAnalyticsFromHistory
  )
}
