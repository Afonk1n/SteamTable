// Invest module
// Используем константы из Constants.gs
const INVEST_CONFIG = {
  STEAM_APPID: STEAM_APP_ID,
  STEAM_FEE: STEAM_FEE,
  COLUMNS: INVEST_COLUMNS,
}

// Форматирование новой строки Invest (при добавлении из History)
function invest_formatNewRow_(sheet, row) {
  const numberFormatConfig = {
    QUANTITY: NUMBER_FORMATS.INTEGER,      // C: Количество
    BUY_PRICE: NUMBER_FORMATS.CURRENCY,    // D: Цена покупки
    CURRENT_PRICE: NUMBER_FORMATS.CURRENCY, // E: Текущая цена
    GOAL: NUMBER_FORMATS.CURRENCY,         // F: Цель
    TOTAL_INVESTMENT: NUMBER_FORMATS.CURRENCY, // G: Общие вложения
    CURRENT_VALUE_AFTER_FEE: NUMBER_FORMATS.CURRENCY, // H: Текущая стоимость с комиссией
    PROFIT: NUMBER_FORMATS.CURRENCY,       // I: Профит
    PROFIT_AFTER_FEE: NUMBER_FORMATS.PERCENT, // J: Прибыль % с комиссией
    MIN_PRICE: NUMBER_FORMATS.CURRENCY,    // L: Min цена
    MAX_PRICE: NUMBER_FORMATS.CURRENCY     // M: Max цена
    // N-Z: Investment Score, Рекомендация, Фаза, Потенциал, Тренд, Дней смены, Hero Trend, Метрики, Risk Level, Чекбоксы
  }
  
  formatNewRowUniversal_(sheet, row, INVEST_CONFIG, numberFormatConfig, false)
}

// Функции getInvestSheet_ и getOrCreateInvestSheet_ перенесены в SheetService.gs

function invest_applySale(row, qtySold, sellPricePerUnit) {
  const sheet = getInvestSheet_()
  if (!sheet) return
  const name = sheet.getRange(`${INVEST_CONFIG.COLUMNS.NAME}${row}`).getValue()
  const qtyAvailable = Number(sheet.getRange(`${INVEST_CONFIG.COLUMNS.QUANTITY}${row}`).getValue())
  
  // Проверка валидности данных
  if (!name || !Number.isFinite(qtyAvailable) || qtyAvailable <= 0) {
    console.error('Invest: некорректные данные для продажи в строке', row)
    return
  }
  
  // Проверка, что продаваемое количество не превышает доступное
  if (!Number.isFinite(qtySold) || qtySold <= 0 || qtySold > qtyAvailable) {
    console.error(`Invest: некорректное количество для продажи: ${qtySold} (доступно: ${qtyAvailable})`)
    SpreadsheetApp.getUi().alert(`Ошибка: нельзя продать ${qtySold} шт., доступно только ${qtyAvailable} шт.`)
    return
  }

  const remaining = qtyAvailable - qtySold
  if (remaining > 0) {
    sheet.getRange(`${INVEST_CONFIG.COLUMNS.QUANTITY}${row}`).setValue(remaining)
    const currentPrice = Number(sheet.getRange(`${INVEST_CONFIG.COLUMNS.CURRENT_PRICE}${row}`).getValue()) || 0
    invest_calculateSingle_(sheet, row, currentPrice)
  } else {
    // remaining === 0 - удаляем строку
    sheet.deleteRow(row)
  }

  // Логирование продажи
  try {
    logOperation_('SELL', name, qtySold, sellPricePerUnit, qtySold * sellPricePerUnit, 'Invest')
  } catch (e) {
    console.error('Invest: ошибка при логировании продажи:', e)
  }

  // Синхронизация с Sales - используем максимальную цену продажи
  const sales = getSalesSheet_()
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
      
      const historySheet = getHistorySheet_()
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

  // ИСПРАВЛЕНИЕ: Синхронизация аналитики убрана отсюда, так как она выполняется в syncPricesFromHistoryToInvestAndSales()
  // Это предотвращает двойную синхронизацию аналитики
  
  // Примечание: История портфеля теперь сохраняется автоматически в unified_priceUpdate() после дневного сбора

  try {
    logAutoAction_('Invest', 'Ежедневный сброс', 'OK')
  } catch (e) {
    console.error('Invest: ошибка при логировании ежедневного сброса:', e)
  }
}

function invest_updateSinglePrice(row) {
  const sheet = getInvestSheet_()
  if (!sheet) return 'error'
  const historySheet = getHistorySheet_()
  if (!historySheet) return 'error'
  
  const itemName = sheet.getRange(`${INVEST_CONFIG.COLUMNS.NAME}${row}`).getValue()
  if (!itemName) return 'error'

  const priceResult = getHistoryPriceForPeriod_(historySheet, itemName, getCurrentPricePeriod())
  
  if (!priceResult.found) {
    console.warn(`Invest: цена не найдена в History для "${itemName}"`)
    return 'error'
  }

  // ВАЛИДАЦИЯ: Проверяем цену перед использованием
  const priceValidation = validatePrice_(priceResult.price, itemName)
  if (!priceValidation.valid) {
    console.warn(`Invest: некорректная цена для "${itemName}": ${priceResult.price}, ошибка: ${priceValidation.error}`)
    return 'error'
  }

  const priceColIndex = getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE)
  invest_calculateSingle_(sheet, row, priceValidation.price)
  
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
  
  // ВАЛИДАЦИЯ: Проверяем цену перед расчетом
  const priceValidation = validatePrice_(currentPrice, `строка ${row}`)
  if (!priceValidation.valid) {
    console.warn(`Invest: некорректная цена для пересчета в строке ${row}: ${currentPrice}`)
    return
  }
  
  invest_calculateSingle_(sheet, row, priceValidation.price)
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
    const historySheet = getHistorySheet_()
    if (historySheet) {
      const period = getCurrentPricePeriod()
      const priceResult = getHistoryPriceForPeriod_(historySheet, name, period)
      if (priceResult && priceResult.found && priceResult.price > 0) {
        // ВАЛИДАЦИЯ: Проверяем цену перед использованием
        const priceValidation = validatePrice_(priceResult.price, name)
        if (priceValidation.valid) {
          currentPrice = priceValidation.price
        } else {
          console.warn(`Invest: некорректная цена для "${name}" при добавлении позиции: ${priceResult.price}`)
        }
      }
    }
    
    // ВАЛИДАЦИЯ: Проверяем цену покупки перед установкой
    const buyPriceValidation = validatePrice_(buyPricePerUnit, `${name} (buyPrice)`)
    if (!buyPriceValidation.valid) {
      console.warn(`Invest: некорректная цена покупки для "${name}": ${buyPricePerUnit}`)
      return
    }
    
    // Устанавливаем текущую цену и вычисляем расчёты
    sheet.getRange(target, getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE)).setValue(currentPrice)
    sheet.getRange(target, getColumnIndex(INVEST_COLUMNS.BUY_PRICE)).setValue(buyPriceValidation.price)
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
  let currentQty = Number(sheet.getRange(row, quantityCol).getValue()) || 0
  let currentBuy = Number(sheet.getRange(row, buyPriceCol).getValue()) || 0
  
  // Проверка валидности данных
  if (!Number.isFinite(currentQty) || currentQty < 0) {
    console.error('Invest: некорректное текущее количество в строке', row)
    currentQty = 0 // Исправляем на 0
  }
  if (!Number.isFinite(currentBuy) || currentBuy < 0) {
    console.error('Invest: некорректная текущая цена покупки в строке', row)
    currentBuy = 0 // Исправляем на 0
  }
  
  const newQty = currentQty + qtyToAdd
  
  // Усредняем цену: если newQty > 0, используем формулу усреднения, иначе используем новую цену
  const newAvg = newQty > 0 && Number.isFinite(currentBuy) && currentBuy > 0
    ? (currentQty * currentBuy + qtyToAdd * buyPricePerUnit) / newQty
    : buyPricePerUnit
  sheet.getRange(row, quantityCol).setValue(newQty)
  sheet.getRange(row, buyPriceCol).setValue(newAvg)
  const currentPriceCol = getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE)
  invest_calculateSingle_(sheet, row, Number(sheet.getRange(row, currentPriceCol).getValue()) || 0)
  return { created: false, row }
}

function invest_formatTable() {
  const sheet = getOrCreateInvestSheet_()
  const headers = HEADERS.INVEST // 28 колонок (новая структура)
  
  if (!headers || !Array.isArray(headers) || headers.length === 0) {
    console.error('Invest: HEADERS.INVEST не определен или пуст')
    SpreadsheetApp.getUi().alert('Ошибка: HEADERS.INVEST не определен в Constants.gs')
    return
  }
  
  // Базовое форматирование таблицы
  const lastRow = formatTableBase_(sheet, headers, INVEST_COLUMNS, getInvestSheet_, 'Invest')
  if (lastRow === 0) return

  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.IMAGE), COLUMN_WIDTHS.IMAGE)
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.NAME), COLUMN_WIDTHS.NAME)
  sheet.setColumnWidths(3, 9, COLUMN_WIDTHS.WIDE) // C-K (9 колонок)
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.MIN_PRICE), COLUMN_WIDTHS.MEDIUM) // L
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.MAX_PRICE), COLUMN_WIDTHS.MEDIUM) // M
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.INVESTMENT_SCORE), 130) // N
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.RECOMMENDATION), COLUMN_WIDTHS.EXTRA_WIDE) // O
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.PHASE), COLUMN_WIDTHS.WIDE) // P
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.POTENTIAL), COLUMN_WIDTHS.MEDIUM) // Q
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.TREND), COLUMN_WIDTHS.WIDE) // R - Тренд (объединенный формат: "🟨 Боковик 39 д.")
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.HERO_TREND), COLUMN_WIDTHS.MEDIUM) // S (перемещено из T, убрали DAYS_CHANGE)
  sheet.setColumnWidth(getColumnIndex(INVEST_COLUMNS.RISK_LEVEL), COLUMN_WIDTHS.MEDIUM) // T (перемещено из U)

  if (lastRow > 1) {
    sheet.getRange(`D2:I${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // D-G, H (с комиссией), I (Профит)
    sheet.getRange(`F2:F${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Цель - явное форматирование для гарантии
    sheet.getRange(`J2:J${lastRow}`).setNumberFormat(NUMBER_FORMATS.PERCENT) // Прибыль % с комиссией
    sheet.getRange(`L2:M${lastRow}`).setNumberFormat(NUMBER_FORMATS.CURRENCY) // Min, Max
    // Форматирование колонки Потенциал (Q) как процент с знаком "+"
    const potentialCol = getColumnIndex(INVEST_COLUMNS.POTENTIAL)
    sheet.getRange(DATA_START_ROW, potentialCol, lastRow - 1, 1).setNumberFormat('+0%;-0%;"—"')
    // Метрики удалены из отображения (остаются в коде для расчетов)

    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - 1, headers.length)
    dataRange.setVerticalAlignment('middle').setWrap(true)

    sheet.getRange(`A2:A${lastRow}`).setHorizontalAlignment('center')
    sheet.getRange(`B2:B${lastRow}`).setHorizontalAlignment('left')
    sheet.getRange(`C2:AA${lastRow}`).setHorizontalAlignment('center') // До AA (чекбокс Продать)
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

  // Заморозка строки уже выполнена в formatTableBase_()
  // Добавляем колонку чекбокса «Продать» если отсутствует (убрали «Купить?»)
  const lastCol = sheet.getLastColumn()
  const sellHeader = 'Продать'
  let sellCol = null
  
  for (let c = 1; c <= lastCol; c++) {
    const header = sheet.getRange(1, c).getValue()
    if (header === sellHeader) sellCol = c
  }
  
  if (!sellCol) {
    sellCol = getColumnIndex(INVEST_COLUMNS.SELL_CHECKBOX)
    sheet.getRange(1, sellCol).setValue(sellHeader)
    formatHeaderRange_(sheet.getRange(HEADER_ROW, sellCol, 1, 1))
    if (lastRow > 1) {
      const rng = sheet.getRange(DATA_START_ROW, sellCol, lastRow - 1, 1)
      rng.insertCheckboxes()
      rng.setHorizontalAlignment('center')
    }
  }
  
  console.log('Invest: форматирование завершено')
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

  // INVEST_COLUMNS.MIN_PRICE = 'L', INVEST_COLUMNS.MAX_PRICE = 'M'
  const minColIndex = getColumnIndex(INVEST_COLUMNS.MIN_PRICE)
  const maxColIndex = getColumnIndex(INVEST_COLUMNS.MAX_PRICE)
  
  return syncMinMaxFromHistoryUniversal_(sheet, minColIndex, maxColIndex, updateAll)
}

// Синхронизация Тренд из листа History по названию (теперь объединенный формат)
function invest_syncTrendDaysFromHistory(updateAll = true) {
  const sheet = getInvestSheet_()
  if (!sheet) return

  // INVEST_COLUMNS.TREND = 'R' (теперь содержит объединенный формат "🟥 Падает 35 дн.")
  const trendColIndex = getColumnIndex(INVEST_COLUMNS.TREND)
  
  return syncTrendFromHistoryUniversal_(sheet, trendColIndex, updateAll)
}

// Синхронизация расширенной аналитики (Фаза/Потенциал/Рекомендация) из History
function invest_syncExtendedAnalyticsFromHistory(updateAll = true) {
  const sheet = getInvestSheet_()
  if (!sheet) return

  // INVEST_COLUMNS: PHASE = 'P', POTENTIAL = 'Q', RECOMMENDATION = 'O'
  const phaseColIndex = getColumnIndex(INVEST_COLUMNS.PHASE)
  const potentialColIndex = getColumnIndex(INVEST_COLUMNS.POTENTIAL)
  const recommendationColIndex = getColumnIndex(INVEST_COLUMNS.RECOMMENDATION)
  
  return syncExtendedAnalyticsFromHistoryUniversal_(sheet, phaseColIndex, potentialColIndex, recommendationColIndex, updateAll)
}

/**
 * Комплексное обновление всей аналитики (Min/Max + Тренд (объединенный) + Фаза/Потенциал/Рекомендация)
 */
function invest_updateAllAnalytics() {
  updateAllAnalyticsManual_(
    'Invest',
    invest_syncMinMaxFromHistory,
    invest_syncTrendDaysFromHistory,
    invest_syncExtendedAnalyticsFromHistory
  )
}

// ===== СИСТЕМА ИНВЕСТИЦИОННЫХ РЕКОМЕНДАЦИЙ =====

/**
 * Получение данных из SteamWebAPI для расчета метрик
 * @param {Array<string>} itemNames - Массив названий предметов
 * @returns {Object} Объект {itemName: itemData}
 */
function invest_updateMetricsFromSteamWebAPI(itemNames) {
  const itemsData = {}
  
  // Batch запросы (до 50 предметов за раз)
  const batchSize = API_CONFIG.STEAM_WEB_API.MAX_ITEMS_PER_REQUEST
  for (let i = 0; i < itemNames.length; i += batchSize) {
    const batch = itemNames.slice(i, i + batchSize)
    const result = steamWebAPI_fetchItems(batch, 'dota2')
    if (result.ok && result.items) {
      result.items.forEach(item => {
        if (item.marketname) {
          itemsData[item.marketname] = item
        }
      })
    }
    // Задержка между batch запросами
    if (i + batchSize < itemNames.length) {
      Utilities.sleep(LIMITS.METRICS_UPDATE_DELAY_MS)
    }
  }
  
  return itemsData
}

/**
 * Расчет всех метрик для позиций в Invest
 */
function invest_calculateAllMetrics() {
  const sheet = getInvestSheet_()
  if (!sheet) return
  
  const lastRow = sheet.getLastRow()
  if (lastRow < DATA_START_ROW) return
  
  const itemNames = sheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.NAME), lastRow - HEADER_ROW, 1).getValues()
  const itemNamesList = itemNames.map(row => String(row[0] || '').trim()).filter(name => name)
  
  // Получаем данные из SteamWebAPI
  const itemsData = invest_updateMetricsFromSteamWebAPI(itemNamesList)
  
  // Получаем маппинги героев
  const mappings = heroMapping_getAllMappings()
  
  // Получаем историю цен из History
  const historySheet = getHistorySheet_()
  
  // Подготовка данных для batch-операций
  const liquidityScores = []
  const demandRatios = []
  const priceMomenta = []
  const salesTrends = []
  const volatilityIndices = []
  const heroTrends = []
  const historyNames = historySheet ? historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), historySheet.getLastRow() - HEADER_ROW, 1).getValues() : []
  
  // Рассчитываем метрики для всех строк
  const startedAt = Date.now()
  const MAX_EXECUTION_TIME_MS = 300000 // 5 минут
  
  for (let i = 0; i < itemNames.length; i++) {
    // Проверка таймаута
    if (Date.now() - startedAt > MAX_EXECUTION_TIME_MS) {
      console.warn(`Invest: превышено время выполнения calculateAllMetrics (${MAX_EXECUTION_TIME_MS}ms), прервано на строке ${i + 1}`)
      break
    }
    
    try {
      const itemName = String(itemNames[i][0] || '').trim()
      if (!itemName) {
        liquidityScores.push([null])
        demandRatios.push([null])
        priceMomenta.push([null])
        salesTrends.push([null])
        volatilityIndices.push([null])
        heroTrends.push([null])
        continue
      }
      
      const itemData = itemsData[itemName]
      if (!itemData) {
        console.warn(`Invest: нет данных из SteamWebAPI для "${itemName}"`)
        liquidityScores.push([null])
        demandRatios.push([null])
        priceMomenta.push([null])
        salesTrends.push([null])
        volatilityIndices.push([null])
        heroTrends.push([null])
        continue
      }
      
      // ВАЛИДАЦИЯ: Проверяем валидность цен в itemData
      if (itemData.pricelatest !== undefined && itemData.pricelatest !== null) {
        const priceValidation = validatePrice_(itemData.pricelatest, `${itemName} (pricelatest)`)
        if (!priceValidation.valid) {
          console.warn(`Invest: некорректная цена pricelatest для "${itemName}": ${itemData.pricelatest}, пропускаем расчет метрик`)
          liquidityScores.push([null])
          demandRatios.push([null])
          priceMomenta.push([null])
          salesTrends.push([null])
          volatilityIndices.push([null])
          heroTrends.push([null])
          continue
        }
      }
      
      const mapping = mappings[itemName]
      const category = mapping ? mapping.category : 'Common Item'
      const heroId = mapping && mapping.heroId ? mapping.heroId : null
      const rankCategory = mapping && mapping.heroId ? 'High Rank' : null
      
      // Получаем историю цен
      let historyData = null
      if (historySheet && historyNames.length > 0) {
        const historyRowIndex = historyNames.findIndex(r => String(r[0] || '').trim() === itemName)
        if (historyRowIndex >= 0) {
          historyData = history_getPriceHistoryForItem_(historySheet, historyRowIndex + DATA_START_ROW)
        } else {
          console.warn(`Invest: предмет "${itemName}" не найден в History`)
        }
      }
      
      // Рассчитываем метрики с обработкой ошибок
      try {
        liquidityScores.push([analytics_calculateLiquidityScore(itemData)])
        demandRatios.push([analytics_calculateDemandRatio(itemData)])
        priceMomenta.push([analytics_calculatePriceMomentum(itemData, historyData)])
        salesTrends.push([analytics_calculateSalesTrend(itemData)])
        volatilityIndices.push([analytics_calculateVolatilityIndex(itemData, historyData)])
      } catch (e) {
        console.error(`Invest: ошибка расчета метрик для "${itemName}":`, e)
        liquidityScores.push([null])
        demandRatios.push([null])
        priceMomenta.push([null])
        salesTrends.push([null])
        volatilityIndices.push([null])
      }
      
      // Hero Trend Score (только для Hero Items)
      let heroTrendValue = null
      if (category === 'Hero Item' && heroId && rankCategory) {
        try {
          const latestStats = heroStats_getLatestStats(heroId, rankCategory)
          if (latestStats) {
            const heroStatsObj = {[rankCategory]: latestStats}
            const heroTrendScore = analytics_calculateHeroTrendScore(heroId, rankCategory, heroStatsObj)
            heroTrendValue = analytics_formatScore(heroTrendScore)
          }
        } catch (e) {
          console.error(`Invest: ошибка расчета Hero Trend для "${itemName}":`, e)
        }
      }
      heroTrends.push([heroTrendValue])
      
    } catch (e) {
      console.error(`Invest: ошибка обработки строки ${i + 1} в calculateAllMetrics:`, e)
      // В случае ошибки добавляем null значения для всех метрик
      liquidityScores.push([null])
      demandRatios.push([null])
      priceMomenta.push([null])
      salesTrends.push([null])
      volatilityIndices.push([null])
      heroTrends.push([null])
    }
  }
  
  // Batch-запись Hero Trend (метрики удалены из отображения, но расчеты остаются для Investment Score)
  const count = heroTrends.length
  if (count > 0) {
    sheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.HERO_TREND), count, 1).setValues(heroTrends)
  }
  // Метрики (liquidityScores, demandRatios, priceMomenta, salesTrends, volatilityIndices) 
  // рассчитываются, но не записываются в таблицу - используются только для расчета Investment Score
}

/**
 * Расчет Investment Score для всех позиций в Invest
 */
function invest_updateInvestmentScores() {
  const sheet = getInvestSheet_()
  if (!sheet) return
  
  const lastRow = sheet.getLastRow()
  if (lastRow < DATA_START_ROW) return
  
  const itemNames = sheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.NAME), lastRow - HEADER_ROW, 1).getValues()
  const itemNamesList = itemNames.map(row => String(row[0] || '').trim()).filter(name => name)
  
  // Получаем данные из SteamWebAPI
  const itemsData = invest_updateMetricsFromSteamWebAPI(itemNamesList)
  
  // Получаем маппинги героев
  const mappings = heroMapping_getAllMappings()
  
  // Получаем историю цен из History
  const historySheet = getHistorySheet_()
  
  // Подготовка данных для batch-операций
  const investmentScores = []
  const riskLevels = []
  const historyNames = historySheet ? historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), historySheet.getLastRow() - HEADER_ROW, 1).getValues() : []
  
  // Рассчитываем Investment Score и Risk Level для всех строк
  const startedAt = Date.now()
  const MAX_EXECUTION_TIME_MS = 300000 // 5 минут
  
  for (let i = 0; i < itemNames.length; i++) {
    // Проверка таймаута
    if (Date.now() - startedAt > MAX_EXECUTION_TIME_MS) {
      console.warn(`Invest: превышено время выполнения updateInvestmentScores (${MAX_EXECUTION_TIME_MS}ms), прервано на строке ${i + 1}`)
      break
    }
    
    try {
      const itemName = String(itemNames[i][0] || '').trim()
      if (!itemName) {
        investmentScores.push([null])
        riskLevels.push([null])
        continue
      }
      
      const itemData = itemsData[itemName]
      if (!itemData) {
        console.warn(`Invest: нет данных из SteamWebAPI для "${itemName}" при расчете Investment Score`)
        investmentScores.push([null])
        riskLevels.push([null])
        continue
      }
      
      // ВАЛИДАЦИЯ: Проверяем валидность цен в itemData
      if (itemData.pricelatest !== undefined && itemData.pricelatest !== null) {
        const priceValidation = validatePrice_(itemData.pricelatest, `${itemName} (pricelatest)`)
        if (!priceValidation.valid) {
          console.warn(`Invest: некорректная цена pricelatest для "${itemName}": ${itemData.pricelatest}, пропускаем расчет Investment Score`)
          investmentScores.push([null])
          riskLevels.push([null])
          continue
        }
      }
      
      const mapping = mappings[itemName]
      const category = mapping ? mapping.category : 'Common Item'
      const heroId = mapping && mapping.heroId ? mapping.heroId : null
      const rankCategory = mapping && mapping.heroId ? 'High Rank' : null
      
      // Получаем историю цен
      let historyData = null
      if (historySheet && historyNames.length > 0) {
        const historyRowIndex = historyNames.findIndex(r => String(r[0] || '').trim() === itemName)
        if (historyRowIndex >= 0) {
          historyData = history_getPriceHistoryForItem_(historySheet, historyRowIndex + DATA_START_ROW)
        } else {
          console.warn(`Invest: предмет "${itemName}" не найден в History при расчете Investment Score`)
        }
      }
      
      // ВАЛИДАЦИЯ: Проверяем наличие критических данных для расчета
      // Если нет itemData.priceLatest и нет historyData - пропускаем расчет
      if ((!itemData.pricelatest || itemData.pricelatest <= 0) && (!historyData || !historyData.prices || historyData.prices.length === 0)) {
        console.warn(`Invest: недостаточно данных для расчета Investment Score для "${itemName}" (нет цен)`)
        investmentScores.push([null])
        riskLevels.push([null])
        continue
      }
      
      // Получаем статистику героя
      let heroStats = null
      if (heroId && rankCategory) {
        try {
          const latestStats = heroStats_getLatestStats(heroId, rankCategory)
          if (latestStats) {
            heroStats = {[rankCategory]: latestStats}
          }
        } catch (e) {
          console.error(`Invest: ошибка получения статистики героя для "${itemName}":`, e)
        }
      }
      
      // Рассчитываем Investment Score
      let investmentScore = 0.5 // Значение по умолчанию
      try {
        investmentScore = analytics_calculateInvestmentScore(
          itemData,
          heroStats,
          historyData,
          category,
          heroId,
          rankCategory
        )
        // Валидация Investment Score
        if (!Number.isFinite(investmentScore) || investmentScore < 0 || investmentScore > 1) {
          console.warn(`Invest: некорректный Investment Score для "${itemName}": ${investmentScore}, используем значение по умолчанию`)
          investmentScore = 0.5
        }
      } catch (e) {
        console.error(`Invest: ошибка расчета Investment Score для "${itemName}":`, e)
        investmentScore = 0.5
      }
      
      investmentScores.push([analytics_formatScore(investmentScore)])
      
      // Рассчитываем Risk Level
      let riskLevel = 'Medium' // Значение по умолчанию
      try {
        const volatilityIndex = analytics_calculateVolatilityIndex(itemData, historyData)
        const demandRatio = analytics_calculateDemandRatio(itemData)
        riskLevel = analytics_calculateRiskLevel(investmentScore, volatilityIndex, demandRatio)
      } catch (e) {
        console.error(`Invest: ошибка расчета Risk Level для "${itemName}":`, e)
      }
      riskLevels.push([riskLevel])
      
    } catch (e) {
      console.error(`Invest: ошибка обработки строки ${i + 1} в updateInvestmentScores:`, e)
      investmentScores.push([null])
      riskLevels.push([null])
    }
  }
  
  // Batch-запись Investment Scores и Risk Levels
  const count = investmentScores.length
  if (count > 0) {
    sheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.INVESTMENT_SCORE), count, 1).setValues(investmentScores)
    sheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.RISK_LEVEL), count, 1).setValues(riskLevels)
  }
}

/**
 * Генерация рекомендации на основе Investment Score
 * @param {number} row - Номер строки
 * @returns {string} Рекомендация
 */
function invest_generateRecommendation(row) {
  const sheet = getInvestSheet_()
  if (!sheet) return '👀 НАБЛЮДАТЬ'
  
  const investmentScoreStr = String(sheet.getRange(row, getColumnIndex(INVEST_COLUMNS.INVESTMENT_SCORE)).getValue() || '').trim()
  if (!investmentScoreStr || investmentScoreStr === '—') return '👀 НАБЛЮДАТЬ'
  
  // Парсим число из формата "🟩 0.93"
  const scoreMatch = investmentScoreStr.match(/(\d+\.?\d*)/)
  if (!scoreMatch) return '👀 НАБЛЮДАТЬ'
  
  const investmentScore = parseFloat(scoreMatch[1])
  
  const heroTrendStr = String(sheet.getRange(row, getColumnIndex(INVEST_COLUMNS.HERO_TREND)).getValue() || '').trim()
  const heroTrend = heroTrendStr !== '—' ? heroTrendStr : '—'
  
  if (investmentScore >= 0.75) {
    return `🟩 КУПИТЬ (Score: ${(investmentScore * 100).toFixed(0)}%, Hero: ${heroTrend})`
  }
  if (investmentScore >= 0.60) {
    return `🟨 ДЕРЖАТЬ (Score: ${(investmentScore * 100).toFixed(0)}%)`
  }
  if (investmentScore < 0.40) {
    return `🟥 ПРОДАТЬ (Score: ${(investmentScore * 100).toFixed(0)}%)`
  }
  return `👀 НАБЛЮДАТЬ (Score: ${(investmentScore * 100).toFixed(0)}%)`
}


