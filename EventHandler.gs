/**
 * EventHandler - Обработка событий Google Sheets
 * 
 * Объединяет обработку покупок/продаж и других событий
 */

// Обработчик изменения листа для автоматической обработки чекбоксов
// Отслеживает чекбоксы "Купить?" (History, колонка E) и "Продать?" (Invest, колонка V)
function onEdit(e) {
  if (!e?.range) return
  
  const sheet = e.source.getActiveSheet()
  if (!sheet) return
  
  const sheetName = sheet.getName()
  const row = e.range.getRow()
  const col = e.range.getColumn()
  
  // Ранний выход если изменение не в нужных листах или строках
  if (row <= HEADER_ROW) return
  
  const buyCheckboxCol = getColumnIndex(HISTORY_COLUMNS.BUY_CHECKBOX)
  const sellCheckboxCol = getColumnIndex(INVEST_COLUMNS.SELL_CHECKBOX)
  
  // Обработка чекбокса "Купить?" в History
  if (sheetName === SHEET_NAMES.HISTORY && col === buyCheckboxCol) {
    const isChecked = e.range.getValue() === true
    if (isChecked) {
      try {
        handleBuyFromHistory_(sheet, row)
        Utilities.sleep(100)
        sheet.getRange(row, col).setValue(false)
      } catch (error) {
        console.error('EventHandler: ошибка при обработке покупки:', error)
        sheet.getRange(row, col).setValue(false)
        SpreadsheetApp.getUi().alert('Ошибка при обработке покупки: ' + error.toString())
      }
    }
    return
  }
  
  // Обработка чекбокса "Продать?" в Invest
  if (sheetName === SHEET_NAMES.INVEST && col === sellCheckboxCol) {
    const isChecked = e.range.getValue() === true
    if (isChecked) {
      try {
        handleSellFromInvest_(sheet, row)
        Utilities.sleep(100)
        sheet.getRange(row, col).setValue(false)
      } catch (error) {
        console.error('EventHandler: ошибка при обработке продажи:', error)
        sheet.getRange(row, col).setValue(false)
        SpreadsheetApp.getUi().alert('Ошибка при обработке продажи: ' + error.toString())
      }
    }
    return
  }
}

/**
 * Обработка покупки из History через чекбокс
 * Вызывается автоматически при установке чекбокса "Купить?" в true
 * @param {Sheet} historySheet - Лист History
 * @param {number} row - Номер строки
 */
function handleBuyFromHistory_(historySheet, row) {
  const name = historySheet.getRange(row, getColumnIndex(HISTORY_COLUMNS.NAME)).getValue()
  if (!name) return
  
  const ui = SpreadsheetApp.getUi()
  
  // Запрашиваем количество
  const qtyResp = ui.prompt('Покупка', `Введите количество для «${name}»:`, ui.ButtonSet.OK_CANCEL)
  if (qtyResp.getSelectedButton() !== ui.Button.OK) {
    SpreadsheetApp.getUi().alert('Покупка отменена')
    return
  }
  
  const qtyParsed = parseRuNumber_(qtyResp.getResponseText())
  if (!qtyParsed.ok || qtyParsed.value <= 0 || !Number.isInteger(qtyParsed.value)) {
    SpreadsheetApp.getUi().alert('Некорректное количество (должно быть целым числом)')
    return
  }
  
  // Запрашиваем цену покупки
  const priceResp = ui.prompt('Покупка', 'Введите цену покупки за 1 шт (РУБ):', ui.ButtonSet.OK_CANCEL)
  if (priceResp.getSelectedButton() !== ui.Button.OK) {
    SpreadsheetApp.getUi().alert('Покупка отменена')
    return
  }
  
  const priceParsed = parseRuNumber_(priceResp.getResponseText())
  if (!priceParsed.ok || priceParsed.value <= 0) {
    SpreadsheetApp.getUi().alert('Некорректная цена')
    return
  }
  
  const buyPrice = priceParsed.value
  const quantity = qtyParsed.value
  
  // Выполняем покупку
  const result = invest_addOrUpdatePosition_(name, quantity, buyPrice)
  
  // Логирование
  try {
    logOperation_('BUY', name, quantity, buyPrice, quantity * buyPrice, 'History')
  } catch (e) {
    console.error('EventHandler: ошибка при логировании покупки:', e)
  }
  
  // Если создана новая строка - синхронизируем всю аналитику
  if (result && result.created && result.row) {
    const investSheet = getInvestSheet_()
    if (investSheet) {
      // Получаем текущую цену из History для расчётов
      const period = getCurrentPricePeriod()
      const priceResult = getHistoryPriceForPeriod_(historySheet, name, period)
      let currentPrice = 0
      if (priceResult && priceResult.found && priceResult.price > 0) {
        currentPrice = priceResult.price
      }
      
      // Обновляем текущую цену и расчёты
      const currentPriceCol = getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE)
      investSheet.getRange(result.row, currentPriceCol).setValue(currentPrice)
      invest_calculateSingle_(investSheet, result.row, currentPrice)
      
      // Синхронизируем всю аналитику из History для новой строки
      invest_syncMinMaxFromHistory(false)
      invest_syncTrendDaysFromHistory(false)
      invest_syncExtendedAnalyticsFromHistory(false)
    }
  } else if (result && !result.created && result.row) {
    // Обновление существующей позиции - пересчитываем только расчёты
    const investSheet = getInvestSheet_()
    if (investSheet) {
      const currentPriceCol = getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE)
      const currentPrice = Number(investSheet.getRange(result.row, currentPriceCol).getValue()) || 0
      if (currentPrice > 0) {
        invest_calculateSingle_(investSheet, result.row, currentPrice)
      }
    }
  }
  
  SpreadsheetApp.getUi().alert(`Покупка выполнена: ${quantity} шт × ${buyPrice.toFixed(2)} ₽ = ${(quantity * buyPrice).toFixed(2)} ₽`)
}

/**
 * Обработка продажи из Invest через чекбокс
 * Вызывается автоматически при установке чекбокса "Продать?" в true
 * @param {Sheet} investSheet - Лист Invest
 * @param {number} row - Номер строки
 */
function handleSellFromInvest_(investSheet, row) {
  const name = investSheet.getRange(row, getColumnIndex(INVEST_COLUMNS.NAME)).getValue()
  if (!name) {
    SpreadsheetApp.getUi().alert('В строке нет названия предмета')
    return
  }
  
  const qtyAvailable = Number(investSheet.getRange(row, getColumnIndex(INVEST_COLUMNS.QUANTITY)).getValue())
  if (!Number.isFinite(qtyAvailable) || qtyAvailable <= 0) {
    SpreadsheetApp.getUi().alert('Неверное количество в строке')
    return
  }
  
  const ui = SpreadsheetApp.getUi()
  
  // Запрашиваем количество
  const qtyResp = ui.prompt('Продажа', `Введите количество (1…${qtyAvailable}) для «${name}»:`, ui.ButtonSet.OK_CANCEL)
  if (qtyResp.getSelectedButton() !== ui.Button.OK) {
    SpreadsheetApp.getUi().alert('Продажа отменена')
    return
  }
  
  const qtyParsed = parseRuNumber_(qtyResp.getResponseText())
  if (!qtyParsed.ok || qtyParsed.value < 1 || qtyParsed.value > qtyAvailable || !Number.isInteger(qtyParsed.value)) {
    SpreadsheetApp.getUi().alert('Некорректное количество')
    return
  }
  
  // Запрашиваем цену продажи
  const priceResp = ui.prompt('Продажа', 'Введите цену продажи за 1 шт (РУБ):', ui.ButtonSet.OK_CANCEL)
  if (priceResp.getSelectedButton() !== ui.Button.OK) {
    SpreadsheetApp.getUi().alert('Продажа отменена')
    return
  }
  
  const priceParsed = parseRuNumber_(priceResp.getResponseText())
  if (!priceParsed.ok || priceParsed.value <= 0) {
    SpreadsheetApp.getUi().alert('Некорректная цена')
    return
  }
  
  // Выполняем продажу
  invest_applySale(row, qtyParsed.value, priceParsed.value)
  
  SpreadsheetApp.getUi().alert(`Продажа выполнена: ${qtyParsed.value} шт × ${priceParsed.value.toFixed(2)} ₽ = ${(qtyParsed.value * priceParsed.value).toFixed(2)} ₽`)
}

