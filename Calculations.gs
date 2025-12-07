/**
 * Calculations - Унифицированная система расчётов
 * 
 * Централизует все расчёты для Invest и Sales с поддержкой batch-операций
 */

/**
 * Оптимизированный batch-расчёт для Invest
 * @param {Sheet} sheet - Лист Invest
 * @param {Array<Array<number>>} currentPrices - Массив текущих цен
 */
function invest_calculateBatch_(sheet, currentPrices) {
  if (!sheet || !currentPrices || currentPrices.length === 0) return
  
  const count = currentPrices.length
  
  const quantityCol = getColumnIndex(INVEST_COLUMNS.QUANTITY)
  const buyPriceCol = getColumnIndex(INVEST_COLUMNS.BUY_PRICE)
  const totalInvCol = getColumnIndex(INVEST_COLUMNS.TOTAL_INVESTMENT)
  const currentValAfterFeeCol = getColumnIndex(INVEST_COLUMNS.CURRENT_VALUE_AFTER_FEE)
  const profitCol = getColumnIndex(INVEST_COLUMNS.PROFIT)
  const profitAfterFeeCol = getColumnIndex(INVEST_COLUMNS.PROFIT_AFTER_FEE)
  
  const quantities = sheet.getRange(DATA_START_ROW, quantityCol, count, 1).getValues()
  const buyPrices = sheet.getRange(DATA_START_ROW, buyPriceCol, count, 1).getValues()
  
  const totalInvestments = []
  const currentValuesAfterFee = []
  const profits = []
  const profitAfterFees = []
  
  for (let i = 0; i < count; i++) {
    const currentPrice = Number(currentPrices[i][0])
    const quantity = Number(quantities[i][0])
    const buyPrice = Number(buyPrices[i][0])
    
    if (Number.isFinite(currentPrice) && currentPrice > 0 &&
        Number.isFinite(quantity) && quantity > 0 &&
        Number.isFinite(buyPrice) && buyPrice > 0) {
      const totalInvestment = quantity * buyPrice
      const currentValue = quantity * currentPrice
      const currentValueAfterFee = currentValue * STEAM_FEE
      const profit = currentValueAfterFee - totalInvestment
      const profitAfterFee = (currentValueAfterFee / totalInvestment) - 1
      
      totalInvestments.push([totalInvestment])
      currentValuesAfterFee.push([currentValueAfterFee])
      profits.push([profit])
      profitAfterFees.push([profitAfterFee])
    } else {
      totalInvestments.push([null])
      currentValuesAfterFee.push([null])
      profits.push([null])
      profitAfterFees.push([null])
    }
  }
  
  sheet.getRange(DATA_START_ROW, totalInvCol, count, 1).setValues(totalInvestments)
  sheet.getRange(DATA_START_ROW, currentValAfterFeeCol, count, 1).setValues(currentValuesAfterFee)
  sheet.getRange(DATA_START_ROW, profitCol, count, 1).setValues(profits)
  sheet.getRange(DATA_START_ROW, profitAfterFeeCol, count, 1).setValues(profitAfterFees)
  
  invest_formatGoalColumn_(DATA_START_ROW, DATA_START_ROW + count - 1)
}

/**
 * Расчёт для одной строки Invest (для обратной совместимости)
 * @param {Sheet} sheet - Лист Invest
 * @param {number} row - Номер строки
 * @param {number} currentPrice - Текущая цена
 */
function invest_calculateSingle_(sheet, row, currentPrice) {
  if (!sheet) return
  
  const quantity = Number(sheet.getRange(row, getColumnIndex(INVEST_COLUMNS.QUANTITY)).getValue())
  const buyPrice = Number(sheet.getRange(row, getColumnIndex(INVEST_COLUMNS.BUY_PRICE)).getValue())
  
  sheet.getRange(row, getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE)).setValue(currentPrice)
  
  if (Number.isFinite(currentPrice) && currentPrice > 0 &&
      Number.isFinite(quantity) && quantity > 0 &&
      Number.isFinite(buyPrice) && buyPrice > 0) {
    const totalInvestment = quantity * buyPrice
    const currentValue = quantity * currentPrice
    const currentValueAfterFee = currentValue * STEAM_FEE
    const profit = currentValueAfterFee - totalInvestment
    const profitAfterFee = (currentValueAfterFee / totalInvestment) - 1
    
    sheet.getRange(row, getColumnIndex(INVEST_COLUMNS.TOTAL_INVESTMENT)).setValue(totalInvestment)
    sheet.getRange(row, getColumnIndex(INVEST_COLUMNS.CURRENT_VALUE_AFTER_FEE)).setValue(currentValueAfterFee)
    sheet.getRange(row, getColumnIndex(INVEST_COLUMNS.PROFIT)).setValue(profit)
    sheet.getRange(row, getColumnIndex(INVEST_COLUMNS.PROFIT_AFTER_FEE)).setValue(profitAfterFee)
  }
  
  invest_formatGoalColumn_(row, row)
}

/**
 * Оптимизированный batch-расчёт для Sales
 * @param {Sheet} sheet - Лист Sales
 * @param {Array<Array<number>>} currentPrices - Массив текущих цен
 * @param {Array<Array<number>>} sellPrices - Массив цен продажи
 */
function sales_calculateBatch_(sheet, currentPrices, sellPrices) {
  if (!sheet || !currentPrices || !sellPrices || currentPrices.length === 0) return
  
  const count = currentPrices.length
  const priceDrops = []
  
  for (let i = 0; i < count; i++) {
    const currentPrice = Number(currentPrices[i][0])
    const sellPrice = Number(sellPrices[i][0])
    
    if (Number.isFinite(sellPrice) && sellPrice > 0 &&
        Number.isFinite(currentPrice) && currentPrice > 0) {
      priceDrops.push([(sellPrice - currentPrice) / sellPrice])
    } else {
      priceDrops.push([null])
    }
  }
  
  const dropCol = getColumnIndex(SALES_COLUMNS.PRICE_DROP)
  sheet.getRange(DATA_START_ROW, dropCol, count, 1).setValues(priceDrops)
}

/**
 * Расчёт для одной строки Sales (для обратной совместимости)
 * @param {Sheet} sheet - Лист Sales
 * @param {number} row - Номер строки
 * @param {number} currentPrice - Текущая цена
 */
function sales_calculateSingle_(sheet, row, currentPrice) {
  if (!sheet) return
  
  const sellPrice = Number(sheet.getRange(row, getColumnIndex(SALES_COLUMNS.SELL_PRICE)).getValue())
  
  sheet.getRange(row, getColumnIndex(SALES_COLUMNS.CURRENT_PRICE)).setValue(currentPrice)
  
  let priceDrop = null
  if (Number.isFinite(sellPrice) && sellPrice > 0 &&
      Number.isFinite(currentPrice) && currentPrice > 0) {
    priceDrop = (sellPrice - currentPrice) / sellPrice
  }
  
  sheet.getRange(row, getColumnIndex(SALES_COLUMNS.PRICE_DROP)).setValue(priceDrop)
}

