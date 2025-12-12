/**
 * Telegram - Интеграция с Telegram Bot API
 * 
 * Функции для отправки уведомлений и отчетов через Telegram бота
 * Использует Telegram Bot API напрямую через UrlFetchApp
 */

/**
 * Сохраняет конфигурацию Telegram бота
 * @param {string} botToken - Токен бота от @BotFather
 * @param {string} chatId - Chat ID пользователя
 */
function telegram_setConfig(botToken, chatId) {
  if (!botToken || !chatId) {
    throw new Error('Telegram: botToken и chatId обязательны')
  }
  
  const props = PropertiesService.getScriptProperties()
  props.setProperty('telegram_bot_token', botToken)
  props.setProperty('telegram_chat_id', chatId)
  
  console.log('Telegram: конфигурация сохранена')
  return { ok: true }
}

/**
 * Получает конфигурацию Telegram бота
 * @returns {Object} {botToken, chatId} или null если не настроено
 */
function telegram_getConfig() {
  const props = PropertiesService.getScriptProperties()
  const botToken = props.getProperty('telegram_bot_token')
  const chatId = props.getProperty('telegram_chat_id')
  
  if (!botToken || !chatId) {
    return null
  }
  
  return { botToken, chatId }
}

/**
 * Отправляет сообщение в Telegram
 * @param {string} message - Текст сообщения
 * @param {string} parseMode - Режим парсинга ('HTML' или 'Markdown')
 * @param {boolean} disablePreview - Отключить превью ссылок (по умолчанию true)
 * @returns {Object} {ok: boolean, error?: string}
 */
function telegram_sendMessage(message, parseMode = 'HTML', disablePreview = true) {
  const config = telegram_getConfig()
  
  if (!config) {
    console.error('Telegram: конфигурация не настроена')
    return { ok: false, error: 'not_configured' }
  }
  
  if (!message || message.trim().length === 0) {
    console.error('Telegram: пустое сообщение')
    return { ok: false, error: 'empty_message' }
  }
  
  // Ограничение длины сообщения Telegram (4096 символов)
  if (message.length > 4096) {
    message = message.substring(0, 4090) + '...'
  }
  
  const url = `https://api.telegram.org/bot${config.botToken}/sendMessage`
  
  try {
    const payload = {
      chat_id: config.chatId,
      text: message,
      parse_mode: parseMode,
      disable_web_page_preview: disablePreview
    }
    
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    })
    
    const result = JSON.parse(response.getContentText())
    
    if (result.ok) {
      return { ok: true }
    } else {
      const errorCode = result.error_code || 'unknown'
      const errorDescription = result.description || 'unknown'
      console.error(`Telegram API error [${errorCode}]: ${errorDescription}`)
      console.error('Full response:', JSON.stringify(result))
      return { ok: false, error: errorDescription, errorCode: errorCode }
    }
  } catch (e) {
    console.error('Telegram send error:', e)
    return { ok: false, error: 'exception' }
  }
}

/**
 * Тест подключения к Telegram боту
 */
function telegram_testConnection() {
  const config = telegram_getConfig()
  
  if (!config) {
    SpreadsheetApp.getUi().alert('Telegram не настроен!\n\nИспользуйте меню: SteamTable → Настроить Telegram')
    return
  }
  
  const message = '✅ Тест подключения к Telegram боту\n\nЕсли вы видите это сообщение, значит все работает!'
  const result = telegram_sendMessage(message)
  
  if (result.ok) {
    SpreadsheetApp.getUi().alert('✅ Сообщение успешно отправлено в Telegram!')
  } else {
    SpreadsheetApp.getUi().alert(`❌ Ошибка отправки: ${result.error || 'unknown'}`)
  }
}

/**
 * Тест ежедневных уведомлений (для ручного запуска из меню)
 */
function telegram_testDailyNotifications() {
  const config = telegram_getConfig()
  
  if (!config) {
    SpreadsheetApp.getUi().alert('Telegram не настроен!\n\nИспользуйте меню: SteamTable → Telegram → Настроить Telegram')
    return
  }
  
  try {
    telegram_checkDailyPriceTargets()
    SpreadsheetApp.getUi().alert('✅ Ежедневные уведомления отправлены (если есть позиции для уведомлений)')
  } catch (e) {
    console.error('Telegram: ошибка при тесте ежедневных уведомлений:', e)
    SpreadsheetApp.getUi().alert('❌ Ошибка: ' + e.message)
  }
}

/**
 * Настройка Telegram через диалог
 */
function telegram_setup() {
  const ui = SpreadsheetApp.getUi()
  
  // Проверяем текущую конфигурацию
  const currentConfig = telegram_getConfig()
  let promptText = 'Настройка Telegram бота\n\n'
  
  if (currentConfig) {
    promptText += 'Текущий Chat ID: ' + currentConfig.chatId + '\n\n'
  }
  
  promptText += 'Введите Bot Token (от @BotFather):'
  
  const botTokenResponse = ui.prompt(
    'Настройка Telegram',
    promptText,
    ui.ButtonSet.OK_CANCEL
  )
  
  if (botTokenResponse.getSelectedButton() !== ui.Button.OK) {
    return
  }
  
  const botToken = botTokenResponse.getResponseText().trim()
  
  if (!botToken) {
    ui.alert('Ошибка', 'Bot Token не может быть пустым', ui.ButtonSet.OK)
    return
  }
  
  promptText = 'Введите Chat ID:'
  if (currentConfig) {
    promptText += '\n\n(Текущий: ' + currentConfig.chatId + ')'
  }
  
  const chatIdResponse = ui.prompt(
    'Настройка Telegram',
    promptText,
    ui.ButtonSet.OK_CANCEL
  )
  
  if (chatIdResponse.getSelectedButton() !== ui.Button.OK) {
    return
  }
  
  const chatId = chatIdResponse.getResponseText().trim()
  
  if (!chatId) {
    ui.alert('Ошибка', 'Chat ID не может быть пустым', ui.ButtonSet.OK)
    return
  }
  
  try {
    telegram_setConfig(botToken, chatId)
    ui.alert('✅ Конфигурация сохранена!\n\nИспользуйте "Тест Telegram" для проверки.')
  } catch (e) {
    ui.alert('Ошибка', 'Не удалось сохранить конфигурацию: ' + e.message, ui.ButtonSet.OK)
  }
}

/**
 * Проверяет целевые цены и отправляет уведомления
 */
function telegram_checkPriceTargets() {
  const investSheet = getInvestSheet_()
  if (!investSheet) {
    console.log('Telegram: лист Invest не найден')
    return
  }
  
  const lastRow = investSheet.getLastRow()
  if (lastRow <= 1) {
    return // Нет данных
  }
  
  const config = telegram_getConfig()
  if (!config) {
    return // Telegram не настроен, просто выходим
  }
  
  // Читаем данные batch-запросом
  const count = lastRow - 1
  const names = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.NAME), count, 1).getValues()
  const currentPrices = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE), count, 1).getValues()
  const goals = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.GOAL), count, 1).getValues()
  const profits = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.PROFIT), count, 1).getValues()
  const profitPercents = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.PROFIT_AFTER_FEE), count, 1).getValues()
  const recommendations = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.RECOMMENDATION), count, 1).getValues()
  const potentials = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.POTENTIAL), count, 1).getValues()
  const maxPrices = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.MAX_PRICE), count, 1).getValues()
  const investmentScores = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.INVESTMENT_SCORE), count, 1).getValues()
  
  let notificationsSent = 0
  
  // Проверяем каждую позицию
  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue
    
    const currentPrice = Number(currentPrices[i][0]) || 0
    const goal = Number(goals[i][0]) || 0
    const profit = Number(profits[i][0]) || 0
    const profitPercent = Number(profitPercents[i][0]) || 0
    const recommendation = String(recommendations[i][0] || '').trim()
    const potential = Number(potentials[i][0]) || null
    const maxPrice = Number(maxPrices[i][0]) || null
    const investmentScore = investmentScores[i][0] || null
    
    if (goal <= 0 || currentPrice <= 0) continue
    
    // Формируем информацию о потенциале
    // Потенциал P85 хранится как доля (0.5 = 50%) - это реалистичная оценка (85-й перцентиль)
    // Это основной потенциал, на который можно рассчитывать
    // Максимум показываем как дополнительную информацию (теоретический максимум, который может не достичь)
    let potentialInfo = ''
    if (potential !== null && !isNaN(potential)) {
      const potentialPercent = potential * 100 // Преобразуем в проценты для отображения
      const potentialPrice = currentPrice * (1 + potential) // Цена при достижении P85
      potentialInfo = `\nПотенциал роста (P85): <b>+${potentialPercent.toFixed(1)}%</b> (до ~${potentialPrice.toFixed(2)} ₽)`
      
      // Показываем теоретический максимум как дополнительную информацию
      // maxPrice - это максимум из всей истории предмета (не локально зафиксированная цена)
      // ВАЖНО: это теоретический максимум, возврат к нему не гарантирован
      if (maxPrice && maxPrice > currentPrice) {
        const potentialToMax = ((maxPrice - currentPrice) / currentPrice) * 100
        potentialInfo += `\nТеоретический максимум: +${potentialToMax.toFixed(1)}% (${maxPrice.toFixed(2)} ₽)`
      }
    }
    
    // Проверка достижения цели
    if (currentPrice >= goal) {
      const formattedName = telegram_formatItemNameWithScore_(name, investmentScore)
      const message = `🎯 <b>Цель достигнута!</b>\n\n` +
        `Предмет: ${formattedName}\n` +
        `Текущая цена: ${currentPrice.toFixed(2)} ₽\n` +
        `Цель: ${goal.toFixed(2)} ₽\n` +
        `Прибыль: ${profit.toFixed(2)} ₽ (${(profitPercent * 100).toFixed(2)}%)` +
        potentialInfo
      
      const result = telegram_sendMessage(message)
      if (result.ok) {
        notificationsSent++
        Utilities.sleep(LIMITS.TELEGRAM_MESSAGE_DELAY_MS)
      }
    }
    
    // Проверка сильной просадки (50%+)
    if (currentPrice <= goal * 0.5) {
      const dropPercent = ((goal - currentPrice) / goal) * 100
      const formattedName = telegram_formatItemNameWithScore_(name, investmentScore)
      const message = `📉 <b>Сильная просадка!</b>\n\n` +
        `Предмет: ${formattedName}\n` +
        `Текущая цена: ${currentPrice.toFixed(2)} ₽ (макс: ${goal.toFixed(2)} ₽)\n` +
        `Просадка: ${dropPercent.toFixed(2)}%` +
        potentialInfo
      
      const result = telegram_sendMessage(message)
      if (result.ok) {
        notificationsSent++
        Utilities.sleep(500)
      }
    }
  }
  
  if (notificationsSent > 0) {
    console.log(`Telegram: отправлено ${notificationsSent} уведомлений`)
  }
}

/**
 * Ежедневная проверка цен и отправка уведомлений
 * Отправляет три сообщения:
 * 1. Общий отчет о портфеле
 * 2. Позиции, достигшие цели (готовы к продаже)
 * 3. Позиции с сильной просадкой (50%+, сигнал покупки)
 */
function telegram_checkDailyPriceTargets() {
  const now = new Date()
  const hour = now.getHours()
  const minute = now.getMinutes()
  console.log(`Telegram: запуск ежедневных уведомлений в ${hour}:${minute.toString().padStart(2, '0')}`)
  
  const config = telegram_getConfig()
  if (!config) {
    console.log('Telegram: конфигурация не настроена, уведомления не отправлены')
    return // Telegram не настроен, просто выходим
  }
  
  // 1. Отправляем общий отчет о портфеле
  try {
    telegram_sendDailyReport()
    Utilities.sleep(LIMITS.TELEGRAM_REPORT_DELAY_MS)
  } catch (e) {
    console.error('Telegram: критическая ошибка при отправке ежедневного отчета:', e)
    console.error('Stack trace:', e.stack)
    // Продолжаем выполнение, даже если отчет не отправился
  }
  
  const investSheet = getInvestSheet_()
  if (!investSheet) {
    console.log('Telegram: лист Invest не найден')
    return
  }
  
  const lastRow = investSheet.getLastRow()
  if (lastRow <= 1) {
    return // Нет данных
  }
  
  // Читаем данные batch-запросом
  const count = lastRow - 1
  const names = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.NAME), count, 1).getValues()
  const currentPrices = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.CURRENT_PRICE), count, 1).getValues()
  const goals = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.GOAL), count, 1).getValues()
  const profits = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.PROFIT), count, 1).getValues()
  const profitPercents = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.PROFIT_AFTER_FEE), count, 1).getValues()
  const potentials = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.POTENTIAL), count, 1).getValues()
  const maxPrices = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.MAX_PRICE), count, 1).getValues()
  const recommendations = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.RECOMMENDATION), count, 1).getValues()
  const investmentScores = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.INVESTMENT_SCORE), count, 1).getValues()
  
  // Собираем позиции, достигшие цели
  const reachedGoal = []
  // Собираем позиции с сильной просадкой
  const strongDrop = []
  
  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue
    
    const currentPrice = Number(currentPrices[i][0]) || 0
    const goal = Number(goals[i][0]) || 0
    const profit = Number(profits[i][0]) || 0
    const profitPercent = Number(profitPercents[i][0]) || 0
    const potential = Number(potentials[i][0]) || null
    const maxPrice = Number(maxPrices[i][0]) || null
    const recommendation = String(recommendations[i][0] || '').trim()
    const investmentScore = investmentScores[i][0] || null
    
    if (goal <= 0 || currentPrice <= 0) continue
    
    // Проверка достижения цели
    if (currentPrice >= goal) {
      reachedGoal.push({
        name,
        currentPrice,
        goal,
        profit,
        profitPercent,
        potential,
        maxPrice,
        recommendation,
        investmentScore
      })
    }
    
    // Проверка сильной просадки (50%+)
    if (currentPrice <= goal * 0.5) {
      const dropPercent = ((goal - currentPrice) / goal) * 100
      strongDrop.push({
        name,
        currentPrice,
        goal,
        dropPercent,
        potential,
        maxPrice,
        recommendation,
        investmentScore
      })
    }
  }
  
  // Сортируем позиции перед отправкой
  // Достигшие цели - от самой прибыльной (по проценту) к менее прибыльной
  reachedGoal.sort((a, b) => b.profitPercent - a.profitPercent)
  
  // Просевшие позиции - от самых просевших (по проценту просадки) к менее просевшим
  strongDrop.sort((a, b) => b.dropPercent - a.dropPercent)
  
  let messagesSent = 0
  let messagesFailed = 0
  
  // Отправляем первое сообщение: достигшие цели
  if (reachedGoal.length > 0) {
    let message = `🎯 <b>Позиции, достигшие цели</b>\n\n`
    
    reachedGoal.forEach((item, index) => {
      const itemUrl = `https://steamcommunity.com/market/listings/${STEAM_APP_ID}/${encodeURIComponent(item.name)}`
      message += `${index + 1}. <b><a href="${itemUrl}">${item.name}</a></b>\n`
      message += `   Цена: ${item.currentPrice.toFixed(2)} ₽ (цель: ${item.goal.toFixed(2)} ₽)\n`
      message += `   Прибыль: ${item.profit.toFixed(2)} ₽ (${(item.profitPercent * 100).toFixed(2)}%)\n`
      
      // Добавляем информацию о потенциале
      // Потенциал P85 - это реалистичная оценка (85-й перцентиль всех цен) - основной потенциал
      // maxPrice - теоретический максимум из всей истории (дополнительная информация)
      if (item.potential !== null && !isNaN(item.potential)) {
        const potentialPercent = item.potential * 100
        const potentialPrice = item.currentPrice * (1 + item.potential)
        message += `   Потенциал (P85): <b>+${potentialPercent.toFixed(1)}%</b> (до ~${potentialPrice.toFixed(2)} ₽)`
        
        // Теоретический максимум как дополнительная информация
        // ВАЖНО: возврат к максимуму не гарантирован, поэтому акцент на P85
        if (item.maxPrice && item.maxPrice > item.currentPrice) {
          const potentialToMax = ((item.maxPrice - item.currentPrice) / item.currentPrice) * 100
          message += `\n   Теор. максимум: +${potentialToMax.toFixed(1)}% (${item.maxPrice.toFixed(2)} ₽)`
        }
        message += `\n`
      }
      
      if (item.recommendation) {
        message += `   ${item.recommendation}\n`
      }
      message += `\n`
    })
    
    message += `Всего: <b>${reachedGoal.length}</b> позиций`
    
    const result = telegram_sendMessage(message)
    if (result.ok) {
      messagesSent++
    } else {
      messagesFailed++
      console.error(`Telegram: ошибка отправки сообщения о достигших цели: ${result.error}`)
    }
    Utilities.sleep(LIMITS.TELEGRAM_REPORT_DELAY_MS)
  }
  
  // Отправляем второе сообщение: просевшие позиции
  if (strongDrop.length > 0) {
    let message = `📉 <b>Позиции с сильной просадкой</b>\n\n`
    
    strongDrop.forEach((item, index) => {
      const itemUrl = `https://steamcommunity.com/market/listings/${STEAM_APP_ID}/${encodeURIComponent(item.name)}`
      const formattedName = telegram_formatItemNameWithScore_(item.name, item.investmentScore)
      message += `${index + 1}. <a href="${itemUrl}">${formattedName}</a>\n`
      message += `   Цена: ${item.currentPrice.toFixed(2)} ₽ (макс: ${item.goal.toFixed(2)} ₽)\n`
      message += `   Просадка: ${item.dropPercent.toFixed(2)}%\n`
      
      // Добавляем информацию о потенциале
      // Потенциал P85 - это реалистичная оценка (85-й перцентиль всех цен) - основной потенциал
      // maxPrice - теоретический максимум из всей истории (дополнительная информация)
      if (item.potential !== null && !isNaN(item.potential)) {
        const potentialPercent = item.potential * 100
        const potentialPrice = item.currentPrice * (1 + item.potential)
        message += `   Потенциал (P85): <b>+${potentialPercent.toFixed(1)}%</b> (до ~${potentialPrice.toFixed(2)} ₽)`
        
        // Теоретический максимум как дополнительная информация
        // ВАЖНО: возврат к максимуму не гарантирован, поэтому акцент на P85
        if (item.maxPrice && item.maxPrice > item.currentPrice) {
          const potentialToMax = ((item.maxPrice - item.currentPrice) / item.currentPrice) * 100
          message += `\n   Теор. максимум: +${potentialToMax.toFixed(1)}% (${item.maxPrice.toFixed(2)} ₽)`
        }
        message += `\n`
      }
      
      message += `\n`
    })
    
    message += `Всего: <b>${strongDrop.length}</b> позиций`
    
    const result = telegram_sendMessage(message)
    if (result.ok) {
      messagesSent++
    } else {
      messagesFailed++
      console.error(`Telegram: ошибка отправки сообщения о просадке: ${result.error}`)
    }
  }
  
  if (reachedGoal.length === 0 && strongDrop.length === 0) {
    console.log('Telegram: нет позиций для уведомлений')
  } else {
    console.log(`Telegram: отправлено уведомлений - достигли цели: ${reachedGoal.length}, просадка: ${strongDrop.length}`)
    if (messagesFailed > 0) {
      console.error(`Telegram: ошибок при отправке: ${messagesFailed} из ${messagesSent + messagesFailed}`)
    }
  }
}

/**
 * Отправляет ежедневный отчет о портфеле
 * Считает данные напрямую из Invest, не зависит от PortfolioStats
 */
function telegram_sendDailyReport() {
  const config = telegram_getConfig()
  if (!config) {
    console.log('Telegram: конфигурация не настроена, отчет не отправлен')
    return // Telegram не настроен
  }
  
  const investSheet = getInvestSheet_()
  if (!investSheet) {
    console.error('Telegram: лист Invest не найден')
    return
  }
  
  try {
    const lastRow = investSheet.getLastRow()
    if (lastRow <= 1) {
      console.log('Telegram: нет данных в Invest')
      return
    }
    
    // Читаем данные из Invest batch-запросом
    const count = lastRow - 1
    const quantities = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.QUANTITY), count, 1).getValues()
    const totalInvestments = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.TOTAL_INVESTMENT), count, 1).getValues()
    const currentValues = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.CURRENT_VALUE_AFTER_FEE), count, 1).getValues()
    const profits = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.PROFIT), count, 1).getValues()
    const profitPercents = investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.PROFIT_AFTER_FEE), count, 1).getValues()
    
    // Считаем метрики напрямую
    let totalInvestment = 0
    let totalCurrentValue = 0
    let totalProfit = 0
    let totalPositions = 0
    let profitableCount = 0
    let unprofitableCount = 0
    
    for (let i = 0; i < count; i++) {
      const quantity = Number(quantities[i][0]) || 0
      if (quantity <= 0) continue // Пропускаем позиции с нулевым количеством
      
      const investment = Number(totalInvestments[i][0]) || 0
      const currentValue = Number(currentValues[i][0]) || 0
      const profit = Number(profits[i][0]) || 0
      const profitPercent = Number(profitPercents[i][0]) || 0
      
      totalInvestment += investment
      totalCurrentValue += currentValue
      totalProfit += profit
      totalPositions++
      
      // Классификация по прибыльности
      if (profitPercent > 0.01) {
        profitableCount++
      } else if (profitPercent < -0.01) {
        unprofitableCount++
      }
    }
    
    // Рассчитываем общий процент прибыли
    const totalProfitPercent = totalInvestment > 0 
      ? ((totalCurrentValue / totalInvestment) - 1) 
      : 0
    
    // Сообщение 1: Общие метрики портфеля
    const message1 = `📊 <b>Отчет по портфелю</b>\n\n` +
      `Общие вложения: <b>${totalInvestment.toFixed(2)}</b> ₽\n` +
      `Текущая стоимость: <b>${totalCurrentValue.toFixed(2)}</b> ₽\n` +
      `Прибыль/убыток: <b>${totalProfit.toFixed(2)}</b> ₽ (${(totalProfitPercent * 100).toFixed(2)}%)\n\n` +
      `Активных позиций: <b>${totalPositions}</b>\n` +
      `Прибыльных: <b>${profitableCount}</b>\n` +
      `Убыточных: <b>${unprofitableCount}</b>`
    
    let result = telegram_sendMessage(message1)
    if (result.ok) {
      console.log('Telegram: общие метрики портфеля отправлены')
      Utilities.sleep(LIMITS.TELEGRAM_REPORT_DELAY_MS)
    }
    
    // Сообщение 2: Топ-5 возможностей из History (Investment Score >= 0.75, НЕ в портфеле)
    const historySheet = getHistorySheet_()
    const investSheet = getInvestSheet_()
    if (historySheet && investSheet) {
      const investLastRow = investSheet.getLastRow()
      const investNames = investLastRow > 1 
        ? investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.NAME), investLastRow - 1, 1).getValues()
        : []
      const portfolioItems = new Set(investNames.map(row => String(row[0] || '').trim()).filter(name => name))
      
      const historyLastRow = historySheet.getLastRow()
      if (historyLastRow > 1) {
        const count = historyLastRow - 1
        const names = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), count, 1).getValues()
        const investmentScores = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.INVESTMENT_SCORE), count, 1).getValues()
        const currentPrices = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.CURRENT_PRICE), count, 1).getValues()
        
        const opportunities = []
        for (let i = 0; i < count; i++) {
          const name = String(names[i][0] || '').trim()
          if (!name || portfolioItems.has(name)) continue
          
          const investmentScoreStr = String(investmentScores[i][0] || '').trim()
          const investmentScore = telegram_parseScore_(investmentScoreStr)
          
          if (investmentScore && investmentScore >= ANALYTICS_THRESHOLDS.INVESTMENT_SCORE_CRITICAL) {
            const currentPrice = Number(currentPrices[i][0]) || 0
            opportunities.push({ name, investmentScore, currentPrice })
          }
        }
        
        opportunities.sort((a, b) => b.investmentScore - a.investmentScore)
        const top5 = opportunities.slice(0, 5)
        
        if (top5.length > 0) {
          let message2 = `🟢 <b>Топ-5 возможностей для покупки</b>\n\n`
          top5.forEach((opp, index) => {
            const itemUrl = `https://steamcommunity.com/market/listings/${STEAM_APP_ID}/${encodeURIComponent(opp.name)}`
            message2 += `${index + 1}. <b><a href="${itemUrl}">${opp.name}</a></b>\n`
            message2 += `   Investment Score: ${analytics_formatScore(opp.investmentScore)}\n`
            message2 += `   Цена: ${opp.currentPrice.toFixed(2)} ₽\n\n`
          })
          
          result = telegram_sendMessage(message2)
          if (result.ok) {
            console.log('Telegram: топ-5 возможностей отправлены')
            Utilities.sleep(LIMITS.TELEGRAM_REPORT_DELAY_MS)
          }
        }
      }
    }
    
    // Сообщение 3: Топ-5 откупов из Sales (Buyback Score >= 0.75)
    const salesSheet = getSalesSheet_()
    if (salesSheet) {
      const salesLastRow = salesSheet.getLastRow()
      if (salesLastRow > 1) {
        const count = salesLastRow - 1
        const names = salesSheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.NAME), count, 1).getValues()
        const buybackScores = salesSheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.BUYBACK_SCORE), count, 1).getValues()
        const currentPrices = salesSheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.CURRENT_PRICE), count, 1).getValues()
        const priceDrops = salesSheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.PRICE_DROP_PERCENT), count, 1).getValues()
        
        const buybacks = []
        for (let i = 0; i < count; i++) {
          const name = String(names[i][0] || '').trim()
          if (!name) continue
          
          const buybackScoreStr = String(buybackScores[i][0] || '').trim()
          const buybackScore = telegram_parseScore_(buybackScoreStr)
          
          if (buybackScore && buybackScore >= ANALYTICS_THRESHOLDS.BUYBACK_SCORE_CRITICAL) {
            const currentPrice = Number(currentPrices[i][0]) || 0
            const priceDrop = Number(priceDrops[i][0]) || 0
            buybacks.push({ name, buybackScore, currentPrice, priceDrop })
          }
        }
        
        buybacks.sort((a, b) => b.buybackScore - a.buybackScore)
        const top5 = buybacks.slice(0, 5)
        
        if (top5.length > 0) {
          let message3 = `💰 <b>Топ-5 откупов</b>\n\n`
          top5.forEach((item, index) => {
            const itemUrl = `https://steamcommunity.com/market/listings/${STEAM_APP_ID}/${encodeURIComponent(item.name)}`
            message3 += `${index + 1}. <b><a href="${itemUrl}">${item.name}</a></b>\n`
            message3 += `   Buyback Score: ${analytics_formatScore(item.buybackScore)}\n`
            message3 += `   Цена: ${item.currentPrice.toFixed(2)} ₽\n`
            message3 += `   Просадка: ${item.priceDrop.toFixed(2)}%\n\n`
          })
          
          result = telegram_sendMessage(message3)
          if (result.ok) {
            console.log('Telegram: топ-5 откупов отправлены')
          }
        }
      }
    }
    
    if (result && result.ok) {
      console.log('Telegram: ежедневный отчет отправлен успешно')
    } else if (result && !result.ok) {
      console.error('Telegram: ошибка отправки отчета:', result.error)
      throw new Error(`Ошибка отправки отчета: ${result.error}`)
    }
  } catch (e) {
    console.error('Telegram: ошибка при формировании отчета:', e)
    throw e
  }
}

/**
 * Получает или создает лист TelegramNotifications
 * @returns {Sheet} Лист TelegramNotifications
 */
function getOrCreateTelegramNotificationsSheet_() {
  const headers = ['Дата/Время', 'Тип', 'Предмет', 'Событие', 'Данные (JSON)', 'Приоритет']
  const columnWidths = [150, 120, 250, 150, 300, 100]
  return createLogSheet_(SHEET_NAMES.TELEGRAM_NOTIFICATIONS, headers, columnWidths)
}

/**
 * Проверяет, было ли уже отправлено уведомление (cooldown проверка)
 * @param {string} type - Тип уведомления (из TELEGRAM_NOTIFICATION_TYPES)
 * @param {string} itemName - Название предмета
 * @param {string} eventId - Уникальный идентификатор события
 * @param {number} cooldownMs - Период cooldown в миллисекундах
 * @returns {boolean} true если уведомление уже было отправлено и cooldown еще не истек
 */
function telegram_checkNotificationSent_(type, itemName, eventId, cooldownMs) {
  const props = PropertiesService.getScriptProperties()
  const key = `telegram_notif_${type}_${itemName}_${eventId}`
  const lastSentTimestamp = props.getProperty(key)
  
  if (!lastSentTimestamp) {
    return false // Уведомление еще не отправлялось
  }
  
  const lastSent = Number(lastSentTimestamp)
  const now = Date.now()
  const elapsed = now - lastSent
  
  return elapsed < cooldownMs // true если cooldown еще не истек
}

/**
 * Сохраняет информацию об отправленном уведомлении
 * @param {string} type - Тип уведомления
 * @param {string} itemName - Название предмета
 * @param {string} eventId - Уникальный идентификатор события
 * @param {Object} data - Данные уведомления (для сохранения в лист)
 * @param {string} priority - Приоритет (из TELEGRAM_PRIORITY)
 */
function telegram_saveNotification_(type, itemName, eventId, data, priority) {
  const props = PropertiesService.getScriptProperties()
  const key = `telegram_notif_${type}_${itemName}_${eventId}`
  const timestamp = Date.now()
  
  // Сохраняем в PropertiesService для быстрой проверки cooldown
  props.setProperty(key, String(timestamp))
  
  // Сохраняем в лист для истории
  const sheet = getOrCreateTelegramNotificationsSheet_()
  const now = new Date()
  const dateTimeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm:ss')
  const dataJson = JSON.stringify(data)
  
  const row = [dateTimeStr, type, itemName, eventId, dataJson, priority]
  insertLogRowUniversal_(sheet, row, null)
}

/**
 * Очищает старые уведомления из PropertiesService (старше 7 дней)
 * История в листе TelegramNotifications хранится бессрочно
 */
function telegram_cleanupOldNotifications_() {
  const props = PropertiesService.getScriptProperties()
  const allProps = props.getProperties()
  const now = Date.now()
  const sevenDaysMs = 7 * 24 * 60 * 60 * 1000
  let cleanedCount = 0
  
  for (const key in allProps) {
    if (key.startsWith('telegram_notif_')) {
      const timestamp = Number(allProps[key])
      if (!isNaN(timestamp) && (now - timestamp) > sevenDaysMs) {
        props.deleteProperty(key)
        cleanedCount++
      }
    }
  }
  
  if (cleanedCount > 0) {
    console.log(`Telegram: очищено ${cleanedCount} старых уведомлений из PropertiesService`)
    logAutoAction_(SHEET_NAMES.TELEGRAM_NOTIFICATIONS, `Очистка старых уведомлений`, `Очищено ${cleanedCount} записей`)
  }
}

/**
 * Парсит форматированный скор (🟢 0.93) в число
 * @param {string} formattedScore - Отформатированный скор
 * @returns {number|null} Числовое значение скора или null
 */
function telegram_parseScore_(formattedScore) {
  if (!formattedScore || typeof formattedScore !== 'string') return null
  // Убираем эмодзи и пробелы, извлекаем число
  const match = formattedScore.match(/[\d.]+/)
  if (match) {
    const score = Number(match[0])
    return isNaN(score) ? null : score
  }
  return null
}

/**
 * Получает смайлик по значению Investment Score
 * @param {number} score - Score от 0 до 1
 * @returns {string} Смайлик (🟢, 🟡, ⚪, 🔴)
 */
function telegram_getScoreEmoji_(score) {
  if (typeof score !== 'number' || isNaN(score)) return '⚪'
  // Круглые эмодзи: 🟢 (>=0.75), 🟡 (>=0.60), ⚪ (>=0.40), 🔴 (<0.40)
  return score >= 0.75 ? '🟢' : score >= 0.60 ? '🟡' : score >= 0.40 ? '⚪' : '🔴'
}

/**
 * Форматирует название предмета с Investment Score
 * @param {string} name - Название предмета
 * @param {number|string|null} investmentScore - Investment Score (число или отформатированная строка)
 * @returns {string} Отформатированное название с смайликом и скором
 */
function telegram_formatItemNameWithScore_(name, investmentScore) {
  if (!investmentScore && investmentScore !== 0) {
    return `Предмет: <b>${name}</b>`
  }
  
  // Парсим скор, если это строка
  let score = typeof investmentScore === 'number' ? investmentScore : telegram_parseScore_(investmentScore)
  if (score === null) {
    return `Предмет: <b>${name}</b>`
  }
  
  const emoji = telegram_getScoreEmoji_(score)
  return `Предмет: <b>${name}</b> ${emoji} ${score.toFixed(2)}`
}

/**
 * Проверяет Investment Score из History для предметов НЕ в портфеле
 * Отправляет критические уведомления для предметов с Investment Score >= 0.75
 */
function telegram_checkHistoryInvestmentOpportunities_() {
  const config = telegram_getConfig()
  if (!config) return
  
  const historySheet = getHistorySheet_()
  const investSheet = getInvestSheet_()
  if (!historySheet || !investSheet) return
  
  const historyLastRow = historySheet.getLastRow()
  if (historyLastRow <= 1) return
  
  // Получаем список предметов в портфеле
  const investLastRow = investSheet.getLastRow()
  const investNames = investLastRow > 1 
    ? investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.NAME), investLastRow - 1, 1).getValues()
    : []
  const portfolioItems = new Set(investNames.map(row => String(row[0] || '').trim()).filter(name => name))
  
  // Читаем данные из History batch-запросом
  const count = historyLastRow - 1
  const names = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), count, 1).getValues()
  const investmentScores = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.INVESTMENT_SCORE), count, 1).getValues()
  const phases = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.PHASE), count, 1).getValues()
  const potentials = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.POTENTIAL), count, 1).getValues()
  const trends = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.TREND), count, 1).getValues()
  const heroTrends = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.HERO_TREND), count, 1).getValues()
  const currentPrices = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.CURRENT_PRICE), count, 1).getValues()
  
  const opportunities = []
  
  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name || portfolioItems.has(name)) continue // Пропускаем предметы в портфеле
    
    const investmentScoreStr = String(investmentScores[i][0] || '').trim()
    const investmentScore = telegram_parseScore_(investmentScoreStr)
    
    if (!investmentScore || investmentScore < ANALYTICS_THRESHOLDS.INVESTMENT_SCORE_CRITICAL) continue
    
    const phase = String(phases[i][0] || '').trim()
    const potential = Number(potentials[i][0]) || null
    const trend = String(trends[i][0] || '').trim()
    const heroTrendStr = String(heroTrends[i][0] || '').trim()
    const heroTrend = telegram_parseScore_(heroTrendStr)
    const currentPrice = Number(currentPrices[i][0]) || 0
    
    const eventId = `investment_${investmentScore.toFixed(2)}`
    const cooldownMs = TELEGRAM_COOLDOWN_MS.INVESTMENT_SCORE
    
    // Проверяем cooldown
    if (telegram_checkNotificationSent_(TELEGRAM_NOTIFICATION_TYPES.INVESTMENT_SCORE, name, eventId, cooldownMs)) {
      continue
    }
    
    opportunities.push({
      name,
      investmentScore,
      phase,
      potential,
      trend,
      heroTrend,
      currentPrice,
      eventId
    })
  }
  
  // Отправляем критические уведомления (отдельное сообщение для каждого)
  for (const opp of opportunities) {
    const potentialInfo = opp.potential !== null && !isNaN(opp.potential)
      ? `\nПотенциал (P85): <b>+${(opp.potential * 100).toFixed(1)}%</b>`
      : ''
    
    const heroTrendInfo = opp.heroTrend !== null
      ? `\nТренд героя: ${analytics_formatScore(opp.heroTrend)}`
      : ''
    
    const formattedName = telegram_formatItemNameWithScore_(opp.name, opp.investmentScore)
    const message = `🟢 <b>Отличная возможность для покупки!</b>\n\n` +
      `Предмет: ${formattedName}\n` +
      `Фаза: ${opp.phase}\n` +
      `Тренд: ${opp.trend}` +
      potentialInfo +
      heroTrendInfo +
      `\nТекущая цена: ${opp.currentPrice.toFixed(2)} ₽`
    
    const result = telegram_sendMessage(message)
    if (result.ok) {
      telegram_saveNotification_(
        TELEGRAM_NOTIFICATION_TYPES.INVESTMENT_SCORE,
        opp.name,
        opp.eventId,
        {
          investmentScore: opp.investmentScore,
          phase: opp.phase,
          potential: opp.potential,
          trend: opp.trend,
          heroTrend: opp.heroTrend,
          currentPrice: opp.currentPrice
        },
        TELEGRAM_PRIORITY.CRITICAL
      )
      Utilities.sleep(LIMITS.TELEGRAM_MESSAGE_DELAY_MS)
    }
  }
  
  if (opportunities.length > 0) {
    console.log(`Telegram: отправлено ${opportunities.length} уведомлений о возможностях из History`)
  }
}

/**
 * Проверяет Buyback Score из Sales
 * Отправляет критические уведомления для предметов с Buyback Score >= 0.75
 */
function telegram_checkSalesBuybackOpportunities_() {
  const config = telegram_getConfig()
  if (!config) return
  
  const salesSheet = getSalesSheet_()
  if (!salesSheet) return
  
  const lastRow = salesSheet.getLastRow()
  if (lastRow <= 1) return
  
  // Читаем данные batch-запросом
  const count = lastRow - 1
  const names = salesSheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.NAME), count, 1).getValues()
  const buybackScores = salesSheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.BUYBACK_SCORE), count, 1).getValues()
  const riskLevels = salesSheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.RISK_LEVEL), count, 1).getValues()
  const heroTrends = salesSheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.HERO_TREND), count, 1).getValues()
  const priceDrops = salesSheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.PRICE_DROP_PERCENT), count, 1).getValues()
  const sellPrices = salesSheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.SELL_PRICE), count, 1).getValues()
  const currentPrices = salesSheet.getRange(DATA_START_ROW, getColumnIndex(SALES_COLUMNS.CURRENT_PRICE), count, 1).getValues()
  
  const opportunities = []
  
  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue
    
    const buybackScoreStr = String(buybackScores[i][0] || '').trim()
    const buybackScore = telegram_parseScore_(buybackScoreStr)
    
    if (!buybackScore || buybackScore < ANALYTICS_THRESHOLDS.BUYBACK_SCORE_CRITICAL) continue
    
    const riskLevel = String(riskLevels[i][0] || '').trim()
    if (riskLevel !== 'Низкий' && riskLevel !== 'Средний') continue
    
    const heroTrendStr = String(heroTrends[i][0] || '').trim()
    const heroTrend = telegram_parseScore_(heroTrendStr)
    const priceDropPercent = Number(priceDrops[i][0]) || 0
    const sellPrice = Number(sellPrices[i][0]) || 0
    const currentPrice = Number(currentPrices[i][0]) || 0
    
    const eventId = `buyback_${buybackScore.toFixed(2)}`
    const cooldownMs = TELEGRAM_COOLDOWN_MS.BUYBACK_SCORE
    
    // Проверяем cooldown
    if (telegram_checkNotificationSent_(TELEGRAM_NOTIFICATION_TYPES.BUYBACK_SCORE, name, eventId, cooldownMs)) {
      continue
    }
    
    opportunities.push({
      name,
      buybackScore,
      riskLevel,
      heroTrend,
      priceDropPercent,
      sellPrice,
      currentPrice,
      eventId
    })
  }
  
  // Отправляем критические уведомления
  for (const opp of opportunities) {
    const heroTrendInfo = opp.heroTrend !== null
      ? `\nТренд героя: ${analytics_formatScore(opp.heroTrend)}`
      : ''
    
    const formattedName = telegram_formatItemNameWithScore_(opp.name, opp.buybackScore)
    const message = `💰 <b>Отличный момент для откупа!</b>\n\n` +
      `Предмет: ${formattedName}\n` +
      `Уровень риска: ${opp.riskLevel}\n` +
      `Просадка: ${opp.priceDropPercent.toFixed(2)}%\n` +
      `Цена продажи: ${opp.sellPrice.toFixed(2)} ₽ (макс: ${opp.sellPrice.toFixed(2)} ₽)\n` +
      `Текущая цена: ${opp.currentPrice.toFixed(2)} ₽` +
      heroTrendInfo
    
    const result = telegram_sendMessage(message)
    if (result.ok) {
      telegram_saveNotification_(
        TELEGRAM_NOTIFICATION_TYPES.BUYBACK_SCORE,
        opp.name,
        opp.eventId,
        {
          buybackScore: opp.buybackScore,
          riskLevel: opp.riskLevel,
          heroTrend: opp.heroTrend,
          priceDropPercent: opp.priceDropPercent,
          sellPrice: opp.sellPrice,
          currentPrice: opp.currentPrice
        },
        TELEGRAM_PRIORITY.CRITICAL
      )
      Utilities.sleep(LIMITS.TELEGRAM_MESSAGE_DELAY_MS)
    }
  }
  
  if (opportunities.length > 0) {
    console.log(`Telegram: отправлено ${opportunities.length} уведомлений об откупах из Sales`)
  }
}

/**
 * Проверяет изменения Hero Trend Score за 24 часа
 * Отправляет важные уведомления для предметов с изменением > 15% за 24ч
 * Группирует топ-5 в одном сообщении
 */
function telegram_checkHeroTrendChanges_() {
  const config = telegram_getConfig()
  if (!config) return
  
  const historySheet = getHistorySheet_()
  const investSheet = getInvestSheet_()
  if (!historySheet) return
  
  const historyLastRow = historySheet.getLastRow()
  if (historyLastRow <= 1) return
  
  // Получаем список предметов в портфеле
  const investLastRow = investSheet ? investSheet.getLastRow() : 0
  const investNames = investLastRow > 1 
    ? investSheet.getRange(DATA_START_ROW, getColumnIndex(INVEST_COLUMNS.NAME), investLastRow - 1, 1).getValues()
    : []
  const portfolioItems = new Set(investNames.map(row => String(row[0] || '').trim()).filter(name => name))
  
  // Читаем данные из History batch-запросом
  const count = historyLastRow - 1
  const names = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), count, 1).getValues()
  const heroTrends = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.HERO_TREND), count, 1).getValues()
  const currentPrices = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.CURRENT_PRICE), count, 1).getValues()
  
  const props = PropertiesService.getScriptProperties()
  const changes = []
  
  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue
    
    const heroTrendStr = String(heroTrends[i][0] || '').trim()
    const currentHeroTrend = telegram_parseScore_(heroTrendStr)
    
    if (!currentHeroTrend || currentHeroTrend === null) continue
    
    // Получаем предыдущее значение из PropertiesService
    const key = `hero_trend_${name}`
    const previousValueStr = props.getProperty(key)
    const previousValue = previousValueStr ? Number(previousValueStr) : null
    
    if (previousValue === null || isNaN(previousValue)) {
      // Сохраняем текущее значение для следующей проверки
      props.setProperty(key, String(currentHeroTrend))
      continue
    }
    
    // Вычисляем изменение в процентах
    const changePercent = ((currentHeroTrend - previousValue) / previousValue) * 100
    const absChange = Math.abs(changePercent)
    
    // Проверяем порог (15% за 24ч)
    if (absChange <= ANALYTICS_THRESHOLDS.HERO_CHANGE_24H * 100) {
      // Обновляем сохраненное значение
      props.setProperty(key, String(currentHeroTrend))
      continue
    }
    
    const eventId = `hero_trend_${changePercent > 0 ? 'up' : 'down'}_${absChange.toFixed(1)}`
    const cooldownMs = TELEGRAM_COOLDOWN_MS.HERO_TREND
    
    // Проверяем cooldown
    if (telegram_checkNotificationSent_(TELEGRAM_NOTIFICATION_TYPES.HERO_TREND, name, eventId, cooldownMs)) {
      // Обновляем сохраненное значение даже если cooldown активен
      props.setProperty(key, String(currentHeroTrend))
      continue
    }
    
    const currentPrice = Number(currentPrices[i][0]) || 0
    const inPortfolio = portfolioItems.has(name)
    
    changes.push({
      name,
      currentHeroTrend,
      previousValue,
      changePercent,
      currentPrice,
      inPortfolio,
      eventId
    })
    
    // Сохраняем новое значение
    props.setProperty(key, String(currentHeroTrend))
  }
  
  // Сортируем по абсолютному изменению (по убыванию)
  changes.sort((a, b) => Math.abs(b.changePercent) - Math.abs(a.changePercent))
  
  // Берем топ-5 и группируем в одно сообщение
  const top5 = changes.slice(0, 5)
  
  if (top5.length > 0) {
    let message = `📈 <b>Изменения Hero Trend Score (>15% за 24ч)</b>\n\n`
    
    top5.forEach((item, index) => {
      const itemUrl = `https://steamcommunity.com/market/listings/${STEAM_APP_ID}/${encodeURIComponent(item.name)}`
      const changeEmoji = item.changePercent > 0 ? '📈' : '📉'
      const portfolioMark = item.inPortfolio ? ' (в портфеле)' : ''
      
      message += `${index + 1}. <b><a href="${itemUrl}">${item.name}</a></b>${portfolioMark}\n`
      message += `   ${changeEmoji} ${item.changePercent > 0 ? '+' : ''}${item.changePercent.toFixed(1)}%\n`
      message += `   Текущий: ${analytics_formatScore(item.currentHeroTrend)}\n`
      message += `   Предыдущий: ${analytics_formatScore(item.previousValue)}\n`
      if (item.currentPrice > 0) {
        message += `   Цена: ${item.currentPrice.toFixed(2)} ₽\n`
      }
      message += `\n`
    })
    
    message += `Всего: <b>${changes.length}</b> предметов с изменениями`
    
    const result = telegram_sendMessage(message)
    if (result.ok) {
      // Сохраняем уведомления для каждого предмета
      top5.forEach(item => {
        telegram_saveNotification_(
          TELEGRAM_NOTIFICATION_TYPES.HERO_TREND,
          item.name,
          item.eventId,
          {
            currentHeroTrend: item.currentHeroTrend,
            previousValue: item.previousValue,
            changePercent: item.changePercent,
            currentPrice: item.currentPrice,
            inPortfolio: item.inPortfolio
          },
          TELEGRAM_PRIORITY.IMPORTANT
        )
      })
      console.log(`Telegram: отправлено уведомление об изменениях Hero Trend Score для ${top5.length} предметов`)
    }
  }
}

/**
 * Проверяет резкие изменения цены (>20% за 24ч)
 * Отправляет критические уведомления немедленно
 */
function telegram_checkPriceChanges_() {
  const config = telegram_getConfig()
  if (!config) return
  
  const historySheet = getHistorySheet_()
  if (!historySheet) return
  
  const lastRow = historySheet.getLastRow()
  if (lastRow <= 1) return
  
  // Получаем первую колонку с датами
  const firstDateCol = getHistoryFirstDateCol_(historySheet)
  if (!firstDateCol) return
  
  // Находим колонку с ценой 24 часа назад
  const now = new Date()
  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000)
  
  // Ищем колонку с датой, ближайшей к 24 часам назад
  const lastCol = historySheet.getLastColumn()
  let price24hAgoCol = null
  let price24hAgoDate = null
  
  for (let col = firstDateCol; col <= lastCol; col++) {
    const header = historySheet.getRange(HEADER_ROW, col).getDisplayValue()
    // Парсим дату из заголовка (формат: "dd.MM.yy ночь" или "dd.MM.yy день")
    const dateMatch = header.match(/^(\d{2})\.(\d{2})\.(\d{2})/)
    if (dateMatch) {
      const day = parseInt(dateMatch[1])
      const month = parseInt(dateMatch[2]) - 1
      const year = 2000 + parseInt(dateMatch[3])
      const colDate = new Date(year, month, day)
      
      // Проверяем, что дата близка к 24 часам назад (в пределах ±12 часов)
      const diffHours = Math.abs((colDate.getTime() - yesterday.getTime()) / (1000 * 60 * 60))
      if (diffHours <= 12) {
        price24hAgoCol = col
        price24hAgoDate = colDate
        break
      }
    }
  }
  
  if (!price24hAgoCol) {
    // Не нашли подходящую колонку, выходим
    return
  }
  
  // Читаем данные batch-запросом
  const count = lastRow - 1
  const names = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), count, 1).getValues()
  const currentPrices = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.CURRENT_PRICE), count, 1).getValues()
  const prices24hAgo = historySheet.getRange(DATA_START_ROW, price24hAgoCol, count, 1).getValues()
  const investmentScores = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.INVESTMENT_SCORE), count, 1).getValues()
  
  const changes = []
  
  for (let i = 0; i < count; i++) {
    const name = String(names[i][0] || '').trim()
    if (!name) continue
    
    const currentPrice = Number(currentPrices[i][0]) || 0
    const price24hAgo = Number(prices24hAgo[i][0]) || 0
    
    if (currentPrice <= 0 || price24hAgo <= 0) continue
    
    // Вычисляем изменение в процентах
    const changePercent = ((currentPrice - price24hAgo) / price24hAgo) * 100
    const absChange = Math.abs(changePercent)
    
    // Проверяем порог (20% за 24ч)
    if (absChange <= ANALYTICS_THRESHOLDS.PRICE_CHANGE_24H * 100) continue
    
    const eventId = `price_change_${changePercent > 0 ? 'up' : 'down'}_${absChange.toFixed(1)}`
    const cooldownMs = TELEGRAM_COOLDOWN_MS.PRICE_CHANGE
    
    // Проверяем cooldown
    if (telegram_checkNotificationSent_(TELEGRAM_NOTIFICATION_TYPES.PRICE_CHANGE, name, eventId, cooldownMs)) {
      continue
    }
    
    const investmentScoreStr = String(investmentScores[i][0] || '').trim()
    const investmentScore = telegram_parseScore_(investmentScoreStr)
    
    changes.push({
      name,
      currentPrice,
      price24hAgo,
      changePercent,
      investmentScore,
      eventId
    })
  }
  
  // Сортируем по абсолютному изменению (по убыванию)
  changes.sort((a, b) => Math.abs(b.changePercent) - Math.abs(a.changePercent))
  
  // Отправляем критические уведомления (отдельное сообщение для каждого)
  for (const change of changes) {
    const changeEmoji = change.changePercent > 0 ? '📈' : '📉'
    const changeType = change.changePercent > 0 ? 'рост' : 'падение'
    
    const formattedName = telegram_formatItemNameWithScore_(change.name, change.investmentScore)
    let message = `${changeEmoji} <b>Резкое изменение цены!</b>\n\n` +
      `Предмет: ${formattedName}\n` +
      `Текущая цена: ${change.currentPrice.toFixed(2)} ₽\n` +
      `Цена 24ч назад: ${change.price24hAgo.toFixed(2)} ₽\n` +
      `${changeType}: <b>${change.changePercent > 0 ? '+' : ''}${change.changePercent.toFixed(2)}%</b>`
    
    const result = telegram_sendMessage(message)
    if (result.ok) {
      telegram_saveNotification_(
        TELEGRAM_NOTIFICATION_TYPES.PRICE_CHANGE,
        change.name,
        change.eventId,
        {
          currentPrice: change.currentPrice,
          price24hAgo: change.price24hAgo,
          changePercent: change.changePercent,
          investmentScore: change.investmentScore
        },
        TELEGRAM_PRIORITY.CRITICAL
      )
      Utilities.sleep(LIMITS.TELEGRAM_MESSAGE_DELAY_MS)
    }
  }
  
  if (changes.length > 0) {
    console.log(`Telegram: отправлено ${changes.length} уведомлений о резких изменениях цены`)
  }
}

