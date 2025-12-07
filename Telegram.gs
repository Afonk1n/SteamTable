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
 * @returns {Object} {ok: boolean, error?: string}
 */
function telegram_sendMessage(message, parseMode = 'HTML') {
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
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        chat_id: config.chatId,
        text: message,
        parse_mode: parseMode
      }),
      muteHttpExceptions: true
    })
    
    const result = JSON.parse(response.getContentText())
    
    if (result.ok) {
      return { ok: true }
    } else {
      console.error('Telegram API error:', result)
      return { ok: false, error: result.description || 'unknown' }
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
    
    if (goal <= 0 || currentPrice <= 0) continue
    
    // Проверка достижения цели
    if (currentPrice >= goal) {
      const message = `🎯 <b>Цель достигнута!</b>\n\n` +
        `Предмет: <b>${name}</b>\n` +
        `Текущая цена: ${currentPrice.toFixed(2)} ₽\n` +
        `Цель: ${goal.toFixed(2)} ₽\n` +
        `Прибыль: ${profit.toFixed(2)} ₽ (${(profitPercent * 100).toFixed(2)}%)\n\n` +
        `Рекомендация: ${recommendation}`
      
      const result = telegram_sendMessage(message)
      if (result.ok) {
        notificationsSent++
        Utilities.sleep(500) // Пауза между сообщениями
      }
    }
    
    // Проверка сильной просадки (50%+)
    if (currentPrice <= goal * 0.5) {
      const dropPercent = ((goal - currentPrice) / goal) * 100
      const message = `📉 <b>Сильная просадка!</b>\n\n` +
        `Предмет: <b>${name}</b>\n` +
        `Текущая цена: ${currentPrice.toFixed(2)} ₽\n` +
        `Цель: ${goal.toFixed(2)} ₽\n` +
        `Просадка: ${dropPercent.toFixed(2)}%\n\n` +
        `Рекомендация: 🟩 КУПИТЬ`
      
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
 * Отправляет ежедневный отчет о портфеле
 */
function telegram_sendDailyReport() {
  const config = telegram_getConfig()
  if (!config) {
    return // Telegram не настроен
  }
  
  const statsSheet = getOrCreatePortfolioStatsSheet_()
  if (!statsSheet) {
    console.log('Telegram: лист PortfolioStats не найден')
    return
  }
  
  try {
    // Читаем данные из PortfolioStats
    const totalInvestment = statsSheet.getRange(4, 2).getValue() || 0
    const totalCurrentValue = statsSheet.getRange(5, 2).getValue() || 0
    const totalProfit = statsSheet.getRange(6, 2).getValue() || 0
    const totalProfitPercent = statsSheet.getRange(6, 3).getValue() || 0
    const totalPositions = statsSheet.getRange(8, 2).getValue() || 0
    
    // Подсчитываем прибыльные/убыточные из данных
    const investSheet = getInvestSheet_()
    let profitableCount = 0
    let unprofitableCount = 0
    
    if (investSheet) {
      const lastRow = investSheet.getLastRow()
      if (lastRow > 1) {
        const profitPercents = investSheet.getRange(
          DATA_START_ROW, 
          getColumnIndex(INVEST_COLUMNS.PROFIT_AFTER_FEE), 
          lastRow - 1, 
          1
        ).getValues()
        
        for (let i = 0; i < profitPercents.length; i++) {
          const percent = Number(profitPercents[i][0]) || 0
          if (percent > 0.01) {
            profitableCount++
          } else if (percent < -0.01) {
            unprofitableCount++
          }
        }
      }
    }
    
    const message = `📊 <b>Отчет по портфелю</b>\n\n` +
      `Общие вложения: <b>${totalInvestment.toFixed(2)}</b> ₽\n` +
      `Текущая стоимость: <b>${totalCurrentValue.toFixed(2)}</b> ₽\n` +
      `Прибыль/убыток: <b>${totalProfit.toFixed(2)}</b> ₽ (${(totalProfitPercent * 100).toFixed(2)}%)\n\n` +
      `Активных позиций: <b>${totalPositions}</b>\n` +
      `Прибыльных: <b>${profitableCount}</b>\n` +
      `Убыточных: <b>${unprofitableCount}</b>`
    
    const result = telegram_sendMessage(message)
    
    if (result.ok) {
      console.log('Telegram: ежедневный отчет отправлен')
    } else {
      console.error('Telegram: ошибка отправки отчета:', result.error)
    }
  } catch (e) {
    console.error('Telegram: ошибка при формировании отчета:', e)
  }
}

