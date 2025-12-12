/**
 * Menu - Меню и триггеры
 * 
 * Управляет созданием меню, триггеров и обработкой событий
 */

// Настройка всех триггеров
// Архитектура:
// 1. Два фиксированных триггера (00:10 и 12:00) для создания колонок периодов и начала сбора
// 2. Один периодический триггер (каждые 10 минут) для продолжения сбора до завершения периода
// 3. Три триггера для обновления статистики героев (08:00, 14:00, 20:00)
// 4. Один триггер для архивации данных HeroStats (воскресенье 03:00)
// 5. Один триггер для синхронизации HeroMapping (ежедневно 04:00)
// Такая архитектура обеспечивает надежность - нет риска что временные триггеры не создадутся или удалятся преждевременно
function setupAllTriggers() {
  removeAllTriggers()
  
  // Создание колонки и запуск сбора для периода "ночь" (00:10)
  // Примечание: atHour не поддерживает минуты, поэтому используем ближайший час (0:00)
  // Проверка точного времени (00:10) выполняется внутри history_ensurePeriodColumn
  ScriptApp.newTrigger('history_createPeriodAndUpdate_morning')
    .timeBased()
    .atHour(UPDATE_INTERVALS.MORNING_HOUR)
    .everyDays(1)
    .create()
  
  // Создание колонки и запуск сбора для периода "день" (12:00)
  ScriptApp.newTrigger('history_createPeriodAndUpdate_evening')
    .timeBased()
    .atHour(UPDATE_INTERVALS.EVENING_HOUR)
    .everyDays(1)
    .create()
  
  // Продолжение сбора цен каждые 10 минут (пока период не завершен)
  ScriptApp.newTrigger('unified_priceUpdate')
    .timeBased()
    .everyMinutes(UPDATE_INTERVALS.PRICES_MINUTES)
    .create()
  
  // Ежедневная проверка цен для Telegram уведомлений (13:00)
  ScriptApp.newTrigger('telegram_checkDailyPriceTargets')
    .timeBased()
    .atHour(13) // 13:00 (час дня)
    .everyDays(1)
    .create()
  
  // Автоматическое обновление статистики героев (если включено в конфигурации)
  if (HERO_STATS_UPDATE_SCHEDULE.ENABLED && HERO_STATS_UPDATE_SCHEDULE.HOURS) {
    HERO_STATS_UPDATE_SCHEDULE.HOURS.forEach(hour => {
      ScriptApp.newTrigger('autoUpdateHeroStats')
        .timeBased()
        .atHour(hour)
        .everyDays(1)
        .create()
    })
  }
  
  // Автоматическая архивация данных HeroStats (воскресенье 03:00)
  // Используем everyWeeks для запуска каждое воскресенье
  ScriptApp.newTrigger('autoArchiveHeroStats')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(3)
    .create()
  
  // Автоматическая синхронизация HeroMapping (ежедневно 04:00)
  ScriptApp.newTrigger('autoSyncHeroMapping')
    .timeBased()
    .atHour(4) // 04:00
    .everyDays(1)
    .create()
  
  SpreadsheetApp.getUi().alert('Все триггеры настроены')
}

// Удаление всех триггеров
function removeAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers()
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger)
  })
  SpreadsheetApp.getUi().alert('Все триггеры удалены')
}

function history_createPeriodAndUpdate_morning() {
  setPriceCollectionState(PRICE_COLLECTION_PERIODS.MORNING, false)
  history_createPeriodAndUpdate()
}

function history_createPeriodAndUpdate_evening() {
  setPriceCollectionState(PRICE_COLLECTION_PERIODS.EVENING, false)
  history_createPeriodAndUpdate()
}

// Создание меню
function onOpen() {
  const ui = SpreadsheetApp.getUi()
  
  // Главное меню SteamTable
  ui.createMenu('SteamTable')
    .addItem('Включить автообновление', 'setupAllTriggers')
    .addItem('Выключить автообновление', 'removeAllTriggers')
    .addSeparator()
    .addItem('Инициализировать все таблицы', 'initializeAllTables')
    .addSeparator()
    .addItem('Полная настройка (все шаги)', 'performFullSetup')
    .addItem('Шаг 1: Расчет Min/Max из SteamWebAPI', 'setupMinMax')
    .addItem('Шаг 2: Настройка HeroMapping', 'setupHeroMapping')
    .addItem('Шаг 3: Обновление статистики героев', 'setupHeroStats')
    .addItem('Шаг 4: Обновление аналитики и метрик', 'setupAnalytics')
    .addSeparator()
    .addItem('Проверить готовность системы', 'checkSystemReadiness')
    .addItem('Проверить состояние автоматизации', 'checkAutomationStatus')
    .addSeparator()
    .addItem('Обновить цены History (ручное)', 'history_updateAllPrices')
    .addItem('Рассчитать Min/Max для всех предметов', 'priceHistory_calculateMinMaxForAllItems')
    .addItem('Обновить Min/Max только у отсутствующих', 'priceHistory_calculateMinMaxForMissingItems')
    .addSeparator()
    // Сокращённые меню
    .addSubMenu(ui.createMenu('Invest')
      .addItem('Форматирование', 'invest_formatTable')
      .addItem('Изображение и ссылки', 'invest_updateImagesAndLinks')
      .addItem('Поиск дублей', 'invest_findDuplicates')
      .addSeparator()
      .addItem('Обновить метрики', 'invest_calculateAllMetrics')
      .addItem('Обновить Investment Scores', 'invest_updateInvestmentScores')
    )
    .addSubMenu(ui.createMenu('Sales')
      .addItem('Форматирование', 'sales_formatTable')
      .addItem('Изображение и ссылки', 'sales_updateImagesAndLinks')
      .addItem('Поиск дублей', 'sales_findDuplicates')
      .addSeparator()
      .addItem('Обновить метрики', 'sales_calculateAllMetrics')
      .addItem('Обновить Buyback Scores', 'sales_updateBuybackScores')
    )
    .addSubMenu(ui.createMenu('History')
      .addItem('Форматирование', 'history_formatTable')
      .addItem('Обновить аналитику', 'history_updateAllAnalytics_')
      .addItem('Изображения и ссылки', 'history_updateImagesAndLinks')
      .addItem('Дубли названий', 'history_findDuplicates')
      .addItem('Создать столбец текущего периода', 'history_ensureTodayColumn')
      .addSeparator()
      .addItem('Нормализовать формат цен в колонках дат', 'history_normalizePriceFormats')
      .addSeparator()
      .addItem('Синхронизировать статистику героев', 'history_syncHeroStats')
      .addItem('Обновить Investment Scores', 'history_updateInvestmentScores')
    )
    .addSubMenu(ui.createMenu('PortfolioStats')
      .addItem('Форматирование', 'portfolioStats_formatTable')
      .addItem('Сохранить строку истории (тест)', 'portfolioStats_saveHistoryManual')
    )
    .addSubMenu(ui.createMenu('HeroStats')
      .addItem('Форматирование', 'heroStats_formatTable')
      .addItem('Обновить статистику (с про-статистикой)', 'heroStats_updateAllStats')
      .addItem('Архивировать старые данные', 'heroStats_archiveOldData')
    )
    .addSubMenu(ui.createMenu('HeroMapping')
      .addItem('Форматирование', 'heroMapping_formatTable')
      .addItem('Автоопределение героев', 'heroMapping_autoDetectFromSteamWebAPI')
      .addItem('Синхронизировать с History', 'heroMapping_syncWithHistory')
      .addSeparator()
      .addItem('Заполнить пустые Hero ID', 'heroMapping_fillMissingHeroIdsMenu')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('Синхронизация')
      .addItem('Синхронизировать цены из History', 'syncPricesFromHistoryToInvestAndSales')
      .addItem('Обновить аналитику Invest/Sales', 'syncAnalyticsForInvestSales_')
      .addSeparator()
      .addItem('Обновить метрики Invest/Sales', 'updateAllMetricsForInvestSales')
      .addItem('Обновить все метрики и скоры', 'updateAllMetricsAndScores_')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('API Settings')
      .addItem('Тест OpenDota API', 'openDota_testConnection')
      .addSeparator()
      .addItem('Тест SteamWebAPI.ru', 'steamWebAPI_testConnection')
    )
    .addSubMenu(ui.createMenu('Telegram')
      .addItem('Настроить Telegram', 'telegram_setup')
      .addItem('Тест Telegram', 'telegram_testConnection')
      .addItem('Тест ежедневных уведомлений', 'telegram_testDailyNotifications')
    )
    .addToUi()
}

// Инициализация всех таблиц (форматирование и настройка)
function initializeAllTables() {
  try {
    invest_formatTable()
    sales_formatTable()
    history_formatTable()
    portfolioStats_formatTable()
    heroStats_formatTable()
    heroMapping_formatTable()
    // Создаем листы логов, если их еще нет
    getOrCreateAutoLogSheet_()
    getOrCreateLogSheet_()
    // Создаем лист TelegramNotifications, если его еще нет
    getOrCreateTelegramNotificationsSheet_()
    console.log('Menu: все таблицы инициализированы и отформатированы')
  } catch (e) {
    console.error('Menu: ошибка инициализации таблиц:', e)
    SpreadsheetApp.getUi().alert('Ошибка инициализации таблиц: ' + e.message)
  }
}

/**
 * Первоначальная настройка новой таблицы
 * Выполняет все необходимые операции для первого запуска последовательно
 * Избегает дублирующихся запросов к API
 */
function performInitialSetup() {
  const ui = SpreadsheetApp.getUi()
  
  // Подтверждение начала настройки
  const startResponse = ui.alert(
    'Первоначальная настройка',
    'Эта функция выполнит все необходимые операции для настройки новой таблицы:\n\n' +
    '1. Заполнение Min/Max цен\n' +
    '2. Синхронизация HeroMapping с History\n' +
    '3. Автоопределение героев\n' +
    '4. Первое обновление статистики героев\n\n' +
    '⚠️ ВАЖНО: Перед запуском убедитесь, что таблицы инициализированы!\n' +
    '(SteamTable → Инициализировать все таблицы)\n\n' +
    'Это может занять несколько минут. Продолжить?',
    ui.ButtonSet.YES_NO
  )
  
  if (startResponse !== ui.Button.YES) {
    return
  }
  
  const results = {
    minMaxCalculated: false,
    heroMappingSynced: false,
    heroesDetected: false,
    statsUpdated: false
  }
  
  try {
    // ШАГ 1: Заполнение Min/Max цен (с выбором режима)
    const minMaxResponse = ui.alert(
      'Шаг 1/4: Заполнение Min/Max цен',
      'Как заполнить Min/Max?\n\n' +
      'ДА - для всех предметов (может занять много времени)\n' +
      'НЕТ - только у отсутствующих (рекомендуется)\n' +
      'ОТМЕНА - пропустить этот шаг',
      ui.ButtonSet.YES_NO_CANCEL
    )
    
    if (minMaxResponse === ui.Button.YES) {
      ui.alert('Начинаем расчет Min/Max для всех предметов...')
      priceHistory_calculateMinMaxForAllItems(false)
      results.minMaxCalculated = true
      console.log('InitialSetup: Min/Max рассчитаны для всех предметов')
    } else if (minMaxResponse === ui.Button.NO) {
      ui.alert('Начинаем расчет Min/Max только для отсутствующих...')
      priceHistory_calculateMinMaxForAllItems(true)
      results.minMaxCalculated = true
      console.log('InitialSetup: Min/Max рассчитаны для отсутствующих')
    } else {
      console.log('InitialSetup: пропущен расчет Min/Max')
    }
    
    // ШАГ 2: Синхронизация HeroMapping с History
    const syncResponse = ui.alert(
      'Шаг 2/4: Синхронизация HeroMapping',
      'Синхронизировать предметы из History в HeroMapping?',
      ui.ButtonSet.YES_NO
    )
    
    if (syncResponse === ui.Button.YES) {
      ui.alert('Начинаем синхронизацию...')
      heroMapping_syncWithHistory()
      results.heroMappingSynced = true
      console.log('InitialSetup: HeroMapping синхронизирован')
    } else {
      console.log('InitialSetup: пропущена синхронизация HeroMapping')
    }
    
    // ШАГ 3: Автоопределение героев
    const detectResponse = ui.alert(
      'Шаг 3/4: Автоопределение героев',
      'Автоматически определить героев для предметов? (может занять несколько минут)',
      ui.ButtonSet.YES_NO
    )
    
    if (detectResponse === ui.Button.YES) {
      ui.alert('Начинаем автоопределение героев...')
      heroMapping_autoDetectFromSteamWebAPI()
      results.heroesDetected = true
      console.log('InitialSetup: герои определены')
    } else {
      console.log('InitialSetup: пропущено автоопределение героев')
    }
    
    // ШАГ 4: Первое обновление статистики героев
    const statsResponse = ui.alert(
      'Шаг 4/4: Обновление статистики героев',
      'Обновить статистику героев через OpenDota API? (может занять несколько минут)',
      ui.ButtonSet.YES_NO
    )
    
    if (statsResponse === ui.Button.YES) {
      ui.alert('Начинаем обновление статистики героев...')
      heroStats_updateAllStats()
      results.statsUpdated = true
      console.log('InitialSetup: статистика героев обновлена')
    } else {
      console.log('InitialSetup: пропущено обновление статистики героев')
    }
    
    // Итоговый отчет
    const completed = Object.values(results).filter(v => v === true).length
    const total = Object.keys(results).length
    
    let summary = `✅ Первоначальная настройка завершена!\n\nВыполнено шагов: ${completed}/${total}\n\n`
    
    if (results.minMaxCalculated) summary += '✅ Расчет Min/Max цен\n'
    if (results.heroMappingSynced) summary += '✅ Синхронизация HeroMapping\n'
    if (results.heroesDetected) summary += '✅ Автоопределение героев\n'
    if (results.statsUpdated) summary += '✅ Обновление статистики героев\n'
    
    summary += '\nСледующие шаги:\n'
    summary += '• Используйте "Полная настройка" для автоматического обновления аналитики и метрик\n'
    summary += '• Или включите автообновление: SteamTable → Включить автообновление'
    
    logAutoAction_('InitialSetup', 'Первоначальная настройка', `Завершено: ${completed}/${total} шагов`)
    ui.alert('Настройка завершена', summary, ui.ButtonSet.OK)
    
  } catch (e) {
    console.error('InitialSetup: ошибка при выполнении настройки:', e)
    ui.alert(
      'Ошибка при выполнении настройки',
      'Произошла ошибка: ' + e.message + '\n\nПроверьте логи для деталей.',
      ui.ButtonSet.OK
    )
  }
}

/**
 * Шаг 2: Настройка HeroMapping
 * Синхронизирует предметы и определяет героев (объединенная операция)
 */
function setupHeroMapping() {
  const startTime = Date.now()
  
  try {
    console.log('Setup: настройка HeroMapping (синхронизация + автоопределение)...')
    // Объединенная операция: синхронизация и автоопределение в одном вызове
    // heroMapping_autoDetectFromSteamWebAPI с autoSync=true сначала синхронизирует, затем определяет героев
    heroMapping_autoDetectFromSteamWebAPI(true, true) // silent=true, autoSync=true
    console.log('Setup: HeroMapping настроен (предметы синхронизированы, герои определены)')
    
    const duration = ((Date.now() - startTime) / 1000).toFixed(1)
    console.log(`Setup: настройка HeroMapping завершена за ${duration} сек`)
    logAutoAction_('Setup', 'Настройка HeroMapping', `OK (${duration} сек)`)
    
  } catch (e) {
    console.error('Setup: ошибка настройки HeroMapping:', e)
    logAutoAction_('Setup', 'Настройка HeroMapping', `Ошибка: ${e.message}`)
    SpreadsheetApp.getUi().alert('Ошибка настройки HeroMapping: ' + e.message)
  }
}

/**
 * Шаг 3: Обновление статистики героев
 * Получает статистику через OpenDota API и синхронизирует в History
 */
/**
 * Шаг 3: Обновление статистики героев
 * 
 * ПРИМЕЧАНИЕ О ПРОИЗВОДИТЕЛЬНОСТИ:
 * Эта функция может выполняться 240-250 секунд (4+ минуты) для ~252 героев.
 * Это нормально, так как:
 * - Выполняется запрос к OpenDota API для каждого героя
 * - Обрабатываются данные для нескольких рангов (Immortal, Divine, Ancient, Legend, Archon, Crusader, Herald)
 * - Синхронизируется статистика в History для всех предметов
 * 
 * Оптимизация уже применена (кэширование, batch-операции), дальнейшее ускорение
 * возможно только за счет уменьшения количества запросов к API, что снизит качество данных.
 */
function setupHeroStats() {
  const startTime = Date.now()
  
  try {
    console.log('Setup: обновление статистики героев...')
    heroStats_updateAllStats()
    console.log('Setup: статистика героев обновлена')
    
    console.log('Setup: заполнение пустых Hero ID в HeroMapping...')
    const fillResult = heroMapping_fillMissingHeroIds()
    console.log(`Setup: заполнено ${fillResult.filled} Hero ID, не найдено ${fillResult.notFound}`)
    
    console.log('Setup: синхронизация статистики в History...')
    history_syncHeroStats()
    console.log('Setup: статистика синхронизирована в History')
    
    const duration = ((Date.now() - startTime) / 1000).toFixed(1)
    console.log(`Setup: настройка статистики героев завершена за ${duration} сек`)
    logAutoAction_('Setup', 'Настройка статистики героев', `OK (${duration} сек, заполнено Hero ID: ${fillResult.filled})`)
    
  } catch (e) {
    console.error('Setup: ошибка настройки статистики героев:', e)
    logAutoAction_('Setup', 'Настройка статистики героев', `Ошибка: ${e.message}`)
    SpreadsheetApp.getUi().alert('Ошибка настройки статистики героев: ' + e.message)
  }
}

/**
 * Шаг 1: Расчет Min/Max из SteamWebAPI (опционально)
 * Получает Min/Max для всех предметов из SteamWebAPI
 * Выполняется только если Min/Max отсутствуют
 */
function setupMinMax() {
  const startTime = Date.now()
  
  try {
    const historySheet = getHistorySheet_()
    if (!historySheet) {
      console.warn('Setup: лист History не найден, пропускаем Min/Max')
      return
    }
    
    const lastRow = historySheet.getLastRow()
    if (lastRow < DATA_START_ROW) {
      console.warn('Setup: нет предметов в History, пропускаем Min/Max')
      return
    }
    
    // Проверяем, нужен ли расчет Min/Max
    const names = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.NAME), lastRow - HEADER_ROW, 1).getValues()
    const minBatch = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.MIN_PRICE), lastRow - HEADER_ROW, 1).getValues()
    const maxBatch = historySheet.getRange(DATA_START_ROW, getColumnIndex(HISTORY_COLUMNS.MAX_PRICE), lastRow - HEADER_ROW, 1).getValues()
    
    let missingCount = 0
    for (let i = 0; i < names.length; i++) {
      const name = String(names[i][0] || '').trim()
      if (!name) continue
      
      const minValue = minBatch[i][0]
      const maxValue = maxBatch[i][0]
      const hasMin = minValue !== null && minValue !== '' && Number.isFinite(Number(minValue)) && Number(minValue) > 0
      const hasMax = maxValue !== null && maxValue !== '' && Number.isFinite(Number(maxValue)) && Number(maxValue) > 0
      
      if (!hasMin || !hasMax) {
        missingCount++
      }
    }
    
    if (missingCount === 0) {
      console.log('Setup: все Min/Max уже заполнены, пропускаем')
      return
    }
    
    console.log(`Setup: требуется расчет Min/Max для ${missingCount} предметов...`)
    priceHistory_calculateMinMaxForAllItems(true) // onlyMissing = true
    console.log('Setup: Min/Max рассчитаны')
    
    const duration = ((Date.now() - startTime) / 1000).toFixed(1)
    console.log(`Setup: настройка Min/Max завершена за ${duration} сек`)
    logAutoAction_('Setup', 'Настройка Min/Max', `OK (${duration} сек, ${missingCount} предметов)`)
    
  } catch (e) {
    console.error('Setup: ошибка настройки Min/Max:', e)
    logAutoAction_('Setup', 'Настройка Min/Max', `Ошибка: ${e.message}`)
    // Не показываем alert, так как это опциональный шаг
  }
}

/**
 * Шаг 4: Обновление аналитики и метрик
 * Обновляет аналитику History (включая Min/Max из существующих цен) и Investment Scores
 */
function setupAnalytics() {
  const startTime = Date.now()
  
  try {
    console.log('Setup: обновление аналитики History...')
    // history_updateAllAnalytics_ включает обновление Min/Max из существующих цен в колонках
    // Пропускаем синхронизацию статистики героев, так как она уже выполнена в setupHeroStats
    history_updateAllAnalytics_(true) // skipHeroStats = true
    console.log('Setup: аналитика History обновлена (включая Min/Max из существующих цен, статистика героев пропущена - уже обновлена)')
    
    console.log('Setup: обновление Investment Scores...')
    history_updateInvestmentScores()
    console.log('Setup: Investment Scores обновлены')
    
    const duration = ((Date.now() - startTime) / 1000).toFixed(1)
    console.log(`Setup: настройка аналитики завершена за ${duration} сек`)
    logAutoAction_('Setup', 'Настройка аналитики', `OK (${duration} сек)`)
    
  } catch (e) {
    console.error('Setup: ошибка настройки аналитики:', e)
    logAutoAction_('Setup', 'Настройка аналитики', `Ошибка: ${e.message}`)
    SpreadsheetApp.getUi().alert('Ошибка настройки аналитики: ' + e.message)
  }
}

/**
 * Полная настройка таблицы - последовательный вызов всех шагов
 * Предполагается, что предметы уже добавлены в History вручную
 */
function performFullSetup() {
  const ui = SpreadsheetApp.getUi()
  
  const response = ui.alert(
    'Полная настройка',
    'Выполнит все шаги настройки последовательно:\n\n' +
    'Шаг 1: Расчет Min/Max из SteamWebAPI (если отсутствуют)\n' +
    'Шаг 2: Настройка HeroMapping\n' +
    'Шаг 3: Обновление статистики героев\n' +
    'Шаг 4: Обновление аналитики и метрик\n\n' +
    '⚠️ Убедитесь, что:\n' +
    '• Таблицы инициализированы\n' +
    '• Предметы добавлены в History\n\n' +
    'Продолжить?',
    ui.ButtonSet.YES_NO
  )
  
  if (response !== ui.Button.YES) {
    return
  }
  
  const totalStartTime = Date.now()
  
  try {
    setupMinMax() // Шаг 1: Min/Max из SteamWebAPI (опционально)
    setupHeroMapping() // Шаг 2: Настройка HeroMapping
    setupHeroStats() // Шаг 3: Обновление статистики героев
    setupAnalytics() // Шаг 4: Обновление аналитики и метрик
    
    const totalDuration = ((Date.now() - totalStartTime) / 1000).toFixed(1)
    console.log(`FullSetup: полная настройка завершена за ${totalDuration} сек`)
    logAutoAction_('FullSetup', 'Полная настройка', `Завершено (${totalDuration} сек)`)
    
    ui.alert(
      'Настройка завершена',
      `✅ Все шаги выполнены успешно!\n\nВремя: ${totalDuration} сек\n\n` +
      `💡 Следующие шаги:\n` +
      `• Метрики Invest/Sales обновятся автоматически при использовании\n` +
      `• Включите автообновление: SteamTable → Включить автообновление`,
      ui.ButtonSet.OK
    )
    
  } catch (e) {
    console.error('FullSetup: критическая ошибка:', e)
    logAutoAction_('FullSetup', 'Полная настройка', `Критическая ошибка: ${e.message}`)
    ui.alert('Ошибка', 'Произошла критическая ошибка: ' + e.message, ui.ButtonSet.OK)
  }
}

/**
 * Проверка готовности системы
 * Проверяет наличие данных в History, заполненность Min/Max, наличие HeroMapping, наличие статистики героев
 */
function checkSystemReadiness() {
  const ui = SpreadsheetApp.getUi()
  
  try {
    const checks = {
      historyHasData: false,
      minMaxFilled: false,
      heroMappingExists: false,
      heroStatsExists: false,
      triggersEnabled: false
    }
    
    let report = '🔍 ПРОВЕРКА ГОТОВНОСТИ СИСТЕМЫ\n\n'
    
    // Проверка 1: История содержит данные
    try {
      const historySheet = getHistorySheet_()
      if (historySheet && historySheet.getLastRow() >= DATA_START_ROW) {
        const itemCount = historySheet.getLastRow() - HEADER_ROW
        checks.historyHasData = true
        report += `✅ History: ${itemCount} предметов\n`
      } else {
        report += `❌ History: нет данных\n`
      }
    } catch (e) {
      report += `❌ History: ошибка проверки\n`
    }
    
    // Проверка 2: Min/Max заполнены
    try {
      const historySheet = getHistorySheet_()
      if (historySheet && historySheet.getLastRow() >= DATA_START_ROW) {
        const minCol = getColumnIndex(HISTORY_COLUMNS.MIN_PRICE)
        const maxCol = getColumnIndex(HISTORY_COLUMNS.MAX_PRICE)
        
        if (!minCol || !maxCol) {
          report += `❌ Min/Max: не найдены колонки (Min: ${minCol}, Max: ${maxCol})\n`
        } else {
          const rowCount = historySheet.getLastRow() - HEADER_ROW
          const minMaxValues = historySheet.getRange(DATA_START_ROW, minCol, rowCount, 2).getValues()
          
          const filledCount = minMaxValues.filter(row => {
            const minVal = row[0]
            const maxVal = row[1]
            return minVal !== null && minVal !== '' && !isNaN(Number(minVal)) && Number(minVal) > 0 &&
                   maxVal !== null && maxVal !== '' && !isNaN(Number(maxVal)) && Number(maxVal) > 0
          }).length
          const totalCount = minMaxValues.length
          const fillPercentage = totalCount > 0 ? (filledCount / totalCount * 100).toFixed(0) : 0
          
          if (fillPercentage >= 80) {
            checks.minMaxFilled = true
            report += `✅ Min/Max: заполнено ${fillPercentage}% (${filledCount}/${totalCount})\n`
          } else {
            report += `⚠️ Min/Max: заполнено только ${fillPercentage}% (${filledCount}/${totalCount})\n`
          }
        }
      } else {
        report += `❌ Min/Max: нет данных в History\n`
      }
    } catch (e) {
      console.error('checkSystemReadiness: ошибка проверки Min/Max:', e)
      report += `❌ Min/Max: ошибка проверки (${e.message})\n`
    }
    
    // Проверка 3: HeroMapping существует и содержит данные
    try {
      const mappingSheet = getHeroMappingSheet_()
      if (mappingSheet && mappingSheet.getLastRow() >= DATA_START_ROW) {
        const mappingCount = mappingSheet.getLastRow() - HEADER_ROW
        checks.heroMappingExists = true
        report += `✅ HeroMapping: ${mappingCount} записей\n`
      } else {
        report += `❌ HeroMapping: нет данных\n`
      }
    } catch (e) {
      report += `❌ HeroMapping: ошибка проверки\n`
    }
    
    // Проверка 4: HeroStats содержит данные
    try {
      const heroStatsSheet = getHeroStatsSheet_()
      if (heroStatsSheet) {
        const lastRow = heroStatsSheet.getLastRow()
        const lastCol = heroStatsSheet.getLastColumn()
        const firstDataCol = HERO_STATS_COLUMNS.FIRST_DATA_COL
        
        if (lastRow >= DATA_START_ROW && lastCol > firstDataCol) {
          // Проверяем, есть ли хотя бы одна непустая ячейка в колонках с данными
          const hasData = heroStatsSheet.getRange(DATA_START_ROW, firstDataCol, lastRow - HEADER_ROW, lastCol - firstDataCol + 1).getValues()
            .some(row => row.some(cell => cell !== null && cell !== ''))
          
          if (hasData) {
            checks.heroStatsExists = true
            const statsCount = lastRow - HEADER_ROW // Общее количество строк (может быть нечетным)
            const recordsCount = lastCol - firstDataCol + 1
            report += `✅ HeroStats: ${statsCount} строк, ${recordsCount} записей статистики\n`
          } else {
            report += `⚠️ HeroStats: лист существует, но нет записей статистики\n`
          }
        } else if (lastRow < DATA_START_ROW) {
          report += `❌ HeroStats: нет данных (лист пуст)\n`
        } else if (lastCol <= firstDataCol) {
          report += `⚠️ HeroStats: нет записей статистики (только заголовки)\n`
        }
      } else {
        report += `❌ HeroStats: лист не найден\n`
      }
    } catch (e) {
      console.error('checkSystemReadiness: ошибка проверки HeroStats:', e)
      report += `❌ HeroStats: ошибка проверки (${e.message})\n`
    }
    
    // Проверка 5: Триггеры включены
    try {
      const triggers = ScriptApp.getProjectTriggers()
      if (triggers.length > 0) {
        checks.triggersEnabled = true
        report += `✅ Триггеры: ${triggers.length} активных\n`
      } else {
        report += `⚠️ Триггеры: не включены\n`
      }
    } catch (e) {
      report += `❌ Триггеры: ошибка проверки\n`
    }
    
    // Итоговая оценка
    const passedChecks = Object.values(checks).filter(v => v === true).length
    const totalChecks = Object.keys(checks).length
    
    report += `\n📊 РЕЗУЛЬТАТ: ${passedChecks}/${totalChecks} проверок пройдено\n\n`
    
    if (passedChecks === totalChecks) {
      report += '🎉 Система полностью готова к работе!'
    } else if (passedChecks >= totalChecks - 1) {
      report += '✅ Система почти готова. Рекомендуется выполнить недостающие шаги.'
    } else {
      report += '⚠️ Система требует настройки. Используйте:\n'
      if (!checks.historyHasData) report += '• Добавьте предметы в History\n'
      if (!checks.minMaxFilled) report += '• Рассчитайте Min/Max цены\n'
      if (!checks.heroMappingExists) report += '• Синхронизируйте HeroMapping\n'
      if (!checks.heroStatsExists) report += '• Обновите статистику героев\n'
      if (!checks.triggersEnabled) report += '• Включите автообновление\n'
    }
    
    ui.alert('Проверка готовности', report, ui.ButtonSet.OK)
    
  } catch (e) {
    console.error('checkSystemReadiness: ошибка:', e)
    ui.alert('Ошибка', 'Не удалось проверить готовность системы: ' + e.message, ui.ButtonSet.OK)
  }
}

/**
 * Проверка состояния автоматизации
 * Показывает активные триггеры и последние выполнения операций
 */
function checkAutomationStatus() {
  const ui = SpreadsheetApp.getUi()
  
  try {
    let report = '🤖 ПРОВЕРКА СОСТОЯНИЯ АВТОМАТИЗАЦИИ\n\n'
    
    // Проверка триггеров
    try {
      const triggers = ScriptApp.getProjectTriggers()
      if (triggers.length === 0) {
        report += '❌ Триггеры: не включены\n'
        report += '\n💡 Используйте: SteamTable → Включить автообновление\n'
      } else {
        report += `✅ Триггеры: ${triggers.length} активных\n\n`
        
        // Группируем триггеры по функциям
        const triggerGroups = {}
        triggers.forEach(trigger => {
          const handler = trigger.getHandlerFunction()
          if (!triggerGroups[handler]) {
            triggerGroups[handler] = []
          }
          triggerGroups[handler].push(trigger)
        })
        
        // Показываем информацию о каждом типе триггера
        for (const [handler, triggerList] of Object.entries(triggerGroups)) {
          const count = triggerList.length
          const firstTrigger = triggerList[0]
          
          let schedule = ''
          if (firstTrigger.getEventType() === ScriptApp.EventType.CLOCK) {
            const timeBased = firstTrigger.getTimeBasedTriggerSource()
            if (timeBased === ScriptApp.TimeBasedTriggerSource.CLOCK) {
              schedule = 'по расписанию'
            } else if (timeBased === ScriptApp.TimeBasedTriggerSource.MINUTES) {
              const everyMinutes = firstTrigger.getEveryMinutes()
              schedule = `каждые ${everyMinutes} минут`
            } else if (timeBased === ScriptApp.TimeBasedTriggerSource.HOURS) {
              const everyHours = firstTrigger.getEveryHours()
              schedule = `каждые ${everyHours} часов`
            } else if (timeBased === ScriptApp.TimeBasedTriggerSource.DAYS) {
              const everyDays = firstTrigger.getEveryDays()
              if (everyDays === 1) {
                const hour = firstTrigger.getHour()
                schedule = `ежедневно в ${hour}:00`
              } else {
                schedule = `каждые ${everyDays} дней`
              }
            } else if (timeBased === ScriptApp.TimeBasedTriggerSource.WEEKS) {
              schedule = 'еженедельно'
            }
          }
          
          report += `  • ${handler}${count > 1 ? ` (${count})` : ''}${schedule ? ` - ${schedule}` : ''}\n`
        }
      }
    } catch (e) {
      report += `❌ Триггеры: ошибка проверки (${e.message})\n`
    }
    
    // Проверка последних выполнений из AutoLog
    try {
      const autoLogSheet = getAutoLogSheet_()
      if (autoLogSheet && autoLogSheet.getLastRow() >= DATA_START_ROW) {
        const lastRow = autoLogSheet.getLastRow()
        const lastEntries = autoLogSheet.getRange(Math.max(DATA_START_ROW, lastRow - 4), 1, Math.min(5, lastRow - 1), 3).getValues()
        
        if (lastEntries.length > 0) {
          report += '\n📋 Последние операции:\n'
          lastEntries.reverse().forEach((row, index) => {
            const date = row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'dd.MM.yy HH:mm') : '—'
            const action = row[1] || '—'
            const status = row[2] || '—'
            report += `  ${index + 1}. ${date} - ${action} (${status})\n`
          })
        }
      }
    } catch (e) {
      report += `\n⚠️ Не удалось получить историю операций: ${e.message}\n`
    }
    
    ui.alert('Проверка автоматизации', report, ui.ButtonSet.OK)
    
  } catch (e) {
    console.error('checkAutomationStatus: ошибка:', e)
    ui.alert('Ошибка', 'Не удалось проверить состояние автоматизации: ' + e.message, ui.ButtonSet.OK)
  }
}

// Единая синхронизация аналитики для Invest/Sales
function syncAnalyticsForInvestSales_() {
  try {
    invest_syncMinMaxFromHistory()
    invest_syncTrendDaysFromHistory()
    invest_syncExtendedAnalyticsFromHistory()
    sales_syncMinMaxFromHistory()
    sales_syncTrendDaysFromHistory()
    sales_syncExtendedAnalyticsFromHistory()
    
    // Обновляем метрики после синхронизации аналитики
    updateAllMetricsForInvestSales()
    
    SpreadsheetApp.getUi().alert('Аналитика и метрики синхронизированы (Invest/Sales)')
  } catch (e) {
    console.error('Menu: ошибка синхронизации аналитики:', e)
    SpreadsheetApp.getUi().alert('Ошибка синхронизации аналитики')
  }
}

/**
 * Комплексное обновление всех метрик и скоров для Invest/Sales/History
 * Оптимизированная функция для ручного использования
 */
function updateAllMetricsAndScores_() {
  const ui = SpreadsheetApp.getUi()
  
  const response = ui.alert(
    'Обновление всех метрик и скоров',
    'Эта функция обновит:\n\n' +
    '• Метрики для Invest/Sales (Liquidity, Demand, Momentum, Sales Trend, Volatility)\n' +
    '• Investment Scores для History и Invest\n' +
    '• Buyback Scores для Sales\n' +
    '• Синхронизацию статистики героев\n\n' +
    'Это может занять несколько минут. Продолжить?',
    ui.ButtonSet.YES_NO
  )
  
  if (response !== ui.Button.YES) {
    return
  }
  
  try {
    ui.alert('Начинаем обновление метрик и скоров...')
    
    // 1. Обновляем метрики для Invest
    console.log('Menu: обновление метрик Invest...')
    invest_calculateAllMetrics()
    
    // 2. Обновляем метрики для Sales
    console.log('Menu: обновление метрик Sales...')
    sales_calculateAllMetrics()
    
    // 3. Обновляем Investment Scores для History
    console.log('Menu: обновление Investment Scores для History...')
    history_updateInvestmentScores()
    
    // 4. Обновляем Investment Scores для Invest
    console.log('Menu: обновление Investment Scores для Invest...')
    invest_updateInvestmentScores()
    
    // 5. Обновляем Buyback Scores для Sales
    console.log('Menu: обновление Buyback Scores для Sales...')
    sales_updateBuybackScores()
    
    // 6. Синхронизируем статистику героев
    console.log('Menu: синхронизация статистики героев...')
    history_syncHeroStats()
    
    ui.alert('✅ Все метрики и скоры обновлены!')
    logAutoAction_('Menu', 'Обновление всех метрик и скоров', 'Завершено')
    
  } catch (e) {
    console.error('Menu: ошибка при обновлении метрик и скоров:', e)
    ui.alert('Ошибка при обновлении метрик и скоров: ' + e.message)
  }
}

/**
 * Автоматическое обновление статистики героев (для триггера)
 * Выполняет полный цикл обновления после получения новой статистики:
 * 1. Обновление статистики героев через OpenDota API
 * 2. Синхронизация статистики в History
 * 3. Обновление Investment Scores для History и Invest
 * 4. Обновление Buyback Scores для Sales
 */
function autoUpdateHeroStats() {
  const lockKey = 'autoUpdateHeroStats'
  
  try {
    // Проверяем блокировку (таймаут 5 минут)
    const lockResult = acquireLock_(lockKey, LIMITS.LOCK_TIMEOUT_SEC)
    if (lockResult.locked) {
      console.warn('autoUpdateHeroStats: операция уже выполняется, пропускаем')
      return
    }
    
    console.log('autoUpdateHeroStats: начало автоматического обновления статистики героев')
    logAutoAction_('AutoUpdate', 'Обновление статистики героев', 'Начало')
    
    // 1. Обновление статистики героев
    try {
      heroStats_updateAllStats()
      console.log('autoUpdateHeroStats: статистика героев обновлена')
    } catch (e) {
      console.error('autoUpdateHeroStats: ошибка обновления статистики героев:', e)
      logAutoAction_('AutoUpdate', 'Обновление статистики героев', `Ошибка: ${e.message}`)
      releaseLock_(lockKey)
      return
    }
    
    // 2. Синхронизация статистики в History
    try {
      history_syncHeroStats()
      console.log('autoUpdateHeroStats: статистика синхронизирована в History')
    } catch (e) {
      console.error('autoUpdateHeroStats: ошибка синхронизации статистики:', e)
      logAutoAction_('AutoUpdate', 'Синхронизация статистики', `Ошибка: ${e.message}`)
      // Продолжаем выполнение, даже если синхронизация не удалась
    }
    
    // 3. Обновление Investment Scores для History
    try {
      history_updateInvestmentScores()
      console.log('autoUpdateHeroStats: Investment Scores обновлены для History')
    } catch (e) {
      console.error('autoUpdateHeroStats: ошибка обновления Investment Scores для History:', e)
      logAutoAction_('AutoUpdate', 'Обновление Investment Scores (History)', `Ошибка: ${e.message}`)
    }
    
    // 4. Обновление Investment Scores для Invest
    try {
      invest_updateInvestmentScores()
      console.log('autoUpdateHeroStats: Investment Scores обновлены для Invest')
    } catch (e) {
      console.error('autoUpdateHeroStats: ошибка обновления Investment Scores для Invest:', e)
      logAutoAction_('AutoUpdate', 'Обновление Investment Scores (Invest)', `Ошибка: ${e.message}`)
    }
    
    // 5. Обновление Buyback Scores для Sales
    try {
      sales_updateBuybackScores()
      console.log('autoUpdateHeroStats: Buyback Scores обновлены для Sales')
    } catch (e) {
      console.error('autoUpdateHeroStats: ошибка обновления Buyback Scores:', e)
      logAutoAction_('AutoUpdate', 'Обновление Buyback Scores', `Ошибка: ${e.message}`)
    }
    
    console.log('autoUpdateHeroStats: автоматическое обновление завершено')
    logAutoAction_('AutoUpdate', 'Обновление статистики героев', 'Завершено')
    
  } catch (e) {
    console.error('autoUpdateHeroStats: критическая ошибка:', e)
    logAutoAction_('AutoUpdate', 'Обновление статистики героев', `Критическая ошибка: ${e.message}`)
    
    // Отправляем уведомление в Telegram, если настроено
    try {
      const telegramConfig = telegram_getConfig()
      if (telegramConfig && telegramConfig.botToken && telegramConfig.chatId) {
        telegram_sendMessage(
          `⚠️ <b>Ошибка автоматического обновления статистики героев</b>\n\n` +
          `Ошибка: ${e.message}`,
          'HTML'
        )
      }
    } catch (telegramError) {
      // Игнорируем ошибки Telegram
      console.warn('autoUpdateHeroStats: не удалось отправить уведомление в Telegram:', telegramError)
    }
  } finally {
    releaseLock_(lockKey)
  }
}

/**
 * Автоматическая архивация старых данных HeroStats (для триггера)
 * Удаляет данные старше HERO_STATS_HISTORY_DAYS дней
 */
function autoArchiveHeroStats() {
  const lockKey = 'autoArchiveHeroStats'
  
  try {
    // Проверяем блокировку (таймаут 5 минут)
    const lockResult = acquireLock_(lockKey, LIMITS.LOCK_TIMEOUT_SEC)
    if (lockResult.locked) {
      console.warn('autoArchiveHeroStats: операция уже выполняется, пропускаем')
      return
    }
    
    console.log('autoArchiveHeroStats: начало автоматической архивации')
    logAutoAction_('AutoArchive', 'Архивация HeroStats', 'Начало')
    
    heroStats_archiveOldData()
    
    console.log('autoArchiveHeroStats: архивация завершена')
    logAutoAction_('AutoArchive', 'Архивация HeroStats', 'Завершено')
    
    // Очистка старых уведомлений Telegram (старше 7 дней)
    try {
      telegram_cleanupOldNotifications_()
      console.log('autoArchiveHeroStats: очистка старых уведомлений завершена')
    } catch (e) {
      console.error('autoArchiveHeroStats: ошибка очистки уведомлений:', e)
      // Не прерываем выполнение, просто логируем ошибку
    }
    
  } catch (e) {
    console.error('autoArchiveHeroStats: ошибка архивации:', e)
    logAutoAction_('AutoArchive', 'Архивация HeroStats', `Ошибка: ${e.message}`)
    
    // Отправляем уведомление в Telegram, если настроено
    try {
      const telegramConfig = telegram_getConfig()
      if (telegramConfig && telegramConfig.botToken && telegramConfig.chatId) {
        telegram_sendMessage(
          `⚠️ <b>Ошибка автоматической архивации HeroStats</b>\n\n` +
          `Ошибка: ${e.message}`,
          'HTML'
        )
      }
    } catch (telegramError) {
      // Игнорируем ошибки Telegram
      console.warn('autoArchiveHeroStats: не удалось отправить уведомление в Telegram:', telegramError)
    }
  } finally {
    releaseLock_(lockKey)
  }
}

/**
 * Автоматическая синхронизация HeroMapping (для триггера)
 * Синхронизирует предметы из History и автоматически определяет героев для новых предметов
 */
function autoSyncHeroMapping() {
  const lockKey = 'autoSyncHeroMapping'
  
  try {
    // Проверяем блокировку (таймаут 5 минут)
    const lockResult = acquireLock_(lockKey, LIMITS.LOCK_TIMEOUT_SEC)
    if (lockResult.locked) {
      console.warn('autoSyncHeroMapping: операция уже выполняется, пропускаем')
      return
    }
    
    console.log('autoSyncHeroMapping: начало автоматической синхронизации HeroMapping')
    logAutoAction_('AutoSync', 'Синхронизация HeroMapping', 'Начало')
    
    // 1. Синхронизируем предметы из History (silent режим для автоматических вызовов)
    try {
      heroMapping_syncWithHistory(true) // silent = true для автоматических вызовов
      console.log('autoSyncHeroMapping: предметы синхронизированы с History')
    } catch (e) {
      console.error('autoSyncHeroMapping: ошибка синхронизации предметов:', e)
      logAutoAction_('AutoSync', 'Синхронизация предметов', `Ошибка: ${e.message}`)
      releaseLock_(lockKey)
      return
    }
    
    // 2. Автоматически определяем героев для новых предметов (silent режим для автоматических вызовов)
    try {
      heroMapping_autoDetectFromSteamWebAPI(true) // silent = true для автоматических вызовов
      console.log('autoSyncHeroMapping: герои определены для новых предметов')
    } catch (e) {
      console.error('autoSyncHeroMapping: ошибка автоопределения героев:', e)
      logAutoAction_('AutoSync', 'Автоопределение героев', `Ошибка: ${e.message}`)
      // Продолжаем выполнение, даже если автоопределение не удалось
    }
    
    // 3. Заполняем пустые Hero ID из HeroStats (если есть данные)
    try {
      const fillResult = heroMapping_fillMissingHeroIds()
      if (fillResult.filled > 0) {
        console.log(`autoSyncHeroMapping: заполнено ${fillResult.filled} Hero ID из HeroStats`)
        logAutoAction_('AutoSync', 'Заполнение Hero ID', `Заполнено: ${fillResult.filled}`)
      }
    } catch (e) {
      console.error('autoSyncHeroMapping: ошибка заполнения Hero ID:', e)
      // Продолжаем выполнение, даже если заполнение не удалось
    }
    
    console.log('autoSyncHeroMapping: автоматическая синхронизация завершена')
    logAutoAction_('AutoSync', 'Синхронизация HeroMapping', 'Завершено')
    
  } catch (e) {
    console.error('autoSyncHeroMapping: критическая ошибка:', e)
    logAutoAction_('AutoSync', 'Синхронизация HeroMapping', `Критическая ошибка: ${e.message}`)
    
    // Отправляем уведомление в Telegram, если настроено
    try {
      const telegramConfig = telegram_getConfig()
      if (telegramConfig && telegramConfig.botToken && telegramConfig.chatId) {
        telegram_sendMessage(
          `⚠️ <b>Ошибка автоматической синхронизации HeroMapping</b>\n\n` +
          `Ошибка: ${e.message}`,
          'HTML'
        )
      }
    } catch (telegramError) {
      // Игнорируем ошибки Telegram
      console.warn('autoSyncHeroMapping: не удалось отправить уведомление в Telegram:', telegramError)
    }
  } finally {
    releaseLock_(lockKey)
  }
}

/**
 * Меню для заполнения пустых Hero ID в HeroMapping
 */
function heroMapping_fillMissingHeroIdsMenu() {
  const ui = SpreadsheetApp.getUi()
  
  const response = ui.alert(
    'Заполнение пустых Hero ID',
    'Эта функция заполнит пустые Hero ID в HeroMapping, используя данные из HeroStats.\n\n' +
    'Для этого нужны:\n' +
    '• Заполненные Hero Name в HeroMapping\n' +
    '• Обновленная статистика в HeroStats\n\n' +
    'Продолжить?',
    ui.ButtonSet.YES_NO
  )
  
  if (response !== ui.Button.YES) {
    return
  }
  
  try {
    ui.alert('Начинаем заполнение Hero ID...')
    
    const result = heroMapping_fillMissingHeroIds()
    
    ui.alert(
      'Заполнение завершено',
      `✅ Заполнено Hero ID: ${result.filled}\n` +
      (result.notFound > 0 ? `⚠️ Не найдено: ${result.notFound}\n\n` : '\n') +
      (result.notFound > 0 
        ? 'Не найденные герои могут быть:\n' +
          '• Не обновлены в HeroStats\n' +
          '• Имеют другое имя в HeroStats\n'
        : 'Все Hero ID успешно заполнены!'),
      ui.ButtonSet.OK
    )
    
    logAutoAction_('HeroMapping', 'Заполнение Hero ID', `Заполнено: ${result.filled}, не найдено: ${result.notFound}`)
    
  } catch (e) {
    console.error('HeroMapping: ошибка заполнения Hero ID:', e)
    ui.alert('Ошибка', 'Произошла ошибка при заполнении Hero ID: ' + e.message, ui.ButtonSet.OK)
  }
}