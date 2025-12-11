/**
 * Menu - Меню и триггеры
 * 
 * Управляет созданием меню, триггеров и обработкой событий
 */

// Настройка всех триггеров
// Архитектура:
// 1. Два фиксированных триггера (00:10 и 12:00) для создания колонок периодов и начала сбора
// 2. Один периодический триггер (каждые 10 минут) для продолжения сбора до завершения периода
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
    .addItem('Первоначальная настройка', 'performInitialSetup')
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
    )
    .addSubMenu(ui.createMenu('Sales')
      .addItem('Форматирование', 'sales_formatTable')
      .addItem('Изображение и ссылки', 'sales_updateImagesAndLinks')
      .addItem('Поиск дублей', 'sales_findDuplicates')
    )
    .addSubMenu(ui.createMenu('History')
      .addItem('Форматирование', 'history_formatTable')
      .addItem('Обновить аналитику', 'history_updateAllAnalytics_')
      .addItem('Изображения и ссылки', 'history_updateImagesAndLinks')
      .addItem('Дубли названий', 'history_findDuplicates')
      .addItem('Создать столбец текущего периода', 'history_ensureTodayColumn')
    )
    .addSubMenu(ui.createMenu('PortfolioStats')
      .addItem('Форматирование', 'portfolioStats_formatTable')
      .addItem('Сохранить строку истории (тест)', 'portfolioStats_saveHistoryManual')
    )
    .addSubMenu(ui.createMenu('HeroStats')
      .addItem('Форматирование', 'heroStats_formatTable')
      .addItem('Обновить статистику (ручное)', 'heroStats_updateAllStats')
      .addItem('Архивировать старые данные', 'heroStats_archiveOldData')
    )
    .addSubMenu(ui.createMenu('HeroMapping')
      .addItem('Форматирование', 'heroMapping_formatTable')
      .addItem('Автоопределение героев', 'heroMapping_autoDetectFromSteamWebAPI')
      .addItem('Синхронизировать с History', 'heroMapping_syncWithHistory')
    )
    .addSeparator()
    .addItem('Синхронизировать цены из History', 'syncPricesFromHistoryToInvestAndSales')
    .addItem('Обновить аналитику Invest/Sales', 'syncAnalyticsForInvestSales_')
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
    SpreadsheetApp.getUi().alert('✅ Все таблицы инициализированы и отформатированы')
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
    summary += '• Включите автообновление: SteamTable → Включить автообновление'
    
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

// Единая синхронизация аналитики для Invest/Sales
function syncAnalyticsForInvestSales_() {
  try {
    invest_syncMinMaxFromHistory()
    invest_syncTrendDaysFromHistory()
    invest_syncExtendedAnalyticsFromHistory()
    sales_syncMinMaxFromHistory()
    sales_syncTrendDaysFromHistory()
    sales_syncExtendedAnalyticsFromHistory()
    SpreadsheetApp.getUi().alert('Аналитика синхронизирована (Invest/Sales)')
  } catch (e) {
    console.error('Menu: ошибка синхронизации аналитики:', e)
    SpreadsheetApp.getUi().alert('Ошибка синхронизации аналитики')
  }
}
