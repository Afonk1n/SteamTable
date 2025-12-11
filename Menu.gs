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
    .addSeparator()
    .addItem('Обновить цены History (ручное)', 'history_updateAllPrices')
    .addItem('Рассчитать Min/Max для всех предметов', 'priceHistory_calculateMinMaxForAllItems')
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
      .addItem('Настроить Stratz API', 'stratz_setup')
      .addItem('Тест Stratz API', 'stratz_testConnection')
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
