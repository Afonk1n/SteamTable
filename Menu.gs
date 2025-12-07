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
    .addItem('Синхронизировать цены из History', 'syncPricesFromHistoryToInvestAndSales')
    .addItem('Обновить цены History (ручное)', 'history_updateAllPrices')
    .addSeparator()
    // Сокращённые меню
    .addSubMenu(ui.createMenu('Invest')
      .addItem('Форматирование', 'invest_formatTable')
      .addItem('Изображение и ссылки', 'invest_updateImagesAndLinks')
      .addItem('Поиск дублей', 'invest_findDuplicates')
      .addItem('Обновить аналитику портфеля', 'portfolioStats_formatTable')
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
    .addSeparator()
    .addItem('Обновить аналитику Invest/Sales', 'syncAnalyticsForInvestSales_')
    .addSeparator()
    .addSubMenu(ui.createMenu('Telegram')
      .addItem('Настроить Telegram', 'telegram_setup')
      .addItem('Тест Telegram', 'telegram_testConnection')
    )
    .addToUi()
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
