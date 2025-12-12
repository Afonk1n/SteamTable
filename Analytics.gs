/**
 * Analytics - Расчет экономических метрик и скоров
 * 
 * Реализует систему инвестиционных рекомендаций на основе:
 * - Параметров предметов (цены, объемы, ликвидность, волатильность)
 * - Параметров героев (пикрейт, винрейт, банрейт, контестрейт)
 * - Корреляций между метой героя и ценами предметов
 */

// === ОСНОВНЫЕ МЕТРИКИ ===

/**
 * Расчет Liquidity Score (оценка ликвидности)
 * @param {Object} itemData - Данные предмета из SteamWebAPI.ru
 * @returns {number} Score от 0 до 1
 */
function analytics_calculateLiquidityScore(itemData) {
  if (!itemData) return 0.5
  
  const sold24h = itemData.sold24h || 0
  const sold7d = itemData.sold7d || 0
  const offervolume = itemData.offervolume || 1
  const hourstosold = itemData.hourstosold || 24
  
  // Компоненты формулы
  const sold24hComponent = sold24h * 0.35
  const sold7dComponent = sold7d * 0.25
  const offerVolumeComponent = (1 / (1 + offervolume / 1000)) * 0.25
  const hoursToSoldComponent = (1 / (1 + hourstosold / 24)) * 0.15
  
  const rawScore = sold24hComponent + sold7dComponent + offerVolumeComponent + hoursToSoldComponent
  
  // Нормализация (примерные пороги, могут быть скорректированы)
  const maxSold24h = 10000
  const maxSold7d = 50000
  const normalizedSold24h = Math.min(sold24h / maxSold24h, 1) * 0.35
  const normalizedSold7d = Math.min(sold7d / maxSold7d, 1) * 0.25
  const normalizedOfferVolume = (1 / (1 + offervolume / 1000)) * 0.25
  const normalizedHoursToSold = (1 / (1 + hourstosold / 24)) * 0.15
  
  return Math.min(1, Math.max(0, normalizedSold24h + normalizedSold7d + normalizedOfferVolume + normalizedHoursToSold))
}

/**
 * Расчет Demand Ratio (соотношение спроса и предложения)
 * @param {Object} itemData - Данные предмета из SteamWebAPI.ru
 * @returns {number} Score от 0 до 1
 */
function analytics_calculateDemandRatio(itemData) {
  if (!itemData) return 0.5
  
  const buyordervolume = itemData.buyordervolume || 0
  const offervolume = itemData.offervolume || 1
  
  // Формула: buyordervolume / (buyordervolume + offervolume + 1)
  const ratio = buyordervolume / (buyordervolume + offervolume + 1)
  
  // Нормализация уже в диапазоне 0-1
  return Math.min(1, Math.max(0, ratio))
}

/**
 * Расчет Price Momentum (импульс цены)
 * @param {Object} itemData - Данные предмета из SteamWebAPI.ru
 * @param {Object} historyData - Данные из History (цены за периоды)
 * @returns {number} Score от 0 до 1
 */
function analytics_calculatePriceMomentum(itemData, historyData) {
  if (!itemData) return 0.5
  
  const currentPrice = itemData.pricelatest || itemData.pricelatestsell || 0
  if (!currentPrice || currentPrice <= 0) return 0.5
  
  // Приоритет: pricelatestsell7d/30d > priceavg7d/30d
  const price7d = itemData.pricelatestsell7d || itemData.priceavg7d || currentPrice
  const price30d = itemData.pricelatestsell30d || itemData.priceavg30d || currentPrice
  const priceavg7d = itemData.priceavg7d || currentPrice
  
  // Расчет компонентов
  const change7d = price7d > 0 ? (currentPrice - price7d) / price7d : 0
  const change30d = price30d > 0 ? (currentPrice - price30d) / price30d : 0
  const changeAvg7d = priceavg7d > 0 ? (currentPrice - priceavg7d) / priceavg7d : 0
  
  // Формула: взвешенная сумма
  const momentum = (change7d * 0.5) + (change30d * 0.3) + (changeAvg7d * 0.2)
  
  // Нормализация: -1 (сильное падение) до 1 (сильный рост), затем сдвиг к 0-1
  return analytics_normalizeToRange(momentum, -0.5, 0.5, 0, 1)
}

/**
 * Расчет Sales Trend (тренд продаж)
 * @param {Object} itemData - Данные предмета из SteamWebAPI.ru
 * @returns {number} Score от 0 до 1
 */
function analytics_calculateSalesTrend(itemData) {
  if (!itemData) return 0.5
  
  const sold24h = itemData.sold24h || 0
  const sold7d = itemData.sold7d || 0
  const sold30d = itemData.sold30d || 0
  
  if (sold7d === 0 && sold30d === 0) return 0.5
  
  // Формула: ((sold24h * 7 - sold7d) / sold7d * 0.5) + ((sold7d * 4 - sold30d) / sold30d * 0.5)
  let trend = 0
  
  if (sold7d > 0) {
    trend += ((sold24h * 7 - sold7d) / sold7d) * 0.5
  }
  
  if (sold30d > 0) {
    trend += ((sold7d * 4 - sold30d) / sold30d) * 0.5
  }
  
  // Нормализация: -1 (падение) до 1 (рост), затем сдвиг к 0-1
  return analytics_normalizeToRange(trend, -1, 1, 0, 1)
}

/**
 * Расчет Volatility Index (индекс волатильности)
 * @param {Object} itemData - Данные предмета из SteamWebAPI.ru
 * @param {Object} historyData - Данные из History (цены за периоды)
 * @returns {number} Score от 0 до 1 (высокая волатильность = лучше)
 */
function analytics_calculateVolatilityIndex(itemData, historyData) {
  if (!itemData && !historyData) return 0.5
  
  let minPrice, maxPrice, avgPrice, stdDev
  
  // Приоритет: данные из History (если есть)
  if (historyData && historyData.prices && historyData.prices.length > 0) {
    const prices = historyData.prices.filter(p => p > 0)
    if (prices.length > 0) {
      minPrice = Math.min(...prices)
      maxPrice = Math.max(...prices)
      avgPrice = prices.reduce((a, b) => a + b, 0) / prices.length
      
      // Стандартное отклонение
      const variance = prices.reduce((sum, price) => sum + Math.pow(price - avgPrice, 2), 0) / prices.length
      stdDev = Math.sqrt(variance)
    }
  }
  
  // Fallback: данные из itemData (min/max из истории)
  if (!minPrice || !maxPrice || !avgPrice) {
    minPrice = itemData.pricemin || 0
    maxPrice = itemData.pricemax || 0
    avgPrice = itemData.priceavg || itemData.pricelatest || 0
    
    // Приблизительное стандартное отклонение
    if (minPrice > 0 && maxPrice > 0 && avgPrice > 0) {
      stdDev = (maxPrice - minPrice) / 4 // Приблизительная оценка
    } else {
      stdDev = 0
    }
  }
  
  if (!avgPrice || avgPrice <= 0) return 0.5
  
  // Формула: (maxPrice - minPrice) / avgPrice * 0.5 + (stdDev / avgPrice) * 0.5
  const rangeVolatility = (maxPrice - minPrice) / avgPrice
  const stdDevVolatility = stdDev / avgPrice
  
  const volatility = (rangeVolatility * 0.5) + (stdDevVolatility * 0.5)
  
  // ПРЯМАЯ нормализация (высокая волатильность = лучше)
  // Нормализуем в диапазон 0-1, где 1 = высокая волатильность
  return Math.min(1, Math.max(0, analytics_normalizeToRange(volatility, 0, 1, 0, 1)))
}

/**
 * Расчет Hero Trend Score (оценка тренда героя)
 * @param {number} heroId - ID героя
 * @param {string} rankCategory - Категория ранга ('High Rank' или 'All Ranks')
 * @param {Object} heroStats - Данные статистики героя из HeroStats
 * @returns {number} Score от 0 до 1
 */
function analytics_calculateHeroTrendScore(heroId, rankCategory, heroStats) {
  if (!heroId || !heroStats) return 0.5
  
  // Получаем последние данные статистики
  const latestStats = heroStats_getLatestStats(heroId, rankCategory)
  if (!latestStats) return 0.5
  
  // Парсим JSON данные
  let stats
  try {
    stats = typeof latestStats === 'string' ? JSON.parse(latestStats) : latestStats
  } catch (e) {
    return 0.5
  }
  
  // Новые данные (Immortal только, убраны фейки)
  const proContestRateChange7d = stats.proContestRateChange7d || 0
  const pickRateChange7d = stats.pickRateChange7d || 0  // Immortal за неделю
  const pickRatePercent = stats.pickRatePercent || 0  // Текущий пикрейт Immortal
  const winRate = stats.winRate || 0
  
  // Формула с весами из ANALYTICS_WEIGHTS
  const weights = ANALYTICS_WEIGHTS.HERO_TREND_SCORE
  
  // Нормализация компонентов
  const proContestRateChangeNorm = analytics_normalizeToRange(proContestRateChange7d, -0.3, 0.3, 0, 1)
  const pickRateChange7dNorm = analytics_normalizeToRange(pickRateChange7d, -0.3, 0.3, 0, 1)
  const pickRateNorm = analytics_normalizeToRange((pickRatePercent - 50) / 50, -1, 1, 0, 1)
  const winRateNorm = analytics_normalizeToRange((winRate - 50) / 50, -1, 1, 0, 1)
  
  // Если про-статистика недоступна, перераспределяем вес на pickRateChange7d
  let proContestRateWeight = weights.PRO_CONTEST_RATE_CHANGE_7D
  let pickRateChangeWeight = weights.PICK_RATE_CHANGE_IMMORTAL_7D
  
  if (!proContestRateChange7d || proContestRateChange7d === 0) {
    pickRateChangeWeight = weights.PICK_RATE_CHANGE_IMMORTAL_7D + weights.PRO_CONTEST_RATE_CHANGE_7D
    proContestRateWeight = 0
  }
  
  const score = 
    (proContestRateChangeNorm * proContestRateWeight) +
    (pickRateChange7dNorm * pickRateChangeWeight) +
    (pickRateNorm * weights.PICK_RATE_IMMORTAL) +
    (winRateNorm * weights.WIN_RATE)
  
  return Math.min(1, Math.max(0, score))
}

/**
 * Расчет Investment Score (итоговая оценка для покупки)
 * @param {Object} itemData - Данные предмета из SteamWebAPI.ru
 * @param {Object} heroStats - Данные статистики героя
 * @param {Object} historyData - Данные из History
 * @param {string} itemCategory - Категория предмета ('Hero Item' или 'Common Item')
 * @param {number} heroId - ID героя (опционально)
 * @param {string} rankCategory - Категория ранга (опционально)
 * @returns {number} Score от 0 до 100
 */
function analytics_calculateInvestmentScore(itemData, heroStats, historyData, itemCategory, heroId, rankCategory) {
  if (!itemData) return 50
  
  const weights = ANALYTICS_WEIGHTS.INVESTMENT_SCORE
  
  // Расчет всех метрик
  const liquidityScore = analytics_calculateLiquidityScore(itemData)
  const demandRatio = analytics_calculateDemandRatio(itemData)
  const priceMomentum = analytics_calculatePriceMomentum(itemData, historyData)
  const salesTrend = analytics_calculateSalesTrend(itemData)
  const volatilityIndex = analytics_calculateVolatilityIndex(itemData, historyData)
  
  // Hero Trend Score (только для Hero Items)
  let heroTrendScore = 0.5
  if (itemCategory === 'Hero Item' && heroId && rankCategory && heroStats) {
    heroTrendScore = analytics_calculateHeroTrendScore(heroId, rankCategory, heroStats)
  }
  
  // Формула для Hero Items
  if (itemCategory === 'Hero Item') {
    const baseScore = 
      (heroTrendScore * weights.HERO_TREND) +
      (volatilityIndex * weights.VOLATILITY) +
      (demandRatio * weights.DEMAND_RATIO) +
      (priceMomentum * weights.PRICE_MOMENTUM) +
      (liquidityScore * weights.LIQUIDITY) +
      (salesTrend * weights.SALES_TREND)
    
    // Применение корреляций (бонусные множители)
    let finalScore = baseScore
    
    // Корреляция 1: Hero Trend + Price Momentum
    if (heroTrendScore > 0.6 && priceMomentum > 0.5) {
      finalScore = Math.min(1.0, finalScore * 1.2)
    }
    
    // Корреляция 2: Demand Ratio + Liquidity Score
    if (demandRatio > 0.7 && liquidityScore > 0.6) {
      finalScore = Math.min(1.0, finalScore * 1.15)
    }
    
    return Math.round(Math.min(100, Math.max(0, finalScore * 100)))
  } else {
    // Формула для Common Items (перераспределенные веса)
    const score = Math.min(1, Math.max(0,
      (volatilityIndex * 0.30) +
      (demandRatio * 0.25) +
      (priceMomentum * 0.20) +
      (salesTrend * 0.15) +
      (liquidityScore * 0.10)
    ))
    return Math.round(score * 100)
  }
}

/**
 * Расчет Buyback Score (итоговая оценка для откупа)
 * @param {Object} itemData - Данные предмета из SteamWebAPI.ru
 * @param {Object} heroStats - Данные статистики героя
 * @param {Object} historyData - Данные из History
 * @param {number} sellPrice - Цена продажи
 * @param {number} currentPrice - Текущая цена
 * @param {number} heroId - ID героя (опционально)
 * @param {string} rankCategory - Категория ранга (опционально)
 * @returns {number} Score от 0 до 100
 */
function analytics_calculateBuybackScore(itemData, heroStats, historyData, sellPrice, currentPrice, heroId, rankCategory) {
  if (!itemData || !sellPrice || !currentPrice) return 50
  
  const weights = ANALYTICS_WEIGHTS.BUYBACK_SCORE
  
  // Процент просадки цены
  const priceDropPercent = (sellPrice - currentPrice) / sellPrice
  
  // Расчет метрик
  const volatilityIndex = analytics_calculateVolatilityIndex(itemData, historyData)
  const demandRatio = analytics_calculateDemandRatio(itemData)
  const priceMomentum = analytics_calculatePriceMomentum(itemData, historyData)
  const liquidityScore = analytics_calculateLiquidityScore(itemData)
  
  // Hero Trend Score
  let heroTrendScore = 0.5
  if (heroId && rankCategory && heroStats) {
    heroTrendScore = analytics_calculateHeroTrendScore(heroId, rankCategory, heroStats)
  }
  
  // Нормализация процента просадки (чем больше просадка, тем лучше)
  const priceDropNorm = analytics_normalizeToRange(priceDropPercent, 0, 1, 0, 1)
  
  // Формула
  const score = 
    (priceDropNorm * weights.PRICE_DROP) +
    (volatilityIndex * weights.VOLATILITY) +
    (heroTrendScore * weights.HERO_TREND) +
    (demandRatio * weights.DEMAND_RATIO) +
    (priceMomentum * weights.PRICE_MOMENTUM) +
    (liquidityScore * weights.LIQUIDITY)
  
  return Math.round(Math.min(100, Math.max(0, score * 100)))
}

/**
 * Расчет Risk Level (уровень риска)
 * @param {number} investmentScore - Investment Score или Buyback Score (0-100)
 * @param {number} volatilityIndex - Volatility Index (0-1)
 * @param {number} demandRatio - Demand Ratio (0-1)
 * @returns {string} 'Низкий', 'Средний', 'Высокий'
 */
function analytics_calculateRiskLevel(investmentScore, volatilityIndex, demandRatio) {
  if (investmentScore >= 70 && volatilityIndex < 0.5 && demandRatio > 0.6) {
    return 'Низкий'
  } else if (investmentScore >= 50 && volatilityIndex < 0.7 && demandRatio > 0.4) {
    return 'Средний'
  } else {
    return 'Высокий'
  }
}

// === НОРМАЛИЗАЦИЯ ===

/**
 * Нормализация метрики в диапазон 0-1
 * @param {number} value - Значение для нормализации
 * @param {number} min - Минимальное значение
 * @param {number} max - Максимальное значение
 * @param {boolean} inverse - Инвертировать (false = больше = лучше, true = меньше = лучше)
 * @returns {number} Нормализованное значение от 0 до 1
 */
function analytics_normalizeMetric(value, min, max, inverse = false) {
  if (max === min) return 0.5
  
  const normalized = (value - min) / (max - min)
  return inverse ? 1 - normalized : Math.min(1, Math.max(0, normalized))
}

/**
 * Нормализация значения в целевой диапазон
 * @param {number} value - Значение для нормализации
 * @param {number} min - Минимальное значение исходного диапазона
 * @param {number} max - Максимальное значение исходного диапазона
 * @param {number} targetMin - Минимальное значение целевого диапазона (по умолчанию 0)
 * @param {number} targetMax - Максимальное значение целевого диапазона (по умолчанию 1)
 * @returns {number} Нормализованное значение
 */
function analytics_normalizeToRange(value, min, max, targetMin = 0, targetMax = 1) {
  if (max === min) return (targetMin + targetMax) / 2
  
  const normalized = ((value - min) / (max - min)) * (targetMax - targetMin) + targetMin
  return Math.min(targetMax, Math.max(targetMin, normalized))
}

/**
 * Форматирование скора для отображения (🟢 85)
 * Использует круглые эмодзи для единообразия
 * Поддерживает оба формата: 0-1 (автоматически умножает на 100) и 0-100
 * @param {number} score - Score от 0 до 1 или от 0 до 100
 * @returns {string} Отформатированная строка
 */
function analytics_formatScore(score) {
  if (typeof score !== 'number' || isNaN(score)) return '—'
  
  // Автоматическое определение формата: если значение < 1, умножаем на 100
  const normalizedScore = score < 1 ? Math.round(score * 100) : Math.round(score)
  
  // Круглые эмодзи: 🟢 (>=75), 🟡 (>=60), ⚪ (>=40), 🔴 (<40)
  const emoji = normalizedScore >= 75 ? '🟢' : normalizedScore >= 60 ? '🟡' : normalizedScore >= 40 ? '⚪' : '🔴'
  return `${emoji} ${normalizedScore}`
}

/**
 * Расчет Мета сигнала (краткосрочный индикатор для патч-имб)
 * Фокус на краткосрочных изменениях (24h)
 * @param {number} heroId - ID героя
 * @param {string} rankCategory - Категория ранга ('High Rank' или 'All Ranks')
 * @param {Object} heroStats - Данные статистики героя из HeroStats
 * @returns {number} Score от 0 до 100
 */
function analytics_calculateMetaSignal(heroId, rankCategory, heroStats) {
  if (!heroId || !heroStats) return 0
  
  // Получаем последние данные статистики
  const latestStats = heroStats_getLatestStats(heroId, rankCategory)
  if (!latestStats) return 0
  
  // Парсим JSON данные
  let stats
  try {
    stats = typeof latestStats === 'string' ? JSON.parse(latestStats) : latestStats
  } catch (e) {
    return 0
  }
  
  // Данные для Мета сигнала (краткосрочные изменения)
  const pickRateChange24h = stats.pickRateChange24h || 0  // Главный индикатор патч-имб
  const proContestRateChange7d = stats.proContestRateChange7d || 0  // Про-мета
  const pickRateChange7d = stats.pickRateChange7d || 0  // Недельный тренд Immortal
  
  // Формула с весами из ANALYTICS_WEIGHTS
  const weights = ANALYTICS_WEIGHTS.META_SIGNAL
  
  // Нормализация компонентов (24h изменения могут быть более резкими)
  const pickRateChange24hNorm = analytics_normalizeToRange(pickRateChange24h, -0.5, 0.5, 0, 1)
  const proContestRateChangeNorm = analytics_normalizeToRange(proContestRateChange7d, -0.3, 0.3, 0, 1)
  const pickRateChange7dNorm = analytics_normalizeToRange(pickRateChange7d, -0.3, 0.3, 0, 1)
  
  // Если 24h данные недоступны, перераспределяем вес на 7d
  let pickRateChange24hWeight = weights.PICK_RATE_CHANGE_IMMORTAL_24H
  let pickRateChange7dWeight = weights.PICK_RATE_CHANGE_IMMORTAL_7D
  
  if (!pickRateChange24h || pickRateChange24h === 0) {
    pickRateChange7dWeight = weights.PICK_RATE_CHANGE_IMMORTAL_7D + weights.PICK_RATE_CHANGE_IMMORTAL_24H
    pickRateChange24hWeight = 0
  }
  
  const score = 
    (pickRateChange24hNorm * pickRateChange24hWeight) +
    (proContestRateChangeNorm * weights.PRO_CONTEST_RATE_CHANGE_7D) +
    (pickRateChange7dNorm * pickRateChange7dWeight)
  
  // Возвращаем в диапазоне 0-100
  return Math.round(Math.min(100, Math.max(0, score * 100)))
}

/**
 * Форматирование Мета сигнала для отображения (🔥 92)
 * Использует специальные эмодзи для краткосрочных сигналов
 * @param {number} score - Score от 0 до 100
 * @returns {string} Отформатированная строка
 */
function analytics_formatMetaSignal(score) {
  if (typeof score !== 'number' || isNaN(score)) return '—'
  
  const normalizedScore = Math.round(score)
  
  // Специальные эмодзи для Мета сигнала: 🔥 (>=75), 🟡 (>=60), ⚪ (>=40), 🔴 (<40)
  const emoji = normalizedScore >= 75 ? '🔥' : normalizedScore >= 60 ? '🟡' : normalizedScore >= 40 ? '⚪' : '🔴'
  return `${emoji} ${normalizedScore}`
}

