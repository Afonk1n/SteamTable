#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для сравнения статистики из таблицы с расчетами по формулам
"""

import re
from datetime import datetime
from typing import List, Tuple, Dict

# Константы из Constants.gs
TREND_ANALYSIS_CONFIG = {
    'SIMPLE_COMPARISON': {
        'BASE_THRESHOLD': 0.08,
        'VOLATILITY_MULTIPLIER': 1.5,
        'SIDEWAYS_FACTOR': 0.4
    },
    'MOVING_AVERAGES': {
        'SHORT_WINDOW': 3,
        'LONG_WINDOW': 7,
        'BASE_THRESHOLD': 0.02,
        'VOLATILITY_MULTIPLIER': 1.5
    },
    'LINEAR_REGRESSION': {
        'WINDOW': 20,
        'GROWTH_THRESHOLD': 0.03,
        'FALL_THRESHOLD': -0.03
    },
    'MOMENTUM': {
        'WINDOW': 5,
        'BASE_THRESHOLD': 0.05,
        'VOLATILITY_MULTIPLIER': 2
    }
}

def parse_price(value_str: str) -> float:
    """Парсит цену из строки вида '20,61 ₽' или '20.61'"""
    if not value_str or value_str.strip() == '':
        return None
    # Убираем валюту и пробелы, заменяем запятую на точку
    cleaned = value_str.replace('₽', '').replace(' ', '').replace(',', '.')
    try:
        return float(cleaned)
    except:
        return None

def parse_date(date_str: str) -> datetime:
    """Парсит дату из формата dd.MM.yy"""
    try:
        return datetime.strptime(date_str, '%d.%m.%y')
    except:
        return None

def calculate_volatility(prices: List[float]) -> float:
    """Вычисляет волатильность"""
    if len(prices) < 2:
        return 0
    sum_squared_changes = 0
    for i in range(1, len(prices)):
        if prices[i-1] > 0:
            change = abs((prices[i] - prices[i-1]) / prices[i-1])
            sum_squared_changes += change * change
    return (sum_squared_changes / (len(prices) - 1)) ** 0.5

def simple_comparison(prices: List[float]) -> str:
    """Метод 1: Простое сравнение последних значений"""
    if len(prices) < 2:
        return '🟪'
    
    config = TREND_ANALYSIS_CONFIG['SIMPLE_COMPARISON']
    recent = prices[-3:] if len(prices) >= 3 else prices
    first = recent[0]
    last = recent[-1]
    
    volatility = calculate_volatility(recent)
    adaptive_threshold = config['BASE_THRESHOLD'] + (volatility * config['VOLATILITY_MULTIPLIER'])
    
    change = abs((last - first) / first) if first > 0 else 0
    
    if change < adaptive_threshold * config['SIDEWAYS_FACTOR']:
        return '🟨'
    return '🟩' if (last > first and change > adaptive_threshold) else ('🟥' if change > adaptive_threshold else '🟨')

def moving_averages(prices: List[float]) -> str:
    """Метод 2: Скользящие средние"""
    if len(prices) < 4:
        return '🟪'
    
    config = TREND_ANALYSIS_CONFIG['MOVING_AVERAGES']
    short_window = min(config['SHORT_WINDOW'], len(prices) // 2)
    long_window = min(config['LONG_WINDOW'], len(prices))
    
    short_ma = sum(prices[-short_window:]) / short_window
    long_ma = sum(prices[-long_window:]) / long_window
    
    if long_ma == 0:
        return '🟪'
    
    diff = (short_ma - long_ma) / long_ma
    volatility = calculate_volatility(prices)
    adaptive_threshold = config['BASE_THRESHOLD'] + (volatility * config['VOLATILITY_MULTIPLIER'])
    
    if diff > adaptive_threshold:
        return '🟩'
    if diff < -adaptive_threshold:
        return '🟥'
    return '🟨'

def linear_regression(prices: List[float]) -> str:
    """Метод 3: Линейная регрессия"""
    if len(prices) < 3:
        return '🟪'
    
    config = TREND_ANALYSIS_CONFIG['LINEAR_REGRESSION']
    window = min(config['WINDOW'], len(prices))
    recent_prices = prices[-window:]
    
    n = len(recent_prices)
    x = list(range(n))
    y = recent_prices
    
    sum_x = sum(x)
    sum_y = sum(y)
    sum_xy = sum(x[i] * y[i] for i in range(n))
    sum_xx = sum(xi * xi for xi in x)
    
    denominator = n * sum_xx - sum_x * sum_x
    if denominator == 0:
        return '🟪'
    
    slope = (n * sum_xy - sum_x * sum_y) / denominator
    avg_price = sum_y / n
    
    if avg_price == 0:
        return '🟪'
    
    relative_slope = slope / avg_price
    
    if relative_slope > config['GROWTH_THRESHOLD']:
        return '🟩'
    if relative_slope < config['FALL_THRESHOLD']:
        return '🟥'
    return '🟨'

def momentum_analysis(prices: List[float]) -> str:
    """Метод 4: Momentum анализ"""
    if len(prices) < 3:
        return '🟪'
    
    config = TREND_ANALYSIS_CONFIG['MOMENTUM']
    recent = prices[-min(config['WINDOW'], len(prices)):]
    momentum = recent[-1] - recent[0]
    avg_price = sum(recent) / len(recent)
    
    if avg_price == 0:
        return '🟪'
    
    momentum_percent = momentum / avg_price
    volatility = calculate_volatility(recent)
    adaptive_threshold = config['BASE_THRESHOLD'] + (volatility * config['VOLATILITY_MULTIPLIER'])
    
    if momentum_percent > adaptive_threshold:
        return '🟩'
    if momentum_percent < -adaptive_threshold:
        return '🟥'
    return '🟨'

def analyze_trend(prices: List[float], dates: List[datetime]) -> Tuple[str, int]:
    """Анализирует тренд и возвращает (тренд, дни_смены)"""
    if len(prices) < 2:
        return ('🟪', 0)
    
    methods = [
        simple_comparison(prices),
        moving_averages(prices),
        linear_regression(prices),
        momentum_analysis(prices)
    ]
    
    # Голосование
    votes = {'🟩': 0, '🟥': 0, '🟨': 0, '🟪': 0}
    for method in methods:
        if method in votes:
            votes[method] += 1
    
    # Определяем победителя с приоритетами
    priority_order = ['🟥', '🟩', '🟨', '🟪']
    trend = '🟪'
    max_votes = 0
    
    for trend_type in priority_order:
        if votes[trend_type] > max_votes:
            max_votes = votes[trend_type]
            trend = trend_type
    
    # Расчет дней смены с использованием реальных дат
    days_change = calculate_days_change(prices, dates, trend)
    
    return (trend, days_change)

def calculate_days_change(prices: List[float], dates: List[datetime], current_trend: str) -> int:
    """Вычисляет количество дней с последней смены тренда"""
    if len(prices) < 3 or len(dates) < 2:
        return 0
    
    # Анализируем тренды для каждого периода, начиная с предпоследнего
    for i in range(len(prices) - 2, 0, -1):
        period_prices = prices[:i+1]
        period_trend = simple_comparison(period_prices)
        
        if period_trend != current_trend:
            # Нашли смену тренда
            change_date = dates[i]
            current_date = dates[-1]
            
            if change_date and current_date:
                diff_time = abs((current_date - change_date).total_seconds())
                days_diff = int(diff_time / (24 * 60 * 60))
                return days_diff if days_diff > 0 else 1
    
    # Если тренд не менялся, вычисляем разницу между первой и последней датой
    if len(dates) >= 2:
        first_date = dates[0]
        last_date = dates[-1]
        if first_date and last_date:
            diff_time = abs((last_date - first_date).total_seconds())
            days_diff = int(diff_time / (24 * 60 * 60))
            return days_diff if days_diff > 0 else 1
    
    return 0

def parse_history_line(line: str, headers: List[str]) -> Dict:
    """Парсит строку из History.md"""
    # Разделяем по табуляции
    parts = line.split('\t')
    
    if len(parts) < 14:
        return None
    
    # Структура: [0]пусто, [1]название, [2]статус, [3]ссылка, [4]купить, [5]текущая_цена, 
    # [6]min, [7]max, [8]тренд, [9]дни_смены, [10]фаза, [11]потенциал, [12]рекомендация, [13+]даты
    name = parts[1].strip() if len(parts) > 1 else ''
    if not name:
        return None
    
    # Структура колонок (проверено):
    # [0]пусто, [1]название, [2]статус, [3]ссылка, [4]купить, [5]текущая_цена, 
    # [6]min, [7]max, [8]тренд, [9]дни_смены, [10]фаза, [11]потенциал, [12]рекомендация, [13+]даты
    current_price = parse_price(parts[5]) if len(parts) > 5 else None
    min_price = parse_price(parts[6]) if len(parts) > 6 else None
    max_price = parse_price(parts[7]) if len(parts) > 7 else None
    table_trend = parts[8].strip() if len(parts) > 8 else ''
    table_days = parts[9].strip() if len(parts) > 9 else ''
    phase = parts[10].strip() if len(parts) > 10 else ''
    potential = parts[11].strip() if len(parts) > 11 else ''
    recommendation = parts[12].strip() if len(parts) > 12 else ''
    
    # Парсим цены по датам (начиная с колонки 13)
    # Важно: группируем по дате, беря последнюю цену за день (если есть "ночь" и "день")
    prices_by_date = {}
    date_entries = []  # Храним все записи для правильной группировки
    
    # Собираем все записи с их позициями
    for i in range(13, min(len(headers), len(parts))):
        if i >= len(headers):
            break
        header = headers[i].strip()
        price_str = parts[i].strip() if i < len(parts) else ''
        
        # Извлекаем дату из заголовка
        date_match = re.match(r'^(\d{2}\.\d{2}\.\d{2})', header)
        if date_match:
            date_key = date_match.group(1)
            price = parse_price(price_str)
            if price is not None:
                is_day = 'день' in header
                is_night = 'ночь' in header
                date_entries.append({
                    'date_key': date_key,
                    'price': price,
                    'col_index': i,
                    'is_day': is_day,
                    'is_night': is_night
                })
    
    # Сортируем по позиции колонки (слева направо = старые -> новые)
    date_entries.sort(key=lambda x: x['col_index'])
    
    # Группируем по дате, беря последнюю цену за день
    date_headers = []
    for entry in date_entries:
        date_key = entry['date_key']
        prices_by_date[date_key] = entry['price']  # Перезаписываем, беря последнюю
        if date_key not in date_headers:
            date_headers.append(date_key)
    
    # Сортируем даты в хронологическом порядке
    def sort_dates(date_str):
        parts = date_str.split('.')
        if len(parts) == 3:
            return (2000 + int(parts[2]), int(parts[1]), int(parts[0]))
        return (0, 0, 0)
    
    sorted_dates = sorted(date_headers, key=sort_dates)
    prices = [prices_by_date[date] for date in sorted_dates if date in prices_by_date]
    
    return {
        'name': name,
        'current_price': current_price,
        'min_price': min_price,
        'max_price': max_price,
        'table_trend': table_trend,
        'table_days': table_days,
        'phase': phase,
        'potential': potential,
        'recommendation': recommendation,
        'prices': prices,
        'dates': sorted_dates
    }

def main():
    """Основная функция"""
    print("Чтение данных из History.md...")
    
    try:
        with open('History.md', 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except Exception as e:
        print(f"Ошибка чтения файла: {e}")
        return
    
    # Парсим заголовок
    header_line = lines[0].strip()
    headers = header_line.split('\t')
    
    # Парсим данные (первые 10 строк для примера)
    results = []
    for i in range(1, min(11, len(lines))):
        line = lines[i].strip()
        if not line:
            continue
        
        data = parse_history_line(line, headers)
        if not data or not data['prices'] or len(data['prices']) < 2:
            continue
        
        # Преобразуем даты в объекты datetime
        dates = [parse_date(date_key) for date_key in data['dates']]
        dates = [d for d in dates if d is not None]  # Убираем None
        
        # Рассчитываем статистику
        calculated_trend, calculated_days = analyze_trend(data['prices'], dates)
        
        # Парсим дни из таблицы
        try:
            # Убираем все нецифровые символы
            days_clean = re.sub(r'[^\d]', '', data['table_days'])
            table_days_int = int(days_clean) if days_clean else 0
        except:
            table_days_int = 0
        
        # Отладочная информация (закомментировано для чистоты вывода)
        # print(f"DEBUG: {data['name']}")
        # print(f"  table_trend: {repr(data['table_trend'])}")
        # print(f"  calculated_trend: {repr(calculated_trend)}")
        # print(f"  table_days: {repr(data['table_days'])} -> {table_days_int}")
        # print(f"  calculated_days: {calculated_days}")
        
        results.append({
            'name': data['name'],
            'prices_count': len(data['prices']),
            'dates_count': len(data['dates']),
            'table_trend': data['table_trend'],
            'calculated_trend': calculated_trend,
            'trend_match': data['table_trend'] == calculated_trend,
            'table_days': table_days_int,
            'calculated_days': calculated_days,
            'days_diff': abs(table_days_int - calculated_days),
            'prices': data['prices'][:5] + ['...'] + data['prices'][-5:] if len(data['prices']) > 10 else data['prices'],
            'dates': data['dates'][:5] + ['...'] + data['dates'][-5:] if len(data['dates']) > 10 else data['dates']
        })
    
    # Выводим результаты
    output = []
    output.append("=" * 100)
    output.append("СРАВНЕНИЕ СТАТИСТИКИ: ТАБЛИЦА vs РАСЧЕТ")
    output.append("=" * 100)
    output.append("")
    
    for i, result in enumerate(results, 1):
        output.append(f"\n{i}. {result['name']}")
        output.append("-" * 100)
        output.append(f"Количество цен (после группировки по дате): {result['prices_count']}")
        output.append(f"Количество уникальных дат: {result['dates_count']}")
        output.append(f"Период данных: {result['dates'][0]} - {result['dates'][-1]}")
        output.append(f"Первые 5 цен: {result['prices'][:5]}")
        output.append(f"Последние 5 цен: {result['prices'][-5:]}")
        output.append("")
        output.append("ТРЕНД:")
        output.append(f"  Таблица:     {result['table_trend']} (repr: {repr(result['table_trend'])})")
        output.append(f"  Расчет:      {result['calculated_trend']}")
        output.append(f"  Совпадение:  {'✓' if result['trend_match'] else '✗'}")
        if not result['trend_match']:
            output.append(f"  ПРИМЕЧАНИЕ: Тренды не совпадают!")
        output.append("")
        output.append("ДНИ СМЕНЫ:")
        output.append(f"  Таблица:     {result['table_days']}")
        output.append(f"  Расчет:      {result['calculated_days']}")
        output.append(f"  Разница:     {result['days_diff']} дней")
        if result['days_diff'] > 5:
            output.append(f"  ПРИМЕЧАНИЕ: Большая разница в днях!")
        output.append("")
    
    if not results:
        output.append("\nОШИБКА: Не удалось распарсить данные из History.md")
        output.append("Проверьте формат файла (должен быть разделен табуляцией)")
    else:
        # Сводка
        output.append("\n" + "=" * 100)
        output.append("СВОДКА")
        output.append("=" * 100)
        trend_matches = sum(1 for r in results if r['trend_match'])
        avg_days_diff = sum(r['days_diff'] for r in results) / len(results) if results else 0
        output.append(f"Совпадение трендов: {trend_matches}/{len(results)} ({trend_matches*100/len(results):.1f}%)")
        output.append(f"Средняя разница в днях: {avg_days_diff:.1f}")
        output.append("")
    
    # Сохраняем в файл
    with open('StatisticsComparison.md', 'w', encoding='utf-8') as f:
        f.write('\n'.join(output))
    
    print(f"\nРезультаты сохранены в StatisticsComparison.md")
    print(f"Обработано строк: {len(results)}")
    if results:
        trend_matches = sum(1 for r in results if r['trend_match'])
        print(f"Совпадение трендов: {trend_matches}/{len(results)}")

if __name__ == '__main__':
    main()

