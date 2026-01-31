import fitz
import re

# Те же паттерны, что в main.py
FORMULA_PATTERNS = [
    (re.compile(r'[∑∫∏√±×÷≤≥≠∞∂∇]'), "Математические символы"),
    (re.compile(r'[α-ωΑ-Ω]'), "Греческие буквы"),
    (re.compile(r'\$[a-zA-Z0-9+\-*/=\(\)\s]{2,}\$'), "LaTeX inline"),
    (re.compile(r'\\begin{equation}'), "LaTeX блочные"),
    (re.compile(r'[a-zA-Z]+\([a-zA-Z0-9,\s]+\)'), "Функции f(t), u(x,t)"),
    (re.compile(r'[a-zA-Z]\([a-zA-Z0-9,\s]*\)'), "Простые функции f(x)"),
    (re.compile(r'\b[a-zA-Z]{1,3}_[a-zA-Z0-9]{1,3}\b'), "Индексы x_1, y_max"),
    (re.compile(r'\^[0-9]+|\^{[^}]+}'), "Степени x^2"),
    (re.compile(r'рН\s*=\s*[0-9]+([,\.][0-9]+)?(\s*[–-]\s*[0-9]+([,\.][0-9]+)?)?'), "pH формулы"),
]

# ФИЛЬТРЫ: что НЕ является формулой
EXCLUDE_PATTERNS = [
    (re.compile(r'\b[A-Z]\(\d+\)'), "Ссылка на литературу"),  # R(3), A(1)
    (re.compile(r'\b[A-Z]{2,}[-–][A-Z0-9][-–A-Z0-9]+'), "Артикул/код"),  # RU-DPP-3, ABC-123
    (re.compile(r'[±]\s*\d+\s*%'), "Погрешность в %"),  # ±21%
    (re.compile(r'\d+\s*[×]\s*\d+(\s*[×]\s*\d+)*\s*(мм|см|м|mm|cm|m)\b'), "Размеры"),  # 50×50×50 мм
    (re.compile(r'\d+\s*[а-яА-Яa-zA-Z]+[·•]\s*[а-яА-Яa-zA-Z]+'), "Единицы измерения"),  # мА·ч, кВт·ч
]

def is_false_positive(text_match):
    """Проверяет, является ли найденное совпадение ложным срабатыванием"""
    for pattern, reason in EXCLUDE_PATTERNS:
        if pattern.search(str(text_match)):
            return True, reason
    return False, None

def analyze_page_formulas(pdf_path, max_pages=10):
    """Анализирует первые max_pages страниц и показывает, какие паттерны срабатывают"""
    doc = fitz.open(pdf_path)
    
    print(f"📄 Анализ файла: {pdf_path}")
    print(f"📊 Всего страниц: {len(doc)}")
    print(f"🔍 Проверяю первые {max_pages} страниц с формулами...\n")
    
    pages_with_formulas = 0
    pages_with_real_formulas = 0
    
    # Паттерны для определения реальных формул
    has_latex = re.compile(r'\$[^$]+\$|\\begin\{equation\}')
    has_chemical = re.compile(r'[A-Z][a-z]?\([A-Z0-9a-z]+\)')  # Si(CH3), Ru(dpp)
    has_equation = re.compile(r'[a-zA-Zα-ωΑ-Ω]\s*=\s*[^0-9]')  # переменная = не-число
    
    # Паттерны единиц измерения
    units_pattern = re.compile(r'(мм|см|км|нм|мкм|%|°C|K|кг|мг|г|ч|мин|с|Вт|А|В|Ом|Гц|Па|Дж|ppm|mm|cm|m|nm|Hz|Pa|mol)\b', re.IGNORECASE)
    
    # Греческие буквы которые НЕ формулы (используются в единицах)
    greek_in_units = ['μ', 'Ω']  # микро, ом
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text()
        
        # Проверяем каждый паттерн
        matches = []
        filtered_out = []
        
        for pattern, name in FORMULA_PATTERNS:
            found = pattern.findall(text)
            if found:
                # Фильтруем ложные срабатывания
                valid = []
                for match in found[:10]:
                    is_false, reason = is_false_positive(match)
                    if is_false:
                        filtered_out.append((match, reason))
                    else:
                        valid.append(match)
                
                if valid:
                    matches.append((name, valid[:5]))
        
        # === КОМБИНИРОВАННАЯ ПРОВЕРКА ===
        has_real_formulas = False
        reason = ""
        
        if matches:
            # ✅ КРИТЕРИЙ 1: Есть LaTeX
            if has_latex.search(text):
                has_real_formulas = True
                reason = "LaTeX формулы"
            
            # ✅ КРИТЕРИЙ 2: Есть химические формулы
            elif has_chemical.search(text):
                has_real_formulas = True
                reason = "Химические формулы"
            
            # ✅ КРИТЕРИЙ 3: Есть уравнения (знак = с переменной)
            elif has_equation.search(text):
                has_real_formulas = True
                reason = "Уравнения с переменными"
            
            # ✅ КРИТЕРИЙ 4: Греческие буквы НЕ в единицах измерения
            elif any(name == "Греческие буквы" for name, _ in matches):
                # Проверяем, что греческие буквы не только в единицах (μ, Ω)
                greek_matches = [ex for name, examples in matches if name == "Греческие буквы" for ex in examples]
                real_greek = [g for g in greek_matches if g not in greek_in_units]
                
                if real_greek:
                    has_real_formulas = True
                    reason = f"Греческие буквы в формулах: {real_greek[:3]}"
            
            # ❌ ИНАЧЕ: Проверяем, не только ли это технические параметры
            if not has_real_formulas:
                # Если нашли только "Математические символы" (±, ×, ≈)
                # И в тексте много единиц измерения
                math_symbols_only = all(name == "Математические символы" for name, _ in matches)
                
                if math_symbols_only:
                    # Считаем единицы измерения
                    units_count = len(units_pattern.findall(text))
                    symbols_count = sum(len(examples) for _, examples in matches)
                    
                    # Если единиц больше или примерно столько же, сколько символов
                    # → это технические параметры
                    if units_count >= symbols_count * 0.5:
                        reason = f"Технические параметры ({symbols_count} символов, {units_count} единиц)"
                    else:
                        # Иначе оставляем как формулы
                        has_real_formulas = True
                        reason = "Математические символы без единиц"
                else:
                    # Если есть другие паттерны (функции, индексы) - вероятно формулы
                    has_real_formulas = True
                    reason = f"Смешанный контент: {[n for n, _ in matches]}"
        
        had_any_matches = len(matches) > 0 or len(filtered_out) > 0
        
        if had_any_matches:
            pages_with_formulas += 1
        
        if has_real_formulas:
            pages_with_real_formulas += 1
            
            if pages_with_real_formulas <= max_pages:
                print(f"═══ Страница {page_num + 1} ═══")
                print(f"✅ ФОРМУЛА: {reason}")
                
                if filtered_out:
                    print(f"\n🚫 Отфильтровано:")
                    for match, filter_reason in filtered_out[:3]:
                        print(f"    - {repr(match)} → {filter_reason}")
                
                print(f"\n📊 Паттерны ({len(matches)} типов):")
                for pattern_name, examples in matches:
                    print(f"  ✓ {pattern_name}: {examples[:3]}")
                
                text_preview = text[:250].replace('\n', ' ')
                print(f"\n  📝 {text_preview}...")
                print()
        
        elif had_any_matches and pages_with_formulas - pages_with_real_formulas <= 5:
            # Показываем первые 5 отфильтрованных страниц
            print(f"═══ Страница {page_num + 1} (ОТФИЛЬТРОВАНА) ═══")
            print(f"❌ НЕ формула: {reason if reason else 'Не прошла критерии'}")
            
            for pattern_name, examples in matches:
                print(f"  • {pattern_name}: {examples[:2]}")
            
            text_preview = text[:200].replace('\n', ' ')
            print(f"  📝 {text_preview}...")
            print()
    
    print(f"\n📈 ИТОГО:")
    print(f"  - Страниц со срабатываниями (ДО фильтрации): {pages_with_formulas}")
    print(f"  - Страниц с РЕАЛЬНЫМИ формулами (ПОСЛЕ фильтрации): {pages_with_real_formulas}")
    print(f"  - Отфильтровано: {pages_with_formulas - pages_with_real_formulas}")
    print(f"  - Всего страниц в документе: {len(doc)}")
    doc.close()

if __name__ == "__main__":
    # Укажите имя файла
    analyze_page_formulas("01.pdf", max_pages=15)
