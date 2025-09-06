import fitz  # PyMuPDF
import pdfplumber
import re
import pandas as pd
import os
import io
import numpy as np
from PIL import Image
from pix2text import Pix2Text
import tempfile
import docx2txt
from docx import Document
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph
import glob
from datetime import datetime, timedelta
import subprocess
import logging
import gc
import threading

# --- Константы и Настройки ---

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('smeta_mu.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)

# Глобальная ML-модель (инициализируется один раз)
_global_p2t_model = None
_model_lock = threading.Lock()

def get_pix2text_model():
    """Получить глобальный экземпляр Pix2Text модели (thread-safe)."""
    global _global_p2t_model
    
    if _global_p2t_model is None:
        with _model_lock:
            if _global_p2t_model is None:  # Double-check locking
                try:
                    logger.info("Инициализация глобальной модели Pix2Text...")
                    _global_p2t_model = Pix2Text()
                    logger.info("Модель Pix2Text успешно инициализирована")
                except Exception as e:
                    logger.error(f"Ошибка инициализации Pix2Text: {e}")
                    _global_p2t_model = None
                    raise
    
    return _global_p2t_model

# Расценки в рублях
PRICING = {
    4: 34,  # Формулы
    3: 25,  # Иллюстрации или таблицы
    2: 21,  # Иероглифы или арабская вязь
    1: 9,   # Простой текст
}

# Паттерны для поиска (можно расширять)
FORMULA_PATTERNS = [
    re.compile(r'[∑∫∏√±×÷≤≥≠∞∂∇]'),  # Математические символы
    re.compile(r'[α-ωΑ-Ω]'),             # Греческие буквы
    re.compile(r'\$[a-zA-Z0-9+\-*/=\(\)\s]{2,}\$'),  # LaTeX inline-формулы (минимум 2 символа, только формула-подобные)
    re.compile(r'\\begin{equation}'),   # LaTeX блочные формулы
    re.compile(r'[a-zA-Z]+\([a-zA-Z0-9,\s]+\)'),  # Функции типа f(t), u(x,t), sin(x)
    re.compile(r'[a-zA-Z]\([a-zA-Z0-9,\s]*\)'),   # Простые функции f(x), g(t)
    re.compile(r'\b[a-zA-Z]{1,3}_[a-zA-Z0-9]{1,3}\b'),  # Индексы типа x_1, y_max (короткие)
    re.compile(r'\^[0-9]+|\^{[^}]+}'),            # Степени x^2, x^{n+1}
    # Только pH паттерны (убираем остальные химические)
    re.compile(r'рН\s*=\s*[0-9]+([,\.][0-9]+)?(\s*[–-]\s*[0-9]+([,\.][0-9]+)?)?'),  # рН=6–7, рН=0,5–2
    re.compile(r'pH\s*=\s*[0-9]+([,\.][0-9]+)?(\s*[–-]\s*[0-9]+([,\.][0-9]+)?)?'),  # pH=6–7, pH=0.5–2
]

# Unicode-диапазоны для иероглифов (CJK) и арабской вязи
SPECIAL_CHARS_PATTERN = re.compile(r'[\u4e00-\u9fff\u0600-\u06ff]+')


def convert_doc_to_pdf(doc_path: str) -> str:
    """
    Конвертирует DOC/DOCX файл в PDF и возвращает путь к временному PDF файлу.
    Использует метаданные для правильного подсчета страниц.
    """
    print(f"Конвертирую {doc_path} в PDF...")
    
    # Создаем временный PDF файл
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_pdf_path = temp_pdf.name
    temp_pdf.close()
    
    try:
        # Определяем тип файла
        file_ext = os.path.splitext(doc_path)[1].lower()
        
        if file_ext == '.docx':
            # Для DOCX файлов используем python-docx
            try:
                doc = Document(doc_path)
                text_content = []
                
                for paragraph in doc.paragraphs:
                    if paragraph.text.strip():
                        text_content.append(paragraph.text)
                
                # Создаем PDF с помощью ReportLab
                doc_pdf = SimpleDocTemplate(temp_pdf_path, pagesize=A4)
                styles = getSampleStyleSheet()
                story = []
                
                for text in text_content:
                    para = Paragraph(text, styles['Normal'])
                    story.append(para)
                
                doc_pdf.build(story)
                print(f"✅ DOCX успешно конвертирован в PDF: {temp_pdf_path}")
                
            except Exception as e:
                print(f"Ошибка при обработке DOCX: {e}")
                # Fallback - используем docx2txt
                text = docx2txt.process(doc_path)
                create_pdf_from_text(text, temp_pdf_path)
                
        elif file_ext == '.doc':
            # Для старых DOC файлов извлекаем метаданные
            print("📊 Анализирую метаданные DOC файла...")
            metadata = get_doc_metadata(doc_path)
            num_pages = metadata['pages']
            num_words = metadata['words']
            
            print(f"📄 Найдено страниц: {num_pages}")
            print(f"📝 Найдено слов: {num_words}")
            
            try:
                # Пытаемся извлечь текст
                text = docx2txt.process(doc_path)
                if text.strip():
                    # Создаем PDF с правильным количеством страниц
                    create_multi_page_pdf(text, temp_pdf_path, num_pages)
                    print(f"✅ DOC конвертирован в {num_pages}-страничный PDF: {temp_pdf_path}")
                else:
                    # Если текст не извлекается, создаем заглушку с правильным количеством страниц
                    fallback_text = f"""DOC файл: {os.path.basename(doc_path)}

Содержимое файла не может быть полностью извлечено из-за ограничений формата DOC.

Информация из метаданных:
- Страниц: {num_pages}
- Слов: {num_words}
- Символов: {metadata['characters']}

Для корректной обработки рекомендуется пересохранить файл в формате DOCX."""
                    
                    create_multi_page_pdf(fallback_text, temp_pdf_path, num_pages)
                    print(f"⚠️ Создан {num_pages}-страничный PDF-заглушка")
                    
            except Exception as e:
                print(f"⚠️ Предупреждение при конвертации DOC файла: {e}")
                # Создаем заглушку с правильным количеством страниц
                error_text = f"""Ошибка конвертации DOC файла: {str(e)}

Файл: {os.path.basename(doc_path)}
Страниц по метаданным: {num_pages}
Слов по метаданным: {num_words}

Рекомендуется пересохранить файл в формате DOCX для корректной обработки."""
                
                create_multi_page_pdf(error_text, temp_pdf_path, num_pages)
                print(f"⚠️ Создан {num_pages}-страничный PDF-заглушка с информацией об ошибке")
                
        return temp_pdf_path
        
    except Exception as e:
        # Удаляем временный файл при ошибке
        if os.path.exists(temp_pdf_path):
            os.unlink(temp_pdf_path)
        raise Exception(f"Не удалось конвертировать {doc_path}: {e}")


def create_pdf_from_text(text: str, output_path: str):
    """
    Создает PDF файл из текстового содержимого.
    """
    try:
        doc = SimpleDocTemplate(output_path, pagesize=A4)
        styles = getSampleStyleSheet()
        story = []
        
        # Разбиваем текст на абзацы
        paragraphs = text.split('\n')
        
        for para_text in paragraphs:
            if para_text.strip():
                # Экранируем HTML символы для ReportLab
                para_text = para_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                para = Paragraph(para_text, styles['Normal'])
                story.append(para)
        
        if story:  # Если есть содержимое
            doc.build(story)
        else:
            # Создаем пустой PDF если нет текста
            c = canvas.Canvas(output_path, pagesize=A4)
            c.drawString(100, 750, "Документ не содержит текста")
            c.save()
            
    except Exception as e:
        print(f"Ошибка при создании PDF: {e}")
        # Создаем минимальный PDF
        c = canvas.Canvas(output_path, pagesize=A4)
        c.drawString(100, 750, f"Ошибка конвертации: {str(e)[:100]}")
        c.save()


def get_doc_metadata(doc_path: str) -> dict:
    """Извлекает метаданные из DOC файла используя команду file."""
    try:
        result = subprocess.run(['file', doc_path], capture_output=True, text=True)
        output = result.stdout
        
        metadata = {
            'pages': 1,  # По умолчанию
            'words': 0,
            'characters': 0
        }
        
        # Парсим вывод команды file
        if 'Number of Pages:' in output:
            pages_match = re.search(r'Number of Pages:\s*(\d+)', output)
            if pages_match:
                metadata['pages'] = int(pages_match.group(1))
        
        if 'Number of Words:' in output:
            words_match = re.search(r'Number of Words:\s*(\d+)', output)
            if words_match:
                metadata['words'] = int(words_match.group(1))
                
        if 'Number of Characters:' in output:
            chars_match = re.search(r'Number of Characters:\s*(\d+)', output)
            if chars_match:
                metadata['characters'] = int(chars_match.group(1))
        
        return metadata
        
    except Exception as e:
        print(f"Предупреждение: не удалось получить метаданные из {doc_path}: {e}")
        return {'pages': 1, 'words': 0, 'characters': 0}


def create_multi_page_pdf(text: str, output_path: str, num_pages: int = 1):
    """Создает PDF файл с указанным количеством страниц."""
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        
        c = canvas.Canvas(output_path, pagesize=A4)
        width, height = A4
        
        # Разбиваем текст на части для разных страниц
        lines_per_page = 50
        words = text.split()
        words_per_page = max(1, len(words) // num_pages) if num_pages > 1 else len(words)
        
        for page_num in range(num_pages):
            y_position = height - 50
            
            # Заголовок страницы
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y_position, f"Страница {page_num + 1} из {num_pages}")
            y_position -= 30
            
            # Основной текст
            c.setFont("Helvetica", 10)
            
            # Берем часть слов для этой страницы
            start_word = page_num * words_per_page
            end_word = min((page_num + 1) * words_per_page, len(words))
            page_words = words[start_word:end_word]
            
            if page_words:
                page_text = ' '.join(page_words)
                # Разбиваем текст на строки по 80 символов
                lines = []
                while len(page_text) > 80:
                    split_pos = page_text.rfind(' ', 0, 80)
                    if split_pos == -1:
                        split_pos = 80
                    lines.append(page_text[:split_pos])
                    page_text = page_text[split_pos:].strip()
                
                if page_text:
                    lines.append(page_text)
                
                # Выводим строки на страницу
                for line in lines:
                    if y_position > 50:  # Оставляем место для нижнего поля
                        c.drawString(50, y_position, line[:100])  # Ограничиваем длину строки
                        y_position -= 15
                    else:
                        break
            else:
                # Если нет текста, показываем информацию о странице
                c.drawString(50, y_position, f"Содержимое страницы {page_num + 1}")
                c.drawString(50, y_position - 20, "Текст DOC файла не может быть извлечен корректно")
            
            if page_num < num_pages - 1:  # Не добавляем новую страницу для последней
                c.showPage()
        
        c.save()
        print(f"✅ Создан PDF с {num_pages} страницами: {output_path}")
        
    except Exception as e:
        print(f"Ошибка при создании многостраничного PDF: {e}")
        # Fallback - создаем простой PDF
        create_pdf_from_text(text, output_path)


def find_supported_files(directory: str) -> list:
    """Находит все поддерживаемые файлы в указанной папке."""
    supported_extensions = ['*.pdf', '*.doc', '*.docx']
    found_files = []
    
    for ext in supported_extensions:
        pattern = os.path.join(directory, ext)
        found_files.extend(glob.glob(pattern))
    
    return sorted(found_files)


def get_directory_from_user() -> str:
    """Интерактивно запрашивает у пользователя путь к папке."""
    max_attempts = 3
    attempts = 0
    
    while attempts < max_attempts:
        print("\n=== СИСТЕМА АНАЛИЗА ДОКУМЕНТОВ ===")
        print("Поддерживаемые форматы: PDF, DOC, DOCX")
        print("Варианты запуска:")
        print("1. Введите путь к папке для пакетной обработки")
        print("2. Введите путь к отдельному файлу")
        print("3. Нажмите Enter для обработки текущей папки")
        
        user_input = input("\nВведите путь: ").strip()
        
        if not user_input:
            return os.getcwd()
        
        if os.path.isfile(user_input):
            return user_input  # Вернем путь к файлу
        elif os.path.isdir(user_input):
            return user_input  # Вернем путь к папке
        else:
            attempts += 1
            print(f"❌ Путь '{user_input}' не существует!")
            if attempts < max_attempts:
                print(f"Осталось попыток: {max_attempts - attempts}")
            else:
                print("Превышено количество попыток. Используется текущая папка.")
                return os.getcwd()


def generate_output_filename(base_name: str, author: str = "Пользователь") -> str:
    """Генерирует структурированное имя файла типа 'Смета_Максим_08_25' (предыдущий_месяц_год)."""
    now = datetime.now()
    
    # Используем предыдущий месяц
    prev_month_date = now.replace(day=1) - timedelta(days=1)
    month = prev_month_date.strftime("%m")
    year = now.strftime("%y")
    
    # Убираем расширение из базового имени и заменяем пробелы на подчеркивания
    clean_base = os.path.splitext(os.path.basename(base_name))[0]
    clean_base = re.sub(r'[^\w\-_\.]', '_', clean_base)
    clean_author = re.sub(r'[^\w\-_\.]', '_', author)
    
    return f"Смета_{clean_author}_{clean_base}_{month}_{year}.xlsx"


def batch_process_directory(directory: str) -> dict:
    """Пакетная обработка всех поддерживаемых файлов в папке."""
    logger.info(f"Начало пакетной обработки папки: {directory}")
    
    print(f"\n🔍 Поиск файлов в папке: {directory}")
    files = find_supported_files(directory)
    
    if not files:
        logger.warning(f"В папке {directory} не найдено поддерживаемых файлов")
        print("❌ В указанной папке не найдено поддерживаемых файлов!")
        return {}
    
    logger.info(f"Найдено {len(files)} файлов для обработки")
    print(f"✅ Найдено файлов: {len(files)}")
    for i, file in enumerate(files, 1):
        print(f"  {i}. {os.path.basename(file)}")
        logger.debug(f"Файл {i}: {file}")
    
    results = {}
    total_cost = 0
    successful_files = 0
    failed_files = 0
    
    print(f"\n🚀 Начинаю обработку {len(files)} файлов...")
    logger.info(f"Начало обработки {len(files)} файлов")
    
    for i, file_path in enumerate(files, 1):
        file_name = os.path.basename(file_path)
        print(f"\n--- [{i}/{len(files)}] Обрабатываю: {file_name} ---")
        logger.info(f"[{i}/{len(files)}] Начинаю обработку файла: {file_path}")
        
        analyzer = None
        start_time = datetime.now()
        
        try:
            analyzer = PDFAnalyzer(file_path)
            analyzer.analyze()
            
            df = analyzer.get_summary_dataframe()
            if not df.empty:
                file_cost = df["Стоимость (руб.)"].sum()
                total_cost += file_cost
                
                results[file_path] = {
                    'cost': file_cost,
                    'pages': len(df),
                    'dataframe': df
                }
                
                processing_time = (datetime.now() - start_time).total_seconds()
                successful_files += 1
                
                print(f"✅ Обработан: {file_cost} руб., {len(df)} страниц")
                logger.info(f"Файл {file_name} успешно обработан: {file_cost} руб., {len(df)} страниц, время: {processing_time:.1f}с")
            else:
                logger.warning(f"Файл {file_name} обработан, но результатов нет")
                print(f"⚠️ Файл обработан, но результатов нет")
                failed_files += 1
                
        except Exception as e:
            processing_time = (datetime.now() - start_time).total_seconds()
            failed_files += 1
            
            logger.error(f"Ошибка обработки файла {file_name}: {e}, время: {processing_time:.1f}с")
            print(f"❌ Ошибка при обработке {file_name}: {e}")
            
        finally:
            # Обязательная очистка ресурсов
            if analyzer:
                try:
                    analyzer.cleanup()
                    logger.debug(f"Ресурсы для файла {file_name} очищены")
                except Exception as cleanup_error:
                    logger.warning(f"Ошибка очистки ресурсов для {file_name}: {cleanup_error}")
            
            # Принудительная очистка памяти каждые 3 файла
            if i % 3 == 0:
                gc.collect()
                logger.debug(f"Очистка памяти после {i} файлов")
    
    # Итоговая статистика
    total_time = (datetime.now() - start_time).total_seconds() if 'start_time' in locals() else 0
    
    print(f"\n🎉 Пакетная обработка завершена!")
    print(f"📊 Обработано файлов: {successful_files}/{len(files)}")
    print(f"💰 Общая стоимость: {total_cost} руб.")
    
    logger.info(f"Пакетная обработка завершена: успешно={successful_files}, ошибок={failed_files}, общая стоимость={total_cost} руб.")
    
    if failed_files > 0:
        logger.warning(f"Обработано с ошибками: {failed_files} файлов")
        print(f"⚠️ Файлов с ошибками: {failed_files}")
    
    return results


def create_summary_by_type_from_dataframe(df: pd.DataFrame, file_name: str) -> pd.DataFrame:
    """Создает сводную таблицу по типам страниц из готового DataFrame."""
    if df.empty:
        return pd.DataFrame()
    
    # Соответствие веса типу страниц
    weight_to_type = {
        4: "формулы",
        3: "иллюстрации/таблицы", 
        2: "нац. символы, сноски",
        1: "простые страницы"
    }
    
    # Группируем по весу
    summary = []
    
    # Добавляем заголовок с именем файла
    summary.append({
        "Тип страниц": file_name,
        "Кол-во стр.": "кол-во стр.",
        "Цена": "цена", 
        "Сумма": "сумма"
    })
    
    for weight in sorted(weight_to_type.keys(), reverse=True):  # От большего к меньшему
        weight_pages = df[df["Вес"] == weight]
        if not weight_pages.empty:
            count = len(weight_pages)
            price = PRICING[weight]  # Цена за страницу
            total = count * price
            summary.append({
                "Тип страниц": weight_to_type[weight],
                "Кол-во стр.": count,
                "Цена": f"{price:g}",  # Убираем .0 для целых чисел
                "Сумма": total
            })
    
    # Добавляем итоговую строку для файла
    if len(summary) > 1:  # Если есть данные кроме заголовка
        total_sum = df["Стоимость (руб.)"].sum()
        summary.append({
            "Тип страниц": "ИТОГО",
            "Кол-во стр.": "",
            "Цена": "",
            "Сумма": int(total_sum)
        })
    
    return pd.DataFrame(summary)

def save_batch_report(results: dict, output_dir: str, author: str = "Пользователь"):
    """Сохраняет сводный отчет по всем обработанным файлам."""
    if not results:
        print("❌ Нет данных для сохранения отчета")
        return
    
    # Создаем развернутую сводную таблицу по типам страниц для каждого файла
    detailed_summary = []
    total_cost = 0
    total_pages = 0
    
    for file_path, data in results.items():
        file_name = os.path.splitext(os.path.basename(file_path))[0]  # Убираем расширение
        df = data['dataframe']
        cost = data['cost']
        pages = data['pages']
        total_cost += cost
        total_pages += pages
        
        # Создаем сводку по типам для каждого файла
        file_summary = create_summary_by_type_from_dataframe(df, file_name)
        
        # Добавляем пустую строку между файлами (кроме первого)
        if detailed_summary:
            detailed_summary.append({
                "Тип страниц": "",
                "Кол-во стр.": "",
                "Цена": "",
                "Сумма": ""
            })
        
        # Добавляем сводку файла
        detailed_summary.extend(file_summary.to_dict('records'))
    
    # Добавляем общий итог
    if detailed_summary:
        detailed_summary.append({
            "Тип страниц": "",
            "Кол-во стр.": "",
            "Цена": "",
            "Сумма": ""
        })
        detailed_summary.append({
            "Тип страниц": "ОБЩИЙ ИТОГ",
            "Кол-во стр.": "",
            "Цена": "",
            "Сумма": int(total_cost)
        })
    
    # Создаем DataFrame для сводки
    summary_df = pd.DataFrame(detailed_summary)
    
    # Генерируем имя файла для сводного отчета
    now = datetime.now()
    # Используем предыдущий месяц
    prev_month_date = now.replace(day=1) - timedelta(days=1)
    month = prev_month_date.strftime("%m")
    year = now.strftime("%y")
    batch_filename = f"Смета_{author}_{month}_{year}.xlsx"
    batch_path = os.path.join(output_dir, batch_filename)
    
    try:
        with pd.ExcelWriter(batch_path, engine='openpyxl') as writer:
            # Развернутая сводная таблица на первом листе
            summary_df.to_excel(writer, sheet_name='Сводка', index=False)
            
            # Детальные отчеты для каждого файла на отдельных листах
            for file_path, data in results.items():
                file_name = os.path.basename(file_path)
                # Убираем расширение и ограничиваем длину для имени листа
                sheet_name = os.path.splitext(file_name)[0][:30]
                df_with_total = data['dataframe']
                
                # Создаем безопасную копию DataFrame для избежания циклических ссылок
                safe_df = pd.DataFrame({
                    'Страница': df_with_total['Страница'].astype(str),
                    'Вес': df_with_total['Вес'].astype(str),
                    'Стоимость (руб.)': df_with_total['Стоимость (руб.)'].astype(float),
                    'Обоснование': df_with_total['Обоснование'].astype(str)
                })
                
                # Добавляем итоговую строку для каждого файла
                file_summary = pd.DataFrame([{
                    "Страница": "ИТОГО",
                    "Вес": "",
                    "Стоимость (руб.)": float(data['cost']),
                    "Обоснование": ""
                }])
                
                df_with_total_final = pd.concat([safe_df, file_summary], ignore_index=True)
                df_with_total_final.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n📊 Сводный отчет сохранен: {batch_path}")
        print(f"📈 Итого: {len(results)} файлов, {total_pages} страниц, {total_cost} руб.")
        
    except Exception as e:
        print(f"❌ Ошибка при сохранении сводного отчета: {e}")


class PDFAnalyzer:
    """
    Класс для анализа страниц PDF/DOC/DOCX документов и классификации их по весу.
    """
    def __init__(self, file_path):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Файл не найден: {file_path}")
        
        logger.info(f"Инициализация анализатора для файла: {file_path}")
        
        self.original_file_path = file_path
        self.temp_pdf_path = None
        
        # Проверяем тип файла и конвертируем если нужно
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext in ['.doc', '.docx']:
            print(f"Обнаружен {file_ext.upper()} файл. Конвертирую в PDF...")
            logger.info(f"Конвертация {file_ext.upper()} файла в PDF")
            try:
                self.temp_pdf_path = convert_doc_to_pdf(file_path)
                self.file_path = self.temp_pdf_path
                print(f"Конвертация завершена. Анализирую PDF: {os.path.basename(self.temp_pdf_path)}")
                logger.info(f"Конвертация успешно завершена: {self.temp_pdf_path}")
            except Exception as e:
                logger.error(f"Ошибка конвертации файла {file_path}: {e}")
                raise
        else:
            self.file_path = file_path
            
        self.analysis_results = []
        
        # Используем глобальную модель вместо создания новой
        try:
            print("Получение модели распознавания формул...")
            self.p2t = get_pix2text_model()
            print("Модель готова.")
            logger.info("Модель Pix2Text получена успешно")
        except Exception as e:
            logger.error(f"Не удалось получить модель Pix2Text: {e}")
            self.p2t = None  # Продолжаем работу без модели

    def _has_formulas(self, text: str) -> bool:
        """Проверяет текст на наличие математических формул по паттернам."""
        # Сначала очищаем текст от лишних пробелов и переносов для лучшего распознавания функций
        cleaned_text = re.sub(r'\s+', ' ', text.strip())
        
        # Также проверяем исходный текст для многострочных формул
        for pattern in FORMULA_PATTERNS:
            if pattern.search(text) or pattern.search(cleaned_text):
                return True
        
        # Дополнительная проверка для многострочных функций типа:
        # ( )
        # f t
        multiline_func_pattern = re.compile(r'\(\s*[a-zA-Z0-9,\s]*\s*\)\s*\n\s*[a-zA-Z]\s+[a-zA-Z0-9,\s]*')
        if multiline_func_pattern.search(text):
            return True
            
        return False

    def _image_has_formulas(self, page: fitz.Page) -> bool:
        """Проверяет изображения на странице на наличие формул с помощью Pix2Text."""
        if self.p2t is None:
            logger.warning("Модель Pix2Text недоступна, пропускаем анализ изображений")
            return False
            
        images = page.get_images(full=True)
        if not images:
            return False

        doc = page.parent  # Получаем объект документа из страницы
        page_num = page.number + 1

        for img_info in images:
            xref = img_info[0]
            try:
                logger.debug(f"Анализ изображения {xref} на странице {page_num}")
                
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                
                # Создаем PIL Image из байтов
                pil_image = Image.open(io.BytesIO(image_bytes))
                
                # Конвертируем в RGB если нужно
                if pil_image.mode != 'RGB':
                    pil_image = pil_image.convert('RGB')
                
                # Используем PIL Image напрямую, а не numpy array
                try:
                    results = self.p2t.recognize(pil_image)
                    logger.debug(f"Результаты распознавания: {len(results) if results else 0} элементов")
                    
                    if results:
                        for res in results:
                            if isinstance(res, dict) and res.get('type') == 'formula':
                                logger.info(f"Найдена формула в изображении на странице {page_num}")
                                return True
                            elif isinstance(res, dict) and 'text' in res:
                                # Проверяем текст на наличие формул
                                if self._has_formulas(res.get('text', '')):
                                    logger.info(f"Найдена формула в тексте изображения на странице {page_num}")
                                    return True
                                    
                except Exception as model_error:
                    logger.warning(f"Ошибка ML-модели при анализе изображения на стр. {page_num}: {model_error}")
                    continue
                    
            except Exception as e:
                logger.warning(f"Не удалось проанализировать изображение на стр. {page_num}: {e}")
                print(f"  - Предупреждение: не удалось проанализировать изображение на стр. {page_num}: {e}")
                continue
                
        return False

    def _has_special_chars(self, text: str) -> bool:
        """Проверяет текст на наличие иероглифов или арабской вязи."""
        return bool(SPECIAL_CHARS_PATTERN.search(text))
    
    def _has_footnotes(self, text: str) -> bool:
        """Проверяет текст на наличие числовых сносок с парным соответствием текст-определение."""
        lines = text.split('\n')
        
        # Паттерн 1: Ссылки в тексте - русские слова с цифрами 
        # Ищем: сандриков3, Мезонин4 (русское слово + цифра) И Маркианович,1 (слово + знак препинания + цифра)
        text_footnote_refs = set()
        
        # Подпаттерн 1.1: слово+цифра (сандриков3)
        pattern1 = re.compile(r'[А-Яа-яё]+(\d+)', re.IGNORECASE)
        for match in pattern1.finditer(text):
            footnote_num = match.group(1)
            if len(footnote_num) <= 2 and not (len(footnote_num) == 4 and footnote_num.startswith(('18', '19', '20'))):
                text_footnote_refs.add(footnote_num)
        
        # Подпаттерн 1.2: буква+знак препинания+цифра (Маркианович,1)
        pattern2 = re.compile(r'[а-яё][,\.;:]\s*(\d+)', re.IGNORECASE)
        for match in pattern2.finditer(text):
            footnote_num = match.group(1)
            if len(footnote_num) <= 2 and not (len(footnote_num) == 4 and footnote_num.startswith(('18', '19', '20'))):
                text_footnote_refs.add(footnote_num)
        
        # Паттерн 2: Определения сносок внизу страницы - строки начинающиеся с "цифра пробел текст"
        footnote_definitions = set()
        for line in lines:
            line = line.strip()
            if line:
                # Ищем строки типа "3 На 1841 год..." или "4 По крайней мере..."
                definition_match = re.match(r'^(\d+)\s+[А-Яа-яё]', line)
                if definition_match:
                    footnote_definitions.add(definition_match.group(1))
        
        # Парное соответствие: ищем пересечения номеров сносок
        matched_footnotes = text_footnote_refs.intersection(footnote_definitions)
        
        # Возвращаем True, если найдено хотя бы одно парное соответствие
        return len(matched_footnotes) > 0
    
    def _is_valid_table(self, table, page_text=""):
        """
        Проверяет, является ли найденная 'таблица' действительной таблицей,
        а не структурированным текстом.
        """
        if not table or len(table) == 0:
            return False
        
        # Получаем размеры таблицы
        rows = len(table)
        cols = len(table[0]) if table[0] else 0
        
        # Критерий 1: Минимальные размеры
        if rows < 2 or cols < 2:
            return False
        
        # Критерий 2: Проверка на дублирование содержимого
        if len(table) > 0 and len(table[0]) >= 2:
            first_cell = table[0][0]
            second_cell = table[0][1]
            
            if first_cell and second_cell:
                if first_cell == second_cell:
                    return False
                if first_cell in second_cell or second_cell in first_cell:
                    return False
        
        # Критерий 3: Проверка структуры данных
        non_empty_cells = 0
        total_cells = 0
        unique_values = set()
        
        for row in table:
            for cell in row:
                total_cells += 1
                if cell and cell.strip():
                    non_empty_cells += 1
                    unique_values.add(cell.strip())
        
        # Если меньше 50% ячеек заполнены, это скорее всего не таблица
        if total_cells > 0 and (non_empty_cells / total_cells) < 0.5:
            return False
        
        # Критерий 4: Проверка на пустые столбцы
        # Если весь столбец состоит из None/пустых значений, это не таблица
        empty_columns = 0
        for col_idx in range(cols):
            column_empty = True
            for row in table:
                if col_idx < len(row) and row[col_idx] and row[col_idx].strip():
                    column_empty = False
                    break
            if column_empty:
                empty_columns += 1
        
        # Если больше половины столбцов пустые, это не таблица
        if empty_columns >= cols / 2:
            return False
        
        # Критерий 5: Проверка на числовые данные (характерно для таблиц)
        numeric_cells = 0
        for row in table:
            for cell in row:
                if cell and any(char.isdigit() for char in cell):
                    numeric_cells += 1
        
        # Критерий 6: Проверка на табличные маркеры
        # Настоящие таблицы часто содержат типичные элементы
        table_indicators = 0
        all_text = " ".join([str(cell) for row in table for cell in row if cell])
        
        # Поиск табличных паттернов (расширенный список для описательных таблиц)
        tabular_words = ['№', 'номер', 'количество', 'сумма', 'итого', 'всего', 
                        'тип', 'вид', 'название', 'наименование', 'характеристика',
                        'описание', 'значение', 'параметр', 'критерий', 'фактор',
                        'аспект', 'влияние', 'воздействие', 'категория', 'группа',
                        'показатель', 'индикатор', 'метод', 'способ', 'форма']
        if any(word in all_text.lower() for word in tabular_words):
            table_indicators += 1
        if numeric_cells >= 3:  # Достаточно чисел для таблицы
            table_indicators += 1
        if len(unique_values) <= rows * 0.7:  # Есть повторяющиеся значения (заголовки, категории)
            table_indicators += 1
            
        # Критерий 7: Проверка на текстовые характеристики
        # Если это просто разбитый текст, то:
        text_indicators = 0
        
        # Проверяем, является ли это последовательными предложениями
        if len(unique_values) == non_empty_cells and non_empty_cells > 5:  # Все уникально
            # Попробуем объединить и посмотреть, получится ли связный текст
            combined = " ".join(unique_values)
            if len(combined.split()) / len(unique_values) > 3:  # Много слов на ячейку
                text_indicators += 1
        
        # Если один столбец полностью пустой, а другой содержит длинные фразы
        if empty_columns > 0 and non_empty_cells > 0:
            avg_cell_length = sum(len(str(cell)) for row in table for cell in row if cell) / non_empty_cells
            if avg_cell_length > 50:  # Длинные фразы
                text_indicators += 1
        
        # Критерий 8: Проверка на равномерность строк
        filled_cells_per_row = []
        for row in table:
            filled = sum(1 for cell in row if cell and cell.strip())
            filled_cells_per_row.append(filled)
        
        if filled_cells_per_row:
            avg_filled = sum(filled_cells_per_row) / len(filled_cells_per_row)
            if avg_filled > 0:
                variance = sum((x - avg_filled) ** 2 for x in filled_cells_per_row) / len(filled_cells_per_row)
                if variance / avg_filled > 2:  # Слишком неравномерно
                    return False
        
        # Итоговое решение: улучшенная логика для описательных таблиц
        # Правило 1: Классические таблицы (числовые, со стандартными индикаторами)
        if table_indicators >= 2 and text_indicators == 0:
            return True
        # Правило 2: Явно не таблицы (только текстовые признаки)
        elif table_indicators == 0 and text_indicators > 0:
            return False
        # Правило 3: Числовые таблицы
        elif numeric_cells > 0 and empty_columns == 0:
            return True
        # Правило 4: Описательные таблицы (новое!)
        elif table_indicators >= 1 and empty_columns == 0 and rows >= 2 and cols == 2:
            # Проверяем, что это структурированная таблица "термин-описание"
            # Если первый столбец содержит короткие термины, а второй - описания
            first_col_lengths = []
            second_col_lengths = []
            for row in table:
                if len(row) >= 2 and row[0] and row[1]:
                    first_col_lengths.append(len(str(row[0])))
                    second_col_lengths.append(len(str(row[1])))
            
            if first_col_lengths and second_col_lengths:
                avg_first = sum(first_col_lengths) / len(first_col_lengths)
                avg_second = sum(second_col_lengths) / len(second_col_lengths)
                # Если второй столбец в среднем значительно длиннее первого
                if avg_second > avg_first * 2 and avg_first < 100:
                    return True
        
        return False
    
    def _has_valid_tables(self, plumber_page, page_text=""):
        """Улучшенное распознавание таблиц с фильтрацией ложных срабатываний."""
        try:
            tables = plumber_page.extract_tables()
            
            if not tables:
                return False
            
            # Проверяем каждую найденную таблицу
            for table in tables:
                if self._is_valid_table(table, page_text):
                    return True
                    
            return False
            
        except Exception as e:
            logger.warning(f"Ошибка улучшенного анализа таблиц: {e}")
            return False

    def analyze(self):
        """
        Основной метод для анализа всего документа.
        Перебирает страницы и определяет вес каждой из них.
        """
        print(f"Начинается анализ файла: {os.path.basename(self.file_path)}")
        logger.info(f"Начало анализа документа: {self.file_path}")
        
        doc = None
        pdf = None
        
        try:
            # Открываем документ с обработкой ошибок
            try:
                doc = fitz.open(self.file_path)
                logger.info(f"PyMuPDF: документ открыт, {len(doc)} страниц")
            except Exception as e:
                logger.error(f"Ошибка открытия документа PyMuPDF: {e}")
                raise RuntimeError(f"Не удалось открыть файл с PyMuPDF: {e}")
            
            try:
                pdf = pdfplumber.open(self.file_path)
                logger.info(f"pdfplumber: документ открыт, {len(pdf.pages)} страниц")
            except Exception as e:
                logger.warning(f"Ошибка открытия документа pdfplumber: {e}")
                # Продолжаем только с PyMuPDF
                pdf = None
                
            if pdf and len(doc) != len(pdf.pages):
                logger.warning("PyMuPDF и pdfplumber распознали разное количество страниц")
                print("Предупреждение: PyMuPDF и pdfplumber распознали разное количество страниц.")
            
            num_pages = len(doc)
            print(f"Всего страниц в документе: {num_pages}")
            logger.info(f"Обработка {num_pages} страниц")

            for i in range(num_pages):
                page_num = i + 1
                
                try:
                    page_weight, reason = self._analyze_single_page(doc, pdf, i)
                    
                    self.analysis_results.append({
                        "Страница": page_num,
                        "Вес": page_weight,
                        "Стоимость (руб.)": PRICING[page_weight],
                        "Обоснование": reason
                    })
                    
                    print(f"  - Страница {page_num}/{num_pages} обработана. Вес: {page_weight} ({reason})")
                    logger.debug(f"Страница {page_num}: вес={page_weight}, причина={reason}")
                    
                    # Принудительная очистка памяти каждые 10 страниц
                    if page_num % 10 == 0:
                        gc.collect()
                        logger.debug(f"Очистка памяти после страницы {page_num}")
                        
                except Exception as page_error:
                    logger.error(f"Ошибка обработки страницы {page_num}: {page_error}")
                    # Добавляем страницу с минимальным весом при ошибке
                    self.analysis_results.append({
                        "Страница": page_num,
                        "Вес": 1,
                        "Стоимость (руб.)": PRICING[1],
                        "Обоснование": f"Ошибка анализа: {str(page_error)[:50]}"
                    })
                    print(f"  - Страница {page_num}/{num_pages} - ОШИБКА, назначен минимальный вес")

        except Exception as e:
            logger.error(f"Критическая ошибка при обработке файла: {e}")
            print(f"Ошибка при обработке файла: {e}")
            raise
        finally:
            # Обязательно закрываем все ресурсы
            try:
                if pdf:
                    pdf.close()
                    logger.debug("pdfplumber ресурсы освобождены")
            except Exception as e:
                logger.warning(f"Ошибка закрытия pdfplumber: {e}")
                
            try:
                if doc:
                    doc.close()
                    logger.debug("PyMuPDF ресурсы освобождены")
            except Exception as e:
                logger.warning(f"Ошибка закрытия PyMuPDF: {e}")
            
            # Принудительная очистка памяти
            gc.collect()
            logger.info(f"Анализ документа завершен. Обработано страниц: {len(self.analysis_results)}")
            
        # Возвращаем результаты анализа
        return self.analysis_results

    def _analyze_single_page(self, doc, pdf, page_index):
        """Анализ одной страницы с обработкой ошибок."""
        page_num = page_index + 1
        page_weight = 1  # Вес по умолчанию
        reason = "Простой текст"
        
        try:
            # Получаем страницу PyMuPDF
            pymupdf_page = doc[page_index]
            page_text = pymupdf_page.get_text("text")
            logger.debug(f"Страница {page_num}: извлечен текст длиной {len(page_text)} символов")

            # --- Иерархическая проверка с обработкой ошибок ---
            
            # 1. Проверка на текстовые формулы
            try:
                if self._has_formulas(page_text):
                    page_weight = 4
                    reason = "Формула (текст)"
                    return page_weight, reason
            except Exception as e:
                logger.warning(f"Ошибка анализа текстовых формул на стр. {page_num}: {e}")
            
            # 2. Проверка на формулы в изображениях
            try:
                if self._image_has_formulas(pymupdf_page):
                    page_weight = 4
                    reason = "Формула (изображение)"
                    return page_weight, reason
            except Exception as e:
                logger.warning(f"Ошибка анализа изображений на стр. {page_num}: {e}")
            
            # 3. Проверка на "обычные" изображения
            try:
                if pymupdf_page.get_images(full=True):
                    page_weight = 3
                    reason = "Изображение"
                    return page_weight, reason
            except Exception as e:
                logger.warning(f"Ошибка поиска изображений на стр. {page_num}: {e}")
            
            # 4. Проверка на таблицы (только если pdfplumber доступен)
            if pdf:
                try:
                    plumber_page = pdf.pages[page_index]
                    if self._has_valid_tables(plumber_page, page_text):
                        page_weight = 3
                        reason = "Таблица"
                        return page_weight, reason
                except Exception as e:
                    logger.warning(f"Ошибка поиска таблиц на стр. {page_num}: {e}")
            
            # 5. Проверка на спецсимволы
            try:
                if self._has_special_chars(page_text):
                    page_weight = 2
                    reason = "Иероглифы/араб. вязь"
                    return page_weight, reason
            except Exception as e:
                logger.warning(f"Ошибка анализа спецсимволов на стр. {page_num}: {e}")
            
            # 6. Проверка на сноски
            try:
                if self._has_footnotes(page_text):
                    page_weight = 2
                    reason = "Текст со сносками"
                    return page_weight, reason
            except Exception as e:
                logger.warning(f"Ошибка анализа сносок на стр. {page_num}: {e}")
                
        except Exception as e:
            logger.error(f"Общая ошибка анализа страницы {page_num}: {e}")
            
        return page_weight, reason

    def get_summary_dataframe(self) -> pd.DataFrame:
        """Возвращает результаты анализа в виде pandas DataFrame."""
        if not self.analysis_results:
            return pd.DataFrame()
        return pd.DataFrame(self.analysis_results)

    def print_total_cost(self):
        """Выводит итоговую таблицу и общую стоимость."""
        df = self.get_summary_dataframe()
        if df.empty:
            print("Анализ не дал результатов.")
            return

        total_cost = df["Стоимость (руб.)"].sum()
        
        print("\n--- Детальный отчет по страницам ---")
        print(df.to_string(index=False))
        
        print("\n--- Итоговая стоимость ---")
        print(f"Общая стоимость обработки документа: {total_cost} руб.")

    def get_summary_by_type(self) -> pd.DataFrame:
        """Создает сводную таблицу по типам страниц."""
        df = self.get_summary_dataframe()
        if df.empty:
            return pd.DataFrame()
        
        # Соответствие веса типу страниц
        weight_to_type = {
            4: "формулы",
            3: "иллюстрации/таблицы", 
            2: "нац. символы, сноски",
            1: "простые страницы"
        }
        
        # Группируем по весу
        summary = []
        for weight in sorted(weight_to_type.keys(), reverse=True):  # От большего к меньшему
            weight_pages = df[df["Вес"] == weight]
            if not weight_pages.empty:
                count = len(weight_pages)
                price = PRICING[weight]  # Цена за страницу
                total = count * price
                summary.append({
                    "Тип страниц": weight_to_type[weight],
                    "Кол-во стр.": count,
                    "Цена": f"{price:g}",  # Убираем .0 для целых чисел
                    "Сумма": total
                })
        
        # Добавляем итоговую строку
        if summary:
            total_sum = sum(item["Сумма"] for item in summary)
            summary.append({
                "Тип страниц": "ИТОГО",
                "Кол-во стр.": "",
                "Цена": "",
                "Сумма": total_sum
            })
        
        return pd.DataFrame(summary)

    def save_to_excel(self, output_path: str):
        """Сохраняет детальный отчет и итоговую стоимость в Excel-файл."""
        df = self.get_summary_dataframe()
        if df.empty:
            print("Нет данных для сохранения в Excel.")
            return

        # Получаем имя файла из пути для заголовка
        file_name = os.path.splitext(os.path.basename(self.file_path))[0]
        
        # Создаем сводную таблицу по типам страниц
        summary_by_type = self.get_summary_by_type()
        
        total_cost = df["Стоимость (руб.)"].sum()
        
        # Создаем строку с итогом для детального отчета
        summary_row = pd.DataFrame([{
            "Страница": "ИТОГО",
            "Вес": "",
            "Стоимость (руб.)": total_cost,
            "Обоснование": ""
        }])
        
        # Добавляем итоговую строку в конец детального отчета
        df_with_total = pd.concat([df, summary_row], ignore_index=True)

        try:
            # Создаем Excel файл с несколькими листами
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Лист 1: Сводная информация
                if not summary_by_type.empty:
                    # Добавляем заголовок с именем файла
                    header_df = pd.DataFrame([{
                        "Тип страниц": file_name,
                        "Кол-во стр.": "кол-во стр.",
                        "Цена": "цена", 
                        "Сумма": "сумма"
                    }])
                    
                    # Объединяем заголовок со сводкой
                    summary_with_header = pd.concat([header_df, summary_by_type], ignore_index=True)
                    summary_with_header.to_excel(writer, sheet_name='Сводка', index=False)
                
                # Лист 2: Детальный отчет по страницам
                df_with_total.to_excel(writer, sheet_name='Детальный отчет', index=False)
                
            print(f"\nОтчет успешно сохранен в файл: {output_path}")
        except Exception as e:
            print(f"\nНе удалось сохранить отчет в Excel: {e}")

    def cleanup(self):
        """Очищает временные файлы и освобождает ресурсы."""
        logger.info("Начало очистки ресурсов PDFAnalyzer")
        
        # Очистка временных файлов
        if hasattr(self, 'temp_pdf_path') and self.temp_pdf_path:
            try:
                if os.path.exists(self.temp_pdf_path):
                    os.unlink(self.temp_pdf_path)
                    print(f"Временный файл удален: {os.path.basename(self.temp_pdf_path)}")
                    logger.info(f"Временный файл удален: {self.temp_pdf_path}")
                else:
                    logger.debug(f"Временный файл уже не существует: {self.temp_pdf_path}")
            except Exception as e:
                logger.warning(f"Не удалось удалить временный файл {self.temp_pdf_path}: {e}")
                print(f"Предупреждение: не удалось удалить временный файл {self.temp_pdf_path}: {e}")
        
        # Очистка результатов анализа для освобождения памяти
        if hasattr(self, 'analysis_results'):
            self.analysis_results.clear()
            logger.debug("Результаты анализа очищены")
        
        # Обнуление ссылки на модель (не удаляем глобальную модель)
        self.p2t = None
        
        # Принудительная очистка памяти
        gc.collect()
        logger.info("Очистка ресурсов PDFAnalyzer завершена")

    def __del__(self):
        """Деструктор для автоматической очистки временных файлов."""
        try:
            # Только удаляем временные файлы, без логирования
            if hasattr(self, 'temp_pdf_path') and self.temp_pdf_path and os.path.exists(self.temp_pdf_path):
                os.unlink(self.temp_pdf_path)
        except:
            pass  # Игнорируем все ошибки в деструкторе


if __name__ == '__main__':
    # --- Интерактивный режим ---
    logger.info("Запуск программы Смета МУ")
    print("📊 Система анализа документов 'Смета МУ' v2.0")
    
    try:
        user_path = get_directory_from_user()
        logger.info(f"Пользователь выбрал путь: {user_path}")
        
        # Запрашиваем имя автора для структурированного вывода
        author = input("\nВведите имя автора для отчета (по умолчанию 'Максим'): ").strip()
        if not author:
            author = "Максим"
        logger.info(f"Автор отчета: {author}")
        
        start_time = datetime.now()
        
        # Определяем, это файл или папка
        if os.path.isfile(user_path):
            # Обработка одного файла
            print(f"\n📄 Обрабатываю отдельный файл: {os.path.basename(user_path)}")
            logger.info(f"Режим: обработка одного файла - {user_path}")
            
            analyzer = None
            try:
                analyzer = PDFAnalyzer(user_path)
                analyzer.analyze()
                analyzer.print_total_cost()

                # Генерируем структурированное имя файла
                output_filename = generate_output_filename(user_path, author)
                output_dir = os.path.dirname(user_path) if os.path.dirname(user_path) else os.getcwd()
                excel_output_path = os.path.join(output_dir, output_filename)
                
                analyzer.save_to_excel(excel_output_path)
                logger.info(f"Отчет сохранен: {excel_output_path}")
                
            except Exception as e:
                logger.error(f"Ошибка обработки файла {user_path}: {e}")
                print(f"❌ Ошибка при обработке файла: {e}")
            finally:
                if analyzer:
                    analyzer.cleanup()
        
        elif os.path.isdir(user_path):
            # Пакетная обработка папки
            print(f"\n📁 Пакетная обработка папки: {user_path}")
            logger.info(f"Режим: пакетная обработка папки - {user_path}")
            
            results = batch_process_directory(user_path)
            
            if results:
                # Сохраняем сводный отчет
                save_batch_report(results, user_path, author)
                logger.info("Сводный отчет успешно создан")
                
                # Принудительная финальная очистка памяти
                gc.collect()
                logger.debug("Финальная очистка памяти")
            else:
                logger.warning("Нет файлов для обработки")
                print("❌ Нет файлов для обработки")
        
        # Итоговая статистика времени
        total_time = (datetime.now() - start_time).total_seconds()
        logger.info(f"Общее время работы программы: {total_time:.1f} секунд")
        print(f"\n⏱️ Время работы: {total_time:.1f}с")
                
    except KeyboardInterrupt:
        logger.info("Программа прервана пользователем")
        print("\n\n👋 Программа прервана пользователем")
    except Exception as e:
        logger.error(f"Критическая ошибка программы: {e}")
        print(f"\n❌ Критическая ошибка: {e}")
    finally:
        # Финальная очистка ресурсов
        gc.collect()
        logger.info("Завершение работы программы")
    
    print("\n✅ Работа завершена!")