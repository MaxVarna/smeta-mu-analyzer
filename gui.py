#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Графический интерфейс для системы анализа документов "Смета МУ"
Поддерживает PDF, DOC, DOCX файлы с весовой классификацией страниц
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys
from datetime import datetime

# Импортируем функции из основного модуля
from main import PDFAnalyzer, find_supported_files, batch_process_directory, save_batch_report, generate_output_filename


class DocumentAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Смета МУ - Анализ документов")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Настройка стилей
        style = ttk.Style()
        style.theme_use('clam')
        
        # Переменные
        self.selected_path = tk.StringVar()
        self.author_name = tk.StringVar(value="Максим")
        self.processing = False
        
        self.setup_ui()
        
    def setup_ui(self):
        """Создает пользовательский интерфейс."""
        # Главный фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройка растягивания
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="🔍 СИСТЕМА АНАЛИЗА ДОКУМЕНТОВ", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        subtitle_label = ttk.Label(main_frame, text="Поддерживаемые форматы: PDF, DOC, DOCX", 
                                  font=('Arial', 10))
        subtitle_label.grid(row=1, column=0, columnspan=3, pady=(0, 20))
        
        # Поле автора
        ttk.Label(main_frame, text="Автор отчета:").grid(row=2, column=0, sticky=tk.W, pady=5)
        author_entry = ttk.Entry(main_frame, textvariable=self.author_name, width=20)
        author_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
        # Выбор пути
        ttk.Label(main_frame, text="Выбранный путь:").grid(row=3, column=0, sticky=tk.W, pady=5)
        path_entry = ttk.Entry(main_frame, textvariable=self.selected_path, state='readonly')
        path_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
        # Кнопки выбора
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=4, column=0, columnspan=3, pady=20)
        
        ttk.Button(buttons_frame, text="📄 Выбрать файл", 
                  command=self.select_file).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(buttons_frame, text="📁 Выбрать папку", 
                  command=self.select_folder).pack(side=tk.LEFT, padx=5)
        
        # Кнопка запуска
        self.process_button = ttk.Button(main_frame, text="🚀 НАЧАТЬ ОБРАБОТКУ", 
                                        command=self.start_processing, style='Accent.TButton')
        self.process_button.grid(row=5, column=0, columnspan=3, pady=20)
        
        # Прогресс-бар
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Область вывода
        ttk.Label(main_frame, text="Результат обработки:").grid(row=7, column=0, sticky=tk.W)
        
        self.output_text = scrolledtext.ScrolledText(main_frame, height=15, width=70)
        self.output_text.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Растягивание текстового поля
        main_frame.rowconfigure(8, weight=1)
        
        # Статус-бар
        self.status_var = tk.StringVar(value="Готов к работе")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
    def select_file(self):
        """Диалог выбора файла."""
        file_path = filedialog.askopenfilename(
            title="Выберите документ для анализа",
            filetypes=[
                ("Все поддерживаемые", "*.pdf *.doc *.docx"),
                ("PDF файлы", "*.pdf"),
                ("Word документы", "*.doc *.docx"),
                ("Все файлы", "*.*")
            ]
        )
        
        if file_path:
            self.selected_path.set(file_path)
            self.status_var.set(f"Выбран файл: {os.path.basename(file_path)}")
            
    def select_folder(self):
        """Диалог выбора папки."""
        folder_path = filedialog.askdirectory(title="Выберите папку с документами")
        
        if folder_path:
            self.selected_path.set(folder_path)
            # Подсчитаем файлы в папке
            files = find_supported_files(folder_path)
            self.status_var.set(f"Выбрана папка с {len(files)} поддерживаемыми файлами")
            
    def log_output(self, message):
        """Добавляет сообщение в область вывода."""
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.root.update_idletasks()
        
    def start_processing(self):
        """Запускает обработку в отдельном потоке."""
        if not self.selected_path.get():
            messagebox.showwarning("Предупреждение", "Пожалуйста, выберите файл или папку для обработки")
            return
            
        if not self.author_name.get().strip():
            messagebox.showwarning("Предупреждение", "Пожалуйста, введите имя автора")
            return
            
        if self.processing:
            messagebox.showinfo("Информация", "Обработка уже выполняется")
            return
            
        # Очищаем область вывода
        self.output_text.delete(1.0, tk.END)
        
        # Запускаем обработку в отдельном потоке
        thread = threading.Thread(target=self.process_documents)
        thread.daemon = True
        thread.start()
        
    def process_documents(self):
        """Обрабатывает документы (выполняется в отдельном потоке)."""
        try:
            self.processing = True
            self.process_button.configure(state='disabled')
            self.progress.start(10)
            self.status_var.set("Обработка...")
            
            path = self.selected_path.get()
            author = self.author_name.get().strip()
            
            if os.path.isfile(path):
                self.process_single_file(path, author)
            elif os.path.isdir(path):
                self.process_folder(path, author)
            else:
                self.log_output("❌ Ошибка: Указанный путь не существует")
                
        except Exception as e:
            self.log_output(f"❌ Критическая ошибка: {e}")
            messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")
        finally:
            self.processing = False
            self.process_button.configure(state='normal')
            self.progress.stop()
            self.status_var.set("Готов к работе")
            
    def process_single_file(self, file_path, author):
        """Обработка одного файла."""
        self.log_output(f"📄 Обработка файла: {os.path.basename(file_path)}")
        self.log_output("=" * 50)
        
        analyzer = None
        try:
            # Перехватываем вывод анализатора
            original_print = print
            def gui_print(*args, **kwargs):
                message = " ".join(str(arg) for arg in args)
                self.log_output(message)
                
            # Временно заменяем print
            import builtins
            builtins.print = gui_print
            
            try:
                analyzer = PDFAnalyzer(file_path)
                analyzer.analyze()
                analyzer.print_total_cost()
                
                # Генерируем имя файла и сохраняем
                output_filename = generate_output_filename(file_path, author)
                output_dir = os.path.dirname(file_path) if os.path.dirname(file_path) else os.getcwd()
                excel_path = os.path.join(output_dir, output_filename)
                
                analyzer.save_to_excel(excel_path)
                
                self.log_output("=" * 50)
                self.log_output("✅ ОБРАБОТКА ЗАВЕРШЕНА УСПЕШНО!")
                self.log_output(f"📊 Отчет сохранен: {output_filename}")
                
                # Показываем диалог успеха
                self.root.after(0, lambda: messagebox.showinfo(
                    "Успех", 
                    f"Файл обработан успешно!\nОтчет сохранен: {output_filename}"
                ))
                
            finally:
                # Восстанавливаем оригинальный print
                builtins.print = original_print
                
        except Exception as e:
            self.log_output(f"❌ Ошибка при обработке: {e}")
        finally:
            if analyzer:
                analyzer.cleanup()
                
    def process_folder(self, folder_path, author):
        """Пакетная обработка папки."""
        self.log_output(f"📁 Пакетная обработка папки: {folder_path}")
        self.log_output("=" * 50)
        
        # Перехватываем вывод для GUI
        original_print = print
        def gui_print(*args, **kwargs):
            message = " ".join(str(arg) for arg in args)
            self.log_output(message)
            
        import builtins
        builtins.print = gui_print
        
        try:
            results = batch_process_directory(folder_path)
            
            if results:
                save_batch_report(results, folder_path, author)
                
                # Очистка временных файлов
                for data in results.values():
                    if 'analyzer' in data:
                        data['analyzer'].cleanup()
                        
                self.log_output("=" * 50)
                self.log_output("✅ ПАКЕТНАЯ ОБРАБОТКА ЗАВЕРШЕНА!")
                
                # Показываем результат
                total_files = len(results)
                total_cost = sum(data['cost'] for data in results.values())
                total_pages = sum(data['pages'] for data in results.values())
                
                success_msg = f"""Обработано файлов: {total_files}
Всего страниц: {total_pages}
Общая стоимость: {total_cost} руб.
Сводный отчет: Смета_{author}_{datetime.now().strftime('%m_%y')}.xlsx"""
                
                self.root.after(0, lambda: messagebox.showinfo("Успех", success_msg))
            else:
                self.log_output("❌ Не найдено файлов для обработки")
                
        finally:
            builtins.print = original_print


def main():
    """Главная функция для запуска GUI."""
    root = tk.Tk()
    
    # Настройка иконки (если есть)
    try:
        # Можно добавить иконку приложения
        # root.iconbitmap('icon.ico')
        pass
    except:
        pass
    
    app = DocumentAnalyzerGUI(root)
    
    # Обработка закрытия окна
    def on_closing():
        if app.processing:
            if messagebox.askokcancel("Выход", "Обработка еще не завершена. Вы уверены, что хотите выйти?"):
                root.destroy()
        else:
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()