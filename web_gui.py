#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Веб-интерфейс для системы анализа документов "Смета МУ"
Работает через браузер, не требует установки дополнительных GUI библиотек
"""

from flask import Flask, render_template_string, request, jsonify, send_file, redirect, url_for
import os
import tempfile
import threading
from datetime import datetime
import json

# Импортируем функции из основного модуля
from main import PDFAnalyzer, find_supported_files, batch_process_directory, save_batch_report, generate_output_filename

app = Flask(__name__)
app.secret_key = 'smeta_mu_secret_key_2024'

# Глобальные переменные для отслеживания прогресса
processing_status = {
    'active': False,
    'progress': 0,
    'current_file': '',
    'log': [],
    'result_file': None
}

# HTML шаблон
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Смета МУ - Анализ документов</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .container {
            max-width: 900px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }

        .header h1 {
            font-size: 2.5rem;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }

        .header p {
            font-size: 1.1rem;
            opacity: 0.9;
        }

        .content {
            padding: 40px;
        }

        .form-group {
            margin-bottom: 25px;
        }

        label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #333;
            font-size: 1.1rem;
        }

        input[type="text"], input[type="file"] {
            width: 100%;
            padding: 12px 15px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-size: 1rem;
            transition: border-color 0.3s;
        }

        input[type="text"]:focus, input[type="file"]:focus {
            outline: none;
            border-color: #4facfe;
        }

        .file-input-wrapper {
            position: relative;
            display: inline-block;
            width: 100%;
        }

        .btn {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            color: white;
            border: none;
            padding: 12px 25px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 1rem;
            font-weight: 600;
            transition: all 0.3s;
            margin-right: 10px;
            margin-bottom: 10px;
        }

        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(79, 172, 254, 0.3);
        }

        .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 15px 30px;
            font-size: 1.2rem;
        }

        .btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }

        .progress-container {
            margin-top: 20px;
            display: none;
        }

        .progress-bar {
            width: 100%;
            height: 20px;
            background: #e0e0e0;
            border-radius: 10px;
            overflow: hidden;
        }

        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #4facfe, #00f2fe);
            border-radius: 10px;
            transition: width 0.3s;
            animation: pulse 2s infinite;
        }

        @keyframes pulse {
            0% { opacity: 1; }
            50% { opacity: 0.7; }
            100% { opacity: 1; }
        }

        .log-container {
            margin-top: 20px;
            display: none;
        }

        .log-box {
            background: #f8f9fa;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: 15px;
            height: 300px;
            overflow-y: auto;
            font-family: 'Courier New', monospace;
            font-size: 0.9rem;
            line-height: 1.4;
        }

        .result-container {
            margin-top: 20px;
            display: none;
            text-align: center;
        }

        .success {
            color: #28a745;
            font-size: 1.2rem;
            margin-bottom: 15px;
        }

        .file-info {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 15px;
        }

        .alert {
            padding: 15px;
            border-radius: 8px;
            margin: 15px 0;
        }

        .alert-warning {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            color: #856404;
        }

        .alert-success {
            background: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
        }

        .current-file {
            font-weight: bold;
            color: #4facfe;
            margin-bottom: 10px;
        }

        @media (max-width: 768px) {
            .header h1 {
                font-size: 2rem;
            }
            
            .content {
                padding: 20px;
            }
            
            .btn {
                width: 100%;
                margin-right: 0;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔍 СМЕТА МУ</h1>
            <p>Система анализа документов с весовой классификацией</p>
            <p>Поддерживаемые форматы: PDF, DOC, DOCX</p>
        </div>
        
        <div class="content">
            <form id="uploadForm" enctype="multipart/form-data">
                
                <div class="form-group">
                    <label for="files">📄 Выберите файлы для анализа:</label>
                    <input type="file" id="files" name="files" multiple 
                           accept=".pdf,.doc,.docx" required>
                    <small style="color: #666; margin-top: 5px; display: block;">
                        Можно выбрать несколько файлов одновременно (Ctrl+Click)
                    </small>
                </div>
                
                <button type="submit" class="btn btn-primary" id="processBtn">
                    🚀 НАЧАТЬ АНАЛИЗ
                </button>
            </form>
            
            <div class="progress-container" id="progressContainer">
                <div class="current-file" id="currentFile"></div>
                <div class="progress-bar">
                    <div class="progress-fill" id="progressFill" style="width: 0%"></div>
                </div>
            </div>
            
            <div class="log-container" id="logContainer">
                <h3>Ход выполнения:</h3>
                <div class="log-box" id="logBox"></div>
            </div>
            
            <div class="result-container" id="resultContainer">
                <div class="success">✅ Обработка завершена успешно!</div>
                <div class="file-info" id="fileInfo"></div>
                <button class="btn" onclick="downloadResult()">📥 СКАЧАТЬ ОТЧЕТ</button>
                <button class="btn" onclick="startNew()">🔄 НОВЫЙ АНАЛИЗ</button>
            </div>
        </div>
    </div>

    <script>
        let checkInterval;
        
        document.getElementById('uploadForm').addEventListener('submit', function(e) {
            e.preventDefault();
            startProcessing();
        });
        
        function startProcessing() {
            const formData = new FormData();
            const files = document.getElementById('files').files;
            
            if (files.length === 0) {
                alert('Пожалуйста, выберите файлы для обработки');
                return;
            }
            
            
            for (let file of files) {
                formData.append('files', file);
            }
            
            // Показываем прогресс
            document.getElementById('processBtn').disabled = true;
            document.getElementById('progressContainer').style.display = 'block';
            document.getElementById('logContainer').style.display = 'block';
            document.getElementById('resultContainer').style.display = 'none';
            
            // Отправляем файлы
            fetch('/process', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    // Начинаем отслеживать прогресс
                    checkInterval = setInterval(checkStatus, 1000);
                } else {
                    alert('Ошибка: ' + data.error);
                    resetForm();
                }
            })
            .catch(error => {
                alert('Ошибка сети: ' + error);
                resetForm();
            });
        }
        
        function checkStatus() {
            fetch('/status')
            .then(response => response.json())
            .then(data => {
                // Обновляем интерфейс
                document.getElementById('currentFile').textContent = data.current_file;
                document.getElementById('progressFill').style.width = data.progress + '%';
                
                // Обновляем лог
                const logBox = document.getElementById('logBox');
                logBox.innerHTML = data.log.map(msg => '<div>' + msg + '</div>').join('');
                logBox.scrollTop = logBox.scrollHeight;
                
                // Проверяем завершение
                if (!data.active && data.result_file) {
                    clearInterval(checkInterval);
                    showResult(data);
                }
            })
            .catch(error => {
                console.error('Ошибка получения статуса:', error);
            });
        }
        
        function showResult(data) {
            document.getElementById('progressContainer').style.display = 'none';
            document.getElementById('resultContainer').style.display = 'block';
            
            // Показываем информацию о результате
            const lastLog = data.log[data.log.length - 1] || '';
            document.getElementById('fileInfo').innerHTML = 
                '<strong>Файл отчета:</strong> ' + (data.result_file || 'Не создан') + '<br>' +
                '<strong>Статус:</strong> Завершено';
            
            document.getElementById('processBtn').disabled = false;
        }
        
        function downloadResult() {
            window.location.href = '/download';
        }
        
        function startNew() {
            resetForm();
            document.getElementById('files').value = '';
        }
        
        function resetForm() {
            document.getElementById('processBtn').disabled = false;
            document.getElementById('progressContainer').style.display = 'none';
            document.getElementById('logContainer').style.display = 'none';
            document.getElementById('resultContainer').style.display = 'none';
            if (checkInterval) {
                clearInterval(checkInterval);
            }
        }
    </script>
</body>
</html>
'''

@app.route('/')
def index():
    """Главная страница."""
    return render_template_string(HTML_TEMPLATE)

@app.route('/process', methods=['POST'])
def process_files():
    """Начинает обработку файлов."""
    try:
        if processing_status['active']:
            return jsonify({'success': False, 'error': 'Обработка уже выполняется'})
        
        files = request.files.getlist('files')
        author = 'Максим'
        
        if not files or not any(f.filename for f in files):
            return jsonify({'success': False, 'error': 'Файлы не выбраны'})
        
        
        # Сохраняем файлы во временную папку
        temp_dir = tempfile.mkdtemp()
        file_paths = []
        
        for file in files:
            if file.filename:
                file_path = os.path.join(temp_dir, file.filename)
                file.save(file_path)
                file_paths.append(file_path)
        
        # Запускаем обработку в отдельном потоке
        thread = threading.Thread(target=process_files_worker, args=(file_paths, author, temp_dir))
        thread.daemon = True
        thread.start()
        
        return jsonify({'success': True})
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

def process_files_worker(file_paths, author, temp_dir):
    """Обрабатывает файлы в отдельном потоке."""
    global processing_status
    
    try:
        processing_status['active'] = True
        processing_status['progress'] = 0
        processing_status['log'] = []
        processing_status['result_file'] = None
        
        def log_message(msg):
            processing_status['log'].append(msg)
            print(msg)
        
        log_message("🚀 Начинаем обработку файлов...")
        log_message(f"Получено файлов: {len(file_paths)}")
        
        results = {}
        total_files = len(file_paths)
        
        for i, file_path in enumerate(file_paths):
            processing_status['current_file'] = f"Обрабатываю файл {i+1}/{total_files}: {os.path.basename(file_path)}"
            processing_status['progress'] = int((i / total_files) * 90)  # 90% на обработку
            
            log_message(f"--- [{i+1}/{total_files}] {os.path.basename(file_path)} ---")
            
            analyzer = None
            try:
                analyzer = PDFAnalyzer(file_path)
                analyzer.analyze()
                
                df = analyzer.get_summary_dataframe()
                if not df.empty:
                    file_cost = df["Стоимость (руб.)"].sum()
                    
                    results[file_path] = {
                        'cost': file_cost,
                        'pages': len(df),
                        'dataframe': df
                    }
                    
                    log_message(f"✅ Обработан: {file_cost} руб., {len(df)} страниц")
                else:
                    log_message("⚠️ Файл обработан, но результатов нет")
                    
            except Exception as e:
                log_message(f"❌ Ошибка: {str(e)}")
                if analyzer:
                    analyzer.cleanup()
        
        # Сохраняем результаты
        processing_status['current_file'] = "Сохраняю отчет..."
        processing_status['progress'] = 95
        
        if results:
            log_message("📊 Создаю отчет...")
            
            try:
                # Используем стандартную функцию создания отчета БЕЗ перехвата print
                save_batch_report(results, temp_dir, author)
                
                # Находим созданный файл
                from datetime import datetime, timedelta
                now = datetime.now()
                # Используем предыдущий месяц
                prev_month_date = now.replace(day=1) - timedelta(days=1)
                month = prev_month_date.strftime("%m")
                year = now.strftime("%y")
                batch_filename = f"Смета_{author}_{month}_{year}.xlsx"
                output_path = os.path.join(temp_dir, batch_filename)
                
                if os.path.exists(output_path):
                    processing_status['result_file'] = output_path
                    
                    # Считаем статистику
                    total_cost = sum(data['cost'] for data in results.values())
                    total_pages = sum(data['pages'] for data in results.values())
                    
                    log_message(f"✅ Отчет создан: {batch_filename}")
                    log_message(f"📊 Обработано файлов: {len(results)}")
                    log_message(f"📄 Всего страниц: {total_pages}")
                    log_message(f"💰 Общая стоимость: {total_cost} руб.")
                else:
                    log_message(f"⚠️ Файл отчета не найден: {batch_filename}")
                    
            except Exception as save_error:
                log_message(f"❌ Ошибка при создании отчета: {save_error}")
        
        processing_status['progress'] = 100
        processing_status['current_file'] = "Готово!"
        
        # Временные файлы анализаторов очищаются автоматически при выходе из области видимости
                
    except Exception as e:
        log_message(f"❌ Критическая ошибка: {str(e)}")
    finally:
        processing_status['active'] = False

@app.route('/status')
def get_status():
    """Возвращает текущий статус обработки."""
    return jsonify(processing_status)

@app.route('/download')
def download_result():
    """Скачивает результирующий файл."""
    if processing_status['result_file'] and os.path.exists(processing_status['result_file']):
        return send_file(
            processing_status['result_file'], 
            as_attachment=True,
            download_name=os.path.basename(processing_status['result_file'])
        )
    else:
        return "Файл не найден", 404

if __name__ == '__main__':
    print("\n🌐 Запуск веб-интерфейса системы 'Смета МУ'")
    print("📊 Откройте браузер и перейдите по адресу: http://localhost:5003")
    print("🛑 Для остановки нажмите Ctrl+C\n")
    
    app.run(host='0.0.0.0', port=5003, debug=False)