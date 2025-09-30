from flask import Flask, render_template, send_file, send_from_directory
import os
import glob
from datetime import datetime
import subprocess
import sys

app = Flask(__name__)

import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


@app.route('/')
def index():
    """Главная страница с графиками"""
    chart_files = glob.glob('charts/*.png')
    charts = []

    for chart_file in chart_files:
        chart_name = os.path.basename(chart_file)
        chart_type = chart_name.replace('.png', '').replace('_', ' ').title()
        charts.append({
            'filename': chart_name,
            'name': chart_type,
            'created_time': datetime.fromtimestamp(os.path.getctime(chart_file)).strftime('%Y-%m-%d %H:%M:%S')
        })

    interactive_charts = glob.glob('charts/*.html')
    interactive_list = []

    for chart_file in interactive_charts:
        chart_name = os.path.basename(chart_file)
        chart_type = chart_name.replace('.html', '').replace('_', ' ').title()
        interactive_list.append({
            'filename': chart_name,
            'name': chart_type
        })

    excel_files = glob.glob('exports/*.xlsx')
    exports = []

    for excel_file in excel_files:
        exports.append({
            'filename': os.path.basename(excel_file),
            'size': f"{os.path.getsize(excel_file) / 1024:.1f} KB"
        })

    return render_template('index.html',
                           charts=charts,
                           interactive_charts=interactive_list,
                           exports=exports)


@app.route('/charts/<filename>')
def serve_chart(filename):
    """Отдача файлов графиков"""
    return send_file(f'charts/{filename}')


@app.route('/interactive/<filename>')
def serve_interactive_chart(filename):
    """Отдача интерактивных HTML графиков"""
    return send_from_directory('charts', filename)


@app.route('/exports/<filename>')
def serve_export(filename):
    """Отдача Excel файлов"""
    return send_file(f'exports/{filename}')


@app.route('/run-analysis')
def run_analysis():
    """Запуск анализа и обновление страницы"""
    try:
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'

        result = subprocess.run(
            [sys.executable, "analytics.py"],
            capture_output=True,
            text=True,
            cwd=os.getcwd(),
            env=env,
            encoding='utf-8',
            errors='replace'
        )

        output = result.stdout if result.stdout else "Нет вывода"
        error = result.stderr if result.stderr else "Нет ошибок"

        if result.returncode == 0:
            return f"""
            <html>
            <head>
                <meta charset="UTF-8">
                <title>Анализ завершен</title>
            </head>
            <body style="font-family: Arial; padding: 20px;">
                <h2 style="color: green;">Анализ успешно выполнен!</h2>
                <h3>Вывод программы:</h3>
                <pre style="background: #f5f5f5; padding: 15px; border-radius: 5px; white-space: pre-wrap;">{output}</pre>
                <a href='/' style="background: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Вернуться на главную</a>
            </body>
            </html>
            """
        else:
            return f"""
            <html>
            <head>
                <meta charset="UTF-8">
                <title>Ошибка анализа</title>
            </head>
            <body style="font-family: Arial; padding: 20px;">
                <h2 style="color: red;">Ошибка при выполнении анализа</h2>
                <h3>Вывод программы:</h3>
                <pre style="background: #f5f5f5; padding: 15px; border-radius: 5px; white-space: pre-wrap;">{output}</pre>
                <h3>Ошибки:</h3>
                <pre style="background: #ffe6e6; padding: 15px; border-radius: 5px; white-space: pre-wrap; color: red;">{error}</pre>
                <a href='/' style="background: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Вернуться на главную</a>
            </body>
            </html>
            """
    except Exception as e:
        return f"""
        <html>
        <head>
            <meta charset="UTF-8">
        </head>
        <body style="font-family: Arial; padding: 20px;">
            <h2 style="color: red;">Ошибка: {e}</h2>
            <a href='/' style="background: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Вернуться на главную</a>
        </body>
        </html>
        """


@app.route('/create-interactive-category')
def create_interactive_category():
    """Создание только интерактивной диаграммы категорий"""
    try:
        from analytics import TicketSalesAnalyzer
        analyzer = TicketSalesAnalyzer()

        analyzer.create_interactive_category_sales()
        analyzer.create_advanced_interactive_dashboard()

        return """
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Интерактивные графики созданы</title>
        </head>
        <body style="font-family: Arial; padding: 20px;">
            <h2 style="color: green;">Интерактивные графики созданы!</h2>
            <p>Были созданы следующие интерактивные диаграммы:</p>
            <ul>
                <li><a href="/interactive/interactive_category_sales.html" target="_blank">Динамика продаж по категориям</a></li>
                <li><a href="/interactive/advanced_sales_dashboard.html" target="_blank">Продвинутая панель продаж</a></li>
            </ul>
            <p>Нажмите на ссылки выше чтобы открыть интерактивные графики в новой вкладке.</p>
            <a href='/' style="background: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; margin-right: 10px;">Вернуться на главную</a>
            <a href='/interactive-charts' style="background: #28a745; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Просмотреть все интерактивные графики</a>
        </body>
        </html>
        """
    except Exception as e:
        return f"""
        <html>
        <head>
            <meta charset="UTF-8">
        </head>
        <body style="font-family: Arial; padding: 20px;">
            <h2 style="color: red;">Ошибка: {e}</h2>
            <a href='/' style="background: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Вернуться на главную</a>
        </body>
        </html>
        """


@app.route('/interactive-charts')
def interactive_charts():
    """Страница со всеми интерактивными графиками"""
    interactive_files = glob.glob('charts/*.html')
    charts = []

    for chart_file in interactive_files:
        chart_name = os.path.basename(chart_file)
        chart_display_name = chart_name.replace('.html', '').replace('_', ' ').title()
        charts.append({
            'filename': chart_name,
            'name': chart_display_name
        })

    return render_template('interactive_charts.html', charts=charts)


if __name__ == '__main__':
    os.makedirs('charts', exist_ok=True)
    os.makedirs('exports', exist_ok=True)
    os.makedirs('templates', exist_ok=True)

    print("Запуск веб-сервера на http://127.0.0.1:56777")
    app.run(host='127.0.0.1', port=56777, debug=False)