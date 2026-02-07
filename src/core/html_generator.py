"""
Генерация HTML отчетов с библиографией
"""

import time


class HTMLGenerator:
    """Генератор HTML отчетов"""

    @staticmethod
    def generate_html_ordered(items, filename, user_email):
        """Генерирует HTML-страницу с сохранением порядка

        Args:
            items: Список обработанных элементов
            filename: Имя исходного файла
            user_email: Email пользователя

        Returns:
            str: HTML контент
        """
        # Сортируем по оригинальному индексу
        items.sort(key=lambda x: x.get('original_index', 0))

        # Подсчет статистики
        total_items = len(items)
        with_apa = sum(1 for item in items if item.get('apa_citation'))
        no_data = sum(1 for item in items if not item.get('has_data', True))
        manual_dois = sum(1 for item in items if item.get('manual_doi'))

        # Подсчет дублей
        doi_tracker = {}
        duplicates = 0
        for item in items:
            if item['dois'] and len(item['dois']) > 0:
                primary_doi = item['dois'][0]
                if primary_doi in doi_tracker:
                    duplicates += 1
                else:
                    doi_tracker[primary_doi] = True

        html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{filename} — Linx2DOI — Библиография в стиле APA</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    line-height: 1.6;
    color: #333;
    background: #f8fafc;
    padding: 20px;
}}
.container {{
    max-width: 800px;
    margin: 0 auto;
    background: white;
    border-radius: 10px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    padding: 30px;
}}
.header {{
    text-align: center;
    margin-bottom: 30px;
    padding-bottom: 20px;
    border-bottom: 2px solid #e2e8f0;
}}
h1 {{
    color: #1a365d;
    margin-bottom: 10px;
}}
.stats {{
    display: flex;
    justify-content: space-around;
    margin: 20px 0;
    padding: 15px;
    background: #f1f5f9;
    border-radius: 8px;
}}
.stat-item {{
    text-align: center;
}}
.stat-value {{
    font-size: 24px;
    font-weight: bold;
    color: #4361ee;
}}
.stat-label {{
    color: #64748b;
    font-size: 14px;
}}
.item {{
    margin-bottom: 25px;
    padding: 20px;
    border-radius: 8px;
    border-left: 5px solid #63b3ed;
    background: #fff;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}}
.item.no-data {{
    border-left-color: #dc3545;
    background: #f8d7da;
}}
.item.manual-doi {{
    border-left-color: #10b981;
    background: #d1fae5;
}}
.item-number {{
    font-weight: bold;
    color: #4361ee;
    font-size: 18px;
    margin-bottom: 10px;
}}
.item-header {{
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    margin-bottom: 10px;
}}
.item-title {{
    font-size: 16px;
    font-weight: 600;
    color: #1a365d;
    flex: 1;
}}
.item-title a {{
    color: #1a365d;
    text-decoration: none;
    border-bottom: 2px solid transparent;
    transition: border-color 0.2s;
}}
.item-title a:hover {{
    border-bottom-color: #4361ee;
}}
.item-url {{
    font-size: 14px;
    color: #4a5568;
    margin-bottom: 10px;
    word-break: break-all;
}}
.no-data-badge {{
    background: #dc3545;
    color: white;
    padding: 5px 10px;
    border-radius: 4px;
    font-size: 12px;
    font-weight: bold;
    margin-left: 10px;
}}
.manual-doi-badge {{
    background: #10b981;
    color: white;
    padding: 5px 10px;
    border-radius: 4px;
    font-size: 12px;
    font-weight: bold;
    margin-left: 10px;
}}
.duplicate-badge {{
    background: #f59e0b;
    color: white;
    padding: 5px 10px;
    border-radius: 4px;
    font-size: 12px;
    font-weight: bold;
    margin-left: 10px;
}}
.item.duplicate {{
    border-left-color: #f59e0b;
    background: #fef3c7;
}}
.doi-container {{
    margin: 10px 0;
}}
.doi-link {{
    display: inline-block;
    background: #dbeafe;
    color: #1e40af;
    padding: 5px 10px;
    border-radius: 4px;
    text-decoration: none;
    margin-right: 10px;
    margin-bottom: 5px;
    font-size: 14px;
}}
.doi-link:hover {{
    background: #bfdbfe;
}}
.apa-citation {{
    font-family: Georgia, serif;
    color: #2d3748;
    font-size: 15px;
    line-height: 1.6;
    margin-top: 15px;
    padding-top: 15px;
    border-top: 1px dashed #cbd5e1;
}}
.apa-citation i {{
    font-style: italic;
}}
.apa-citation b {{
    font-weight: bold;
}}
.no-citation {{
    color: #dc3545;
    font-style: italic;
    font-size: 14px;
    margin-top: 15px;
    padding: 10px;
    background: #fef2f2;
    border-radius: 4px;
}}
.footer {{
    text-align: center;
    margin-top: 30px;
    padding-top: 20px;
    border-top: 1px solid #e2e8f0;
    color: #64748b;
    font-size: 14px;
}}
</style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>📚 Библиография в стиле APA</h1>
        <p><strong>Файл:</strong> {filename}</p>
        <p><strong>Email:</strong> {user_email}</p>
        <p><strong>Сохранена оригинальная последовательность</strong></p>
    </div>

    <div class="stats">
        <div class="stat-item">
            <div class="stat-value">{total_items}</div>
            <div class="stat-label">Всего записей</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">{with_apa}</div>
            <div class="stat-label">С цитатами</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">{no_data}</div>
            <div class="stat-label">Без данных</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">{manual_dois}</div>
            <div class="stat-label">Ручных DOI</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">{duplicates}</div>
            <div class="stat-label">Дублей</div>
        </div>
    </div>
"""

        # Словарь для отслеживания дублей DOI в HTML
        doi_html_tracker = {}

        for item in items:
            item_class = "item"
            badge_html = ""

            # Проверка на дубли DOI
            is_duplicate = False
            duplicate_index = None
            if item['dois'] and len(item['dois']) > 0:
                primary_doi = item['dois'][0]
                if primary_doi in doi_html_tracker:
                    is_duplicate = True
                    duplicate_index = doi_html_tracker[primary_doi]
                else:
                    doi_html_tracker[primary_doi] = item.get('original_index', 0)

            if not item.get('has_data', True):
                item_class += " no-data"
                badge_html = """<span class="no-data-badge">НЕТ ДАННЫХ</span>"""
            elif is_duplicate:
                item_class += " duplicate"
                badge_html = f"""<span class="duplicate-badge">ДУБЛЬ № {duplicate_index}</span>"""
            elif item.get('manual_doi'):
                item_class += " manual-doi"
                badge_html = """<span class="manual-doi-badge">РУЧНОЙ DOI</span>"""

            # Формируем заголовок статьи как гиперссылку на оригинальный URL
            # Если article_title пустой или None, используем оригинальный title
            article_title = item.get('article_title') or item.get('title', 'Без названия')
            original_url = item.get('url', '')
            if original_url:
                # Используем оригинальный URL из документа
                title_html = f"""<a href="{original_url}" target="_blank" style="color: #1a365d; text-decoration: none;">{article_title}</a>"""
            else:
                title_html = article_title

            html += f"""
    <div class="{item_class}">
        <div class="item-number">{item.get('original_index', 0)}.</div>
        <div class="item-header">
            <div class="item-title">{title_html}</div>
            {badge_html}
        </div>
"""

            if item['dois']:
                html += """        <div class="doi-container">"""
                for doi in item['dois']:
                    doi_url = f"https://doi.org/{doi}"
                    html += f"""<a href="{doi_url}" class="doi-link" target="_blank">🔗 doi:{doi}</a>"""
                html += """</div>"""

            if item.get('apa_citation'):
                html += f"""        <div class="apa-citation">{item['apa_citation']}</div>"""
            else:
                html += f"""        <div class="no-citation">⚠️ Библиографическая ссылка недоступна</div>"""

            html += """    </div>"""

        html += f"""
    <div class="footer">
        <p>📄 Сгенерировано: {time.strftime("%d.%m.%Y %H:%M")}</p>
        <p>Позиции сохранены как в исходном файле</p>
    </div>
</div>
</body>
</html>"""
        return html

    @staticmethod
    def generate_table_data(items):
        """Генерирует данные для таблицы результатов

        Args:
            items: Список обработанных элементов

        Returns:
            list: Список строк для таблицы
        """
        # Словарь для отслеживания первых вхождений DOI
        doi_first_occurrence = {}

        table_data = []
        for idx, item in enumerate(items):
            # Формируем статус
            status_parts = []
            is_duplicate = False
            duplicate_index = None

            # Проверка на дубли DOI
            if item['dois'] and len(item['dois']) > 0:
                primary_doi = item['dois'][0]  # Берем первый DOI
                if primary_doi in doi_first_occurrence:
                    # Это дубликат
                    is_duplicate = True
                    duplicate_index = doi_first_occurrence[primary_doi]
                else:
                    # Первое вхождение
                    doi_first_occurrence[primary_doi] = item.get('original_index', idx + 1)

            if not item.get('has_data', True):
                status_parts.append("НЕТ ДАННЫХ")
            elif is_duplicate:
                status_parts.append(f"ДУБЛЬ № {duplicate_index}")
            elif item.get('manual_doi'):
                status_parts.append("РУЧНОЙ DOI")

            status_str = " | ".join(status_parts) if status_parts else "ОК"

            # DOI строка
            dois_str = ", ".join(item['dois']) if item['dois'] else "Данные отсутствуют"

            # URL для заголовка статьи
            doi_url = f"https://doi.org/{item['dois'][0]}" if item['dois'] and len(item['dois']) > 0 else ""

            table_data.append([
                str(item.get('original_index', 0)),  # № в оригинале
                item['title'],  # Название
                item.get('article_title', 'Не найден'),  # Заголовок статьи
                dois_str,  # DOI
                doi_url,  # URL для заголовка статьи
                item.get('manual_doi', ''),  # Ручной DOI
                status_str,  # Статус
                item  # Оригинальный объект
            ])
        return table_data
