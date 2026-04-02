"""
Генерация HTML отчетов с библиографией
"""

import time


class HTMLGenerator:
    """Генератор HTML отчетов"""

    @staticmethod
    def generate_html_ordered(items, filename, user_email):
        """Генерирует HTML-страницу с сохранением порядка"""
        items.sort(key=lambda x: x.get('original_index', 0))

        total_items = len(items)
        with_apa = sum(1 for item in items if item.get('apa_citation'))
        no_data = sum(1 for item in items if not item.get('has_data', True))
        manual_dois = sum(1 for item in items if item.get('manual_doi'))

        doi_tracker = {}
        duplicates = 0
        for item in items:
            if item['dois'] and len(item['dois']) > 0:
                primary_doi = item['dois'][0]
                if primary_doi in doi_tracker:
                    duplicates += 1
                else:
                    doi_tracker[primary_doi] = True

        uncited_items = sum(1 for item in items if not item.get('is_cited', True))

        html = f"""
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>📚 Библиография в стиле APA</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        .item {{ margin: 15px 0; padding: 10px; border-left: 3px solid #63b3ed; }}
        .item.no-data {{ border-left-color: #dc3545; }}
        .item.manual-doi {{ border-left-color: #10b981; }}
        .item.duplicate {{ border-left-color: #f59e0b; }}
        .item.uncited {{ border-left-color: #9e9e9e; }}
        .badge {{ display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 12px; color: white; }}
        .badge-no-data {{ background-color: #dc3545; }}
        .badge-manual {{ background-color: #10b981; }}
        .badge-duplicate {{ background-color: #f59e0b; }}
        .badge-uncited {{ background-color: #9e9e9e; }}
        .doi-link {{ color: #0366d6; text-decoration: none; }}
        .citation {{ font-style: italic; color: #555; }}
        .stats {{ background: #f6f8fa; padding: 15px; border-radius: 5px; margin-bottom: 20px; }}
    </style>
</head>
<body>
    <h1>📚 Библиография в стиле APA</h1>
    <div class="stats">
        <strong>Файл:</strong> {filename}<br>
        <strong>Email:</strong> {user_email}<br>
        <strong>Сохранена оригинальная последовательность</strong><br><br>
        <strong>{total_items}</strong> Всего записей |
        <strong>{with_apa}</strong> С цитатами |
        <strong>{no_data}</strong> Без данных |
        <strong>{manual_dois}</strong> Ручных DOI/ISBN |
        <strong>{duplicates}</strong> Дублей |
        <strong>{uncited_items}</strong> Не цитируются
    </div>
"""

        doi_html_tracker = {}

        for item in items:
            item_class = "item"
            badge_html = ""

            is_duplicate = False
            duplicate_index = None
            if item['dois'] and len(item['dois']) > 0:
                primary_doi = item['dois'][0]
                if primary_doi in doi_html_tracker:
                    is_duplicate = True
                    duplicate_index = doi_html_tracker[primary_doi]
                else:
                    doi_html_tracker[primary_doi] = item.get('original_index', 0)

            if not item.get('is_cited', True):
                item_class += " uncited"
                badge_html = """<span class="badge badge-uncited">ОТСУТСТВУЕТ В ТЕКСТЕ</span>"""
            elif not item.get('has_data', True):
                item_class += " no-data"
                badge_html = """<span class="badge badge-no-data">НЕТ ДАННЫХ</span>"""
            elif is_duplicate:
                item_class += " duplicate"
                badge_html = f"""<span class="badge badge-duplicate">ДУБЛЬ № {duplicate_index}</span>"""
            elif item.get('manual_doi'):
                item_class += " manual-doi"
                badge_html = """<span class="badge badge-manual">РУЧНОЙ DOI/ISBN</span>"""

            article_title = item.get('article_title') or item.get('title', 'Без названия')
            original_url = item.get('url', '')
            if original_url:
                title_html = f"""<a href="{original_url}" class="doi-link">{article_title}</a>"""
            else:
                title_html = article_title

            html += f"""
    <div class="{item_class}">
        <strong>{item.get('original_index', 0)}.</strong> {title_html}
        {badge_html}
"""

            # ✅ DOI или ISBN
            if item['dois'] and len(item['dois']) > 0:
                html += """        <div>"""
                for doi in item['dois']:
                    doi_url = f"https://doi.org/{doi}"
                    html += f"""
            <a href="{doi_url}" class="doi-link">🔗 doi:{doi}</a>
"""
                html += """        </div>
"""
            elif item.get('isbn'):
                html += f"""        <div>📚 ISBN: {item['isbn']}</div>
"""

            if item.get('apa_citation'):
                html += f"""        <div class="citation">{item['apa_citation']}</div>"""
            else:
                html += """        <div class="citation">⚠️ Библиографическая ссылка недоступна</div>"""

            html += """    </div>
"""

        html += f"""
    <hr>
    <p><small>📄 Сгенерировано: {time.strftime("%d.%m.%Y %H:%M")} | Позиции сохранены как в исходном файле</small></p>
</body>
</html>
"""
        return html

    @staticmethod
    def generate_table_data(items):
        """Генерирует данные для таблицы результатов"""
        doi_first_occurrence = {}

        table_data = []
        for idx, item in enumerate(items):
            status_parts = []
            is_duplicate = False
            duplicate_index = None

            if item['dois'] and len(item['dois']) > 0:
                primary_doi = item['dois'][0]
                if primary_doi in doi_first_occurrence:
                    is_duplicate = True
                    duplicate_index = doi_first_occurrence[primary_doi]
                else:
                    doi_first_occurrence[primary_doi] = item.get('original_index', idx + 1)

            if not item.get('is_cited', True):
                status_parts.append("ОТСУТСТВУЕТ В ТЕКСТЕ")
            elif not item.get('has_data', True):
                status_parts.append("НЕТ ДАННЫХ")
            elif is_duplicate:
                status_parts.append(f"ДУБЛЬ № {duplicate_index}")
            elif item.get('manual_doi'):
                status_parts.append("РУЧНОЙ DOI/ISBN")

            status_str = " | ".join(status_parts) if status_parts else "ОК"

            # ✅ ИЗМЕНЕНО: DOI столбец - приоритет DOI, затем ISBN
            if item['dois'] and len(item['dois']) > 0:
                dois_str = ", ".join(item['dois'])
            elif item.get('isbn'):
                dois_str = f"ISBN: {item['isbn']}"
            else:
                dois_str = "Данные отсутствуют"

            doi_url = f"https://doi.org/{item['dois'][0]}" if item['dois'] and len(item['dois']) > 0 else ""

            table_data.append([
                str(item.get('original_index', 0)),
                item['title'],
                item.get('article_title', 'Не найден'),
                dois_str,
                doi_url,
                item.get('manual_doi', ''),
                status_str,
                item
            ])
        return table_data