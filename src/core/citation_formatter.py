"""
Форматирование библиографических цитат в стиле APA
"""

import time
import requests

from .config import CROSSREF_WORKS, MAX_RETRIES, REQUEST_TIMEOUT


class CitationFormatter:
    """Класс для получения и форматирования цитат в APA стиле"""

    def __init__(self, user_email, log_callback=None):
        """
        Args:
            user_email: Email пользователя для API запросов
            log_callback: Функция обратного вызова для логирования
        """
        self.user_email = user_email
        self.log_callback = log_callback
        self.headers = {
            'User-Agent': f'DOIExtractor/1.0 (mailto:{user_email})',
            'Accept': 'application/json',
        }

    def log(self, message):
        """Логирование сообщений"""
        if self.log_callback:
            self.log_callback(message)

    @staticmethod
    def format_authors(authors):
        """Форматирует список авторов в APA стиль

        Args:
            authors: Список словарей с ключами 'family' и 'given'

        Returns:
            str: Отформатированная строка авторов
        """
        if not authors:
            return "Author unknown"

        formatted = []
        for author in authors[:7]:
            family = author.get('family', '').strip()
            given = author.get('given', '').strip()

            if family and given:
                initials = ''.join([part[0].upper() + '.' for part in given.split() if part])
                formatted.append(f"{family}, {initials}")
            elif family:
                formatted.append(family)

        if len(authors) > 7:
            return '; '.join(formatted[:6]) + '; ... ' + formatted[-1] if formatted else "Author unknown"
        elif len(formatted) == 0:
            return "Author unknown"
        elif len(formatted) == 1:
            return formatted[0]
        elif len(formatted) == 2:
            return f"{formatted[0]} & {formatted[1]}"
        else:
            return '; '.join(formatted[:-1]) + f"; & {formatted[-1]}"

    def get_article_title_from_doi(self, doi):
        """Получает заголовок статьи по DOI через CrossRef

        Args:
            doi: DOI статьи

        Returns:
            str или None: Заголовок статьи или None
        """
        try:
            url = f"{CROSSREF_WORKS}/{doi}"

            for attempt in range(MAX_RETRIES - 1):  # 2 попытки
                try:
                    response = requests.get(
                        url,
                        headers=self.headers,
                        timeout=REQUEST_TIMEOUT
                    )

                    if response.status_code == 200:
                        data = response.json()
                        if data.get('status') != 'ok' or 'message' not in data:
                            self.log(f"  ❌ Ошибка получения заголовка для DOI: {doi}")
                            return None

                        msg = data['message']
                        title = msg.get('title', [''])[0].strip() if msg.get('title') else None

                        if title:
                            self.log(f"  📝 Заголовок статьи: {title[:60]}...")

                        return title

                    elif response.status_code == 429:
                        wait = 2 ** attempt
                        time.sleep(wait)
                        continue
                    else:
                        self.log(f"  ❌ Ошибка HTTP {response.status_code} при получении заголовка")
                        return None

                except requests.exceptions.RequestException as e:
                    self.log(f"  ❌ Ошибка запроса заголовка: {str(e)}")
                    if attempt < MAX_RETRIES - 2:
                        time.sleep(2)
                        continue
                    return None

            return None
        except Exception as e:
            self.log(f"  ❌ Исключение при получении заголовка: {str(e)}")
            return None

    def get_apa_citation(self, doi):
        """Получает полную APA цитату через CrossRef

        Args:
            doi: DOI статьи

        Returns:
            str или None: APA цитата в HTML формате или None
        """
        try:
            self.log(f"  📚 Запрос CrossRef для DOI: {doi[:30]}...")
            url = f"{CROSSREF_WORKS}/{doi}"

            for attempt in range(MAX_RETRIES - 1):  # 2 попытки
                try:
                    response = requests.get(
                        url,
                        headers=self.headers,
                        timeout=REQUEST_TIMEOUT
                    )

                    if response.status_code == 200:
                        data = response.json()
                        if data.get('status') != 'ok' or 'message' not in data:
                            self.log(f"  ❌ Ошибка в ответе CrossRef для DOI: {doi}")
                            return None

                        msg = data['message']

                        # Форматируем авторов
                        authors = msg.get('author', [])
                        author_str = self.format_authors(authors)

                        # Получаем год
                        year = ''
                        for field in ['published-print', 'published-online', 'created']:
                            date_parts = msg.get(field, {}).get('date-parts', [])
                            if date_parts and len(date_parts[0]) > 0:
                                year = str(date_parts[0][0])
                                break

                        if not year:
                            year = 'n.d.'

                        # Заголовок
                        title = msg.get('title', [''])[0].strip() if msg.get('title') else 'Title unknown'
                        title = title[0].upper() + title[1:] if title else 'Title unknown'

                        # Журнал
                        journal = msg.get('container-title', [''])[0].strip() if msg.get(
                            'container-title') else 'Journal unknown'

                        # Том, выпуск, страницы
                        volume = msg.get('volume', '')
                        issue = msg.get('issue', '')
                        pages = msg.get('page', '')

                        # Собираем цитату
                        citation = f"{author_str} ({year}). {title}. <i>{journal}</i>"

                        if volume:
                            citation += f", <b>{volume}</b>"
                        if issue:
                            citation += f"({issue})"
                        if pages:
                            pages = pages.replace('-', '–')
                            citation += f", {pages}"

                        citation += f". https://doi.org/{doi}"

                        self.log(f"  ✅ APA цитата получена")
                        return citation

                    elif response.status_code == 429:
                        wait = 2 ** attempt
                        self.log(f"  ⚠️ Rate limit, ожидание {wait} сек...")
                        time.sleep(wait)
                        continue
                    else:
                        self.log(f"  ❌ Ошибка HTTP {response.status_code} при запросе CrossRef")
                        return None

                except requests.exceptions.RequestException as e:
                    self.log(f"  ❌ Ошибка запроса CrossRef: {str(e)}")
                    if attempt < MAX_RETRIES - 2:
                        time.sleep(2)
                        continue
                    return None

            self.log(f"  ❌ Превышено количество попыток для DOI: {doi}")
            return None

        except Exception as e:
            self.log(f"  ❌ Исключение при получении APA: {str(e)}")
            return None
