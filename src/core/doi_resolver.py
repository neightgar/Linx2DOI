"""
Разрешение DOI через различные источники
"""

import re
import time
import requests
import xml.etree.ElementTree as ET
from difflib import SequenceMatcher
from urllib.parse import urlparse, unquote

from .config import ENTREZ_EFETCH, CROSSREF_WORKS, MAX_RETRIES, REQUEST_TIMEOUT
from .document_parser import DocumentParser


class DOIResolver:
    """Класс для поиска DOI через различные источники"""

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

    def extract_doi_from_text(self, text):
        """Извлекает DOI из текста используя regex паттерны"""
        dois = set()
        doi_pattern = r'10\.\d{4,9}/[-._;()/:A-Z0-9]+'

        # Паттерн: doi: 10.xxxx/xxxx
        for match in re.finditer(r'doi:\s*(' + doi_pattern + r')', text, re.IGNORECASE):
            dois.add(match.group(1).strip())

        # Паттерн: https://doi.org/10.xxxx/xxxx
        for match in re.finditer(r'https?://doi\.org/(' + doi_pattern + r')', text, re.IGNORECASE):
            dois.add(match.group(1).strip())

        # Паттерн: просто DOI в тексте
        for match in re.finditer(doi_pattern, text):
            candidate = match.group(0).strip()
            if candidate.startswith('10.'):
                dois.add(candidate)

        return sorted(dois)

    def extract_doi_from_doi_url(self, url):
        """Извлекает DOI из URL вида doi.org/10.xxxx/xxxx"""
        if 'doi.org' in url.lower():
            match = re.search(r'10\.\d{4,9}/[-._;()/:A-Z0-9]+', url)
            if match:
                return match.group(0)
        return None

    def extract_pmid_from_url(self, url):
        """Извлекает PMID из PubMed URL"""
        parsed = urlparse(url)
        if 'pubmed.ncbi.nlm.nih.gov' in parsed.netloc:
            match = re.search(r'/(\d+)/?$', parsed.path)
            if match:
                return match.group(1)
        return None

    def extract_pmcid_from_url(self, url):
        """Извлекает PMCID из PMC URL"""
        parsed = urlparse(url)
        if 'pmc.ncbi.nlm.nih.gov' in parsed.netloc:
            match = re.search(r'/articles/(PMC\d+)/?$', parsed.path, re.IGNORECASE)
            if match:
                return match.group(1)
        return None

    def get_doi_from_pubmed_api(self, pmid):
        """Получает DOI через PubMed API по PMID"""
        try:
            params = {
                'db': 'pubmed',
                'id': pmid,
                'retmode': 'xml',
                'email': self.user_email,
                'tool': 'DOIExtractor'
            }

            for attempt in range(MAX_RETRIES):
                try:
                    response = requests.get(
                        ENTREZ_EFETCH,
                        params=params,
                        headers=self.headers,
                        timeout=REQUEST_TIMEOUT
                    )
                    if response.status_code == 200:
                        root = ET.fromstring(response.content)

                        # Поиск DOI в элементах ELocationID
                        for elocation in root.findall(".//ELocationID[@EIdType='doi']"):
                            doi = elocation.text.strip()
                            if doi.startswith('10.'):
                                return doi

                        # Поиск DOI в элементах ArticleId
                        for aid in root.findall(".//ArticleId[@IdType='doi']"):
                            doi = aid.text.strip()
                            if doi.startswith('10.'):
                                return doi

                        return None

                    elif response.status_code == 429:
                        wait = 2 ** attempt
                        time.sleep(wait)
                        continue
                    else:
                        return None

                except (requests.exceptions.RequestException, ET.ParseError):
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(2)
                        continue
                    return None

            return None
        except Exception:
            return None

    def get_doi_from_pmc_api(self, pmcid):
        """Получает DOI через PMC API по PMCID"""
        try:
            params = {
                'db': 'pmc',
                'id': pmcid,
                'retmode': 'xml',
                'email': self.user_email,
                'tool': 'DOIExtractor'
            }

            for attempt in range(MAX_RETRIES):
                try:
                    response = requests.get(
                        ENTREZ_EFETCH,
                        params=params,
                        headers=self.headers,
                        timeout=REQUEST_TIMEOUT
                    )
                    if response.status_code == 200:
                        root = ET.fromstring(response.content)

                        for article_id in root.findall(".//article-id[@pub-id-type='doi']"):
                            doi = article_id.text.strip()
                            if doi.startswith('10.'):
                                return doi

                        return None

                    elif response.status_code == 429:
                        wait = 2 ** attempt
                        time.sleep(wait)
                        continue
                    else:
                        return None

                except (requests.exceptions.RequestException, ET.ParseError):
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(2)
                        continue
                    return None

            return None
        except Exception:
            return None

    def extract_researchgate_title(self, url):
        """Извлекает заголовок из структуры URL ResearchGate"""
        try:
            parsed = urlparse(url)
            if 'researchgate.net' not in parsed.netloc:
                return None

            path = unquote(parsed.path)
            match = re.search(r'/publication/\d+_(.+?)(?:/|$)', path)
            if not match:
                return None

            title_raw = match.group(1)
            title_clean = title_raw.replace('_', ' ')
            title_clean = re.sub(r'[^a-zA-Z0-9\s\-.,:;]', ' ', title_clean)
            title_clean = re.sub(r'\s+', ' ', title_clean).strip()

            if len(title_clean) < 20 or len(title_clean) > 150:
                return None

            self.log(f"  🔍 Извлечен заголовок из ResearchGate: {title_clean[:50]}...")
            return title_clean

        except Exception as e:
            self.log(f"  ❌ Ошибка извлечения заголовка из ResearchGate: {str(e)}")
            return None

    def search_doi_via_crossref(self, title, url=None):
        """Поиск DOI через CrossRef API по заголовку статьи"""
        try:
            self.log(f"  🔍 Поиск DOI по заголовку: {title[:50]}...")
            clean_title = re.sub(r'[^\w\s\-.,:;]', '', title)

            if len(clean_title) < 15:
                self.log(f"  ⚠️ Заголовок слишком короткий для поиска")
                return None

            if len(clean_title) > 100:
                clean_title = clean_title[:100].rsplit(' ', 1)[0]

            params = {
                'query.bibliographic': clean_title,
                'rows': 3,
                'select': 'DOI,title'
            }

            for attempt in range(MAX_RETRIES - 1):  # 2 попытки
                try:
                    response = requests.get(
                        CROSSREF_WORKS,
                        params=params,
                        headers=self.headers,
                        timeout=REQUEST_TIMEOUT
                    )

                    if response.status_code == 200:
                        data = response.json()
                        if data.get('status') == 'ok' and 'message' in data and data['message'].get('items'):
                            items = data['message']['items']
                            self.log(f"  📊 Найдено {len(items)} возможных совпадений")

                            for item in items:
                                doi = item.get('DOI', '').strip()
                                if not doi.startswith('10.'):
                                    continue

                                found_titles = item.get('title', [])
                                if not found_titles:
                                    continue

                                found_title = found_titles[0].lower() if found_titles else ''
                                clean_found = DocumentParser.normalize_title(found_title)
                                clean_search = DocumentParser.normalize_title(clean_title)

                                # Более гибкое сравнение заголовков
                                # 1. Прямое совпадение первых 25 символов (двунаправленное)
                                if clean_search[:25] in clean_found or clean_found[:25] in clean_search:
                                    self.log(f"  ✅ Найден DOI: {doi}")
                                    return doi

                                # 2. Проверка на короткие заголовки
                                if len(clean_search) < 30:
                                    self.log(f"  ✅ Найден DOI: {doi}")
                                    return doi

                                # 3. Проверка сходства через SequenceMatcher
                                similarity = SequenceMatcher(None, clean_search, clean_found).ratio()
                                if similarity >= 0.75:  # 75% сходства
                                    self.log(f"  ✅ Найден DOI (сходство: {similarity:.2f}): {doi}")
                                    return doi

                        self.log(f"  ❌ DOI по заголовку не найден")
                        return None

                    elif response.status_code == 429:
                        wait = 2 ** attempt
                        self.log(f"  ⚠️ Rate limit, ожидание {wait} сек...")
                        time.sleep(wait)
                        continue

                except requests.exceptions.RequestException as e:
                    self.log(f"  ❌ Ошибка поиска DOI: {str(e)}")
                    if attempt < MAX_RETRIES - 2:
                        time.sleep(2)
                        continue

            return None
        except Exception as e:
            self.log(f"  ❌ Исключение при поиске DOI: {str(e)}")
            return None

    def resolve_doi(self, item, manual_doi=None, auto_search=True):
        """Разрешает DOI для элемента используя различные стратегии

        Args:
            item: Словарь с данными статьи (title, url, text)
            manual_doi: Ручной DOI (приоритетный)
            auto_search: Выполнять автоматический поиск

        Returns:
            list: Список найденных DOI
        """
        self.log(f"\n📄 Обработка записи #{item['original_index']}: {item['title'][:50]}...")
        dois = set()

        # Ручной DOI имеет наивысший приоритет
        if manual_doi and manual_doi.strip():
            self.log(f"  ✅ Используется ручной DOI: {manual_doi}")
            dois.add(manual_doi.strip())
            return sorted(dois)

        if not auto_search:
            return []

        # Автоматический поиск DOI
        self.log(f"  🔍 Поиск DOI в тексте...")
        text_dois = self.extract_doi_from_text(item['text'])
        if text_dois:
            dois.update(text_dois)
            self.log(f"  ✅ Найдено DOI в тексте: {', '.join(text_dois)}")

        if not dois:
            self.log(f"  🔍 Поиск DOI в URL...")
            doi_from_url = self.extract_doi_from_doi_url(item['url'])
            if doi_from_url:
                dois.add(doi_from_url)
                self.log(f"  ✅ DOI извлечен из URL: {doi_from_url}")

        if not dois:
            self.log(f"  🔍 Проверка PubMed...")
            pmid = self.extract_pmid_from_url(item['url'])
            if pmid:
                self.log(f"  📊 Найден PMID: {pmid}")
                doi_from_pubmed = self.get_doi_from_pubmed_api(pmid)
                if doi_from_pubmed:
                    dois.add(doi_from_pubmed)
                    self.log(f"  ✅ PubMed DOI: {doi_from_pubmed}")

        if not dois:
            self.log(f"  🔍 Проверка PMC...")
            pmcid = self.extract_pmcid_from_url(item['url'])
            if pmcid:
                self.log(f"  📊 Найден PMCID: {pmcid}")
                doi_from_pmc = self.get_doi_from_pmc_api(pmcid)
                if doi_from_pmc:
                    dois.add(doi_from_pmc)
                    self.log(f"  ✅ PMC DOI: {doi_from_pmc}")

        if not dois and 'researchgate.net' in item['url'].lower():
            self.log(f"  🔍 Обработка ResearchGate...")
            rg_title = self.extract_researchgate_title(item['url'])
            if rg_title:
                rg_doi = self.search_doi_via_crossref(rg_title, item['url'])
                if rg_doi:
                    dois.add(rg_doi)
                    self.log(f"  ✅ ResearchGate DOI: {rg_doi}")

        if not dois and item['title']:
            self.log(f"  🔍 Резервный поиск по заголовку...")
            doi = self.search_doi_via_crossref(item['title'], item['url'])
            if doi:
                dois.add(doi)
                self.log(f"  ✅ DOI найден по заголовку: {doi}")

        if not dois:
            self.log(f"  ⚠️ DOI не найден")

        return sorted(dois)
