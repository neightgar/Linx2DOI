"""
Конфигурация приложения
"""

# API endpoints
ENTREZ_EFETCH = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
CROSSREF_WORKS = "https://api.crossref.org/works"
GOOGLE_BOOKS_API = "https://www.googleapis.com/books/v1/volumes"

# Настройки приложения
APP_NAME = "DOIExtractor"
APP_VERSION = "1.0"
SETTINGS_ORG = "DOIExtractor"
SETTINGS_APP = "APAFormatter"

# Параметры обработки
MAX_RETRIES = 3
REQUEST_TIMEOUT = 10
RATE_LIMIT_DELAY = 0.5  # секунды между запросами DOI метаданных

# Максимальный размер диапазона цитирований (защита от [1-999999])
MAX_CITATION_RANGE = 100

# Цвета для статусов
STATUS_COLORS = {
    'NOT_IN_TEXT': '#9e9e9e',  # Серый
    'NO_DATA': '#dc3545',      # Красный
    'DUPLICATE': '#f59e0b',    # Оранжевый
    'MANUAL': '#10b981',       # Зелёный
    'OK': '#63b3ed',           # Синий
    'ISBN': '#9b59b6'          # Фиолетовый для ISBN
}