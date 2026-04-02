"""
Скрипт сборки исполняемого файла Linx2DOI с использованием PyInstaller
Исправлена ошибка с jaraco.text и pkg_resources
"""
import os
import sys
import shutil
import subprocess
from pathlib import Path

# ============================================================================
# КОНФИГУРАЦИЯ СБОРКИ
# ============================================================================
PROJECT_NAME = "Linx2DOI"
ENTRY_POINT = "main.py"
OUTPUT_DIR = "dist"
ICON_FILE = "icon.ico"
VERSION = "1.0.0"

# ============================================================================
# Модули PyQt6 для ИСКЛЮЧЕНИЯ (не используются в проекте)
# ============================================================================
PYQT6_EXCLUDE_MODULES = [
    'PyQt6.QtMultimedia',
    'PyQt6.QtNetwork',
    'PyQt6.QtSql',
    'PyQt6.QtWebEngine',
    'PyQt6.QtWebEngineWidgets',
    'PyQt6.QtTest',
    'PyQt6.QtOpenGL',
    'PyQt6.Qt3DCore',
    'PyQt6.Qt3DRender',
    'PyQt6.QtCharts',
    'PyQt6.QtDataVisualization',
    'PyQt6.QtAxContainer',
    'PyQt6.QtBluetooth',
    'PyQt6.QtDesigner',
    'PyQt6.QtHelp',
    'PyQt6.QtLocation',
    'PyQt6.QtNfc',
    'PyQt6.QtPositioning',
    'PyQt6.QtPrintSupport',
    'PyQt6.QtPurchasing',
    'PyQt6.QtQuick',
    'PyQt6.QtQuickWidgets',
    'PyQt6.QtRemoteObjects',
    'PyQt6.QtScxml',
    'PyQt6.QtSensors',
    'PyQt6.QtSerialPort',
    'PyQt6.QtSpeech',
    'PyQt6.QtSvg',
    'PyQt6.QtWebChannel',
    'PyQt6.QtWebSockets',
    'PyQt6.QtX11Extras',
    'PyQt6.QtXmlPatterns',
]

# ============================================================================
# Другие модули для ИСКЛЮЧЕНИЯ (не используются в проекте)
# ============================================================================
OTHER_EXCLUDE_MODULES = [
    'numpy',
    'scipy',
    'matplotlib',
    'PIL',
    'pillow',
    'tkinter',
    'unittest',
    'pytest',
    'pip',
    'setuptools',
    'wheel',
    'IPython',
    'jupyter',
    'notebook',
    'idlelib',
    'lib2to3',
    'pydoc',
    'ensurepip',
    'test',
]

# ============================================================================
# ФЛАГИ PyInstaller
# ============================================================================
PYINSTALLER_FLAGS = [
    sys.executable,
    "-m", "PyInstaller",

    # === Основные настройки ===
    "--onefile",
    "--windowed",
    "--name", PROJECT_NAME,
    f"--distpath={OUTPUT_DIR}",

    # === Иконка ===
    f"--icon={ICON_FILE}",

    # === Включение пакетов ===
    "--add-data", f"src{os.pathsep}src",
    "--collect-all=PyQt6.QtCore",
    "--collect-all=PyQt6.QtGui",
    "--collect-all=PyQt6.QtWidgets",
    "--collect-all=docx",
    "--collect-all=requests",
    "--collect-all=lxml",
    "--collect-all=pkg_resources",

    # === Скрытые импорты ===
    # PyQt6 зависимости
    "--hidden-import=PyQt6.sip",
    # Стандартная библиотека
    "--hidden-import=xml.etree.ElementTree",
    "--hidden-import=difflib",
    "--hidden-import=urllib.parse",
    "--hidden-import=urllib.request",
    # lxml зависимости
    "--hidden-import=lxml",
    "--hidden-import=lxml.etree",
    "--hidden-import=lxml._elementpath",
    # pkg_resources зависимости (исправлено для setuptools >= 60)
    "--hidden-import=pkg_resources",
    "--hidden-import=jaraco",
    "--hidden-import=jaraco.text",
    "--hidden-import=jaraco.functools",
    "--hidden-import=more_itertools",
    "--hidden-import=typing_extensions",
    "--hidden-import=platformdirs",
    # ✅ Для ISBN функциональности (Google Books API)
    "--hidden-import=json",
    "--hidden-import=re",
    "--hidden-import=time",
    # ✅ Для PubMed поиска (RefChecker функционал)
    "--hidden-import=Bio",
    "--hidden-import=Bio.Entrez",
    "--hidden-import=rapidfuzz",
    "--hidden-import=rapidfuzz.fuzz",

    # === Исключение ненужных модулей PyQt6 ===
]

# Добавляем исключаемые модули PyQt6
for module in PYQT6_EXCLUDE_MODULES:
    PYINSTALLER_FLAGS.extend(["--exclude-module", module])

# Добавляем другие исключаемые модули
for module in OTHER_EXCLUDE_MODULES:
    PYINSTALLER_FLAGS.extend(["--exclude-module", module])

# === Очистка и точка входа ===
PYINSTALLER_FLAGS.extend([
    "--clean",
    "--noconfirm",
    ENTRY_POINT,
])


def check_pyinstaller():
    """Проверяет установку PyInstaller"""
    try:
        result = subprocess.run(
            [sys.executable, "-m", "PyInstaller", "--version"],
            capture_output=True,
            text=True,
            check=True
        )
        print(f"✅ PyInstaller {result.stdout.strip()}")
        return True
    except subprocess.CalledProcessError:
        print("❌ PyInstaller не установлен")
        print("   Установите: pip install pyinstaller pyinstaller-hooks-contrib")
        return False


def check_dependencies():
    """Проверяет наличие необходимых файлов и зависимостей"""
    print("\n🔍 Проверка зависимостей...")

    errors = []

    # Проверка точки входа
    if not os.path.exists(ENTRY_POINT):
        errors.append(f"Точка входа '{ENTRY_POINT}' не найдена")

    # Проверка иконки (не критично)
    if not os.path.exists(ICON_FILE):
        print(f"⚠️  Файл '{ICON_FILE}' не найден (иконка не будет добавлена)")
        global PYINSTALLER_FLAGS
        PYINSTALLER_FLAGS = [f for f in PYINSTALLER_FLAGS if '--icon=' not in f]

    # Проверка пакета src
    if not os.path.exists("src"):
        errors.append("Папка 'src' не найдена")

    # Проверка основных зависимостей
    dependencies = [
        ('lxml', 'lxml'),
        ('jaraco.text', 'jaraco.text'),
        ('platformdirs', 'platformdirs'),
        ('requests', 'requests'),
        ('docx', 'python-docx'),
        ('PyQt6', 'PyQt6'),
        ('Bio', 'biopython'),
        ('rapidfuzz', 'rapidfuzz'),
    ]

    for import_name, package_name in dependencies:
        try:
            __import__(import_name)
            print(f"✅ {package_name} установлен")
        except ImportError:
            print(f"⚠️  {package_name} не установлен. Установите: pip install {package_name}")

    # Проверка requirements.txt
    if os.path.exists("requirements.txt"):
        print("✅ requirements.txt найден")

    if errors:
        print("\n❌ Критические ошибки:")
        for error in errors:
            print(f"   • {error}")
        return False

    print("\n✅ Все зависимости проверены")
    return True


def clean_build_artifacts():
    """Очищает артефакты предыдущей сборки"""
    print("\n🧹 Очистка артефактов сборки...")

    cleaned = 0

    # Очистка output directory
    if os.path.exists(OUTPUT_DIR):
        shutil.rmtree(OUTPUT_DIR)
        print(f"   ✅ Удалено: {OUTPUT_DIR}/")
        cleaned += 1

    # Очистка кэша PyInstaller
    build_cache = "build"
    if os.path.exists(build_cache):
        shutil.rmtree(build_cache)
        print(f"   ✅ Удалено: {build_cache}/")
        cleaned += 1

    # Очистка spec файла
    spec_file = f"{PROJECT_NAME}.spec"
    if os.path.exists(spec_file):
        os.remove(spec_file)
        print(f"   ✅ Удалено: {spec_file}")
        cleaned += 1

    # Очистка временных файлов
    for pattern in ["*.onefile-cache", "*.cache", "__pycache__"]:
        for f in Path(".").rglob(pattern):
            if f.is_dir():
                try:
                    shutil.rmtree(f)
                    print(f"   ✅ Удалено: {f}/")
                    cleaned += 1
                except PermissionError:
                    pass

    if cleaned == 0:
        print("   ℹ️  Нечего очищать")

    return True


def build():
    """Запускает процесс сборки"""
    print("\n🚀 Запуск сборки...")
    print("=" * 60)

    try:
        result = subprocess.run(PYINSTALLER_FLAGS, check=True)

        print("=" * 60)
        print("\n✅ Сборка завершена успешно!")

        exe_path = Path(OUTPUT_DIR) / f"{PROJECT_NAME}.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"📦 Исполняемый файл: {exe_path}")
            print(f"📊 Размер: {size_mb:.1f} MB")

            # Предупреждение если размер слишком большой
            if size_mb > 150:
                print(f"⚠️  Размер больше ожидаемого (>150 MB)")
                print("   Проверьте какие модули были включены")
            else:
                print(f"✅ Размер в норме (<150 MB)")
        else:
            print(f"⚠️  Файл не найден: {exe_path}")

        return True

    except subprocess.CalledProcessError as e:
        print("=" * 60)
        print(f"\n❌ Ошибка сборки (код {e.returncode})")
        print("\nВозможные причины:")
        print("   • Отсутствуют зависимости (pip install -r requirements.txt)")
        print("   • Неполная установка PyInstaller")
        print("   • Конфликт версий Python/PyInstaller")
        return False

    except KeyboardInterrupt:
        print("\n\n⚠️  Сборка прервана пользователем")
        return False


def show_usage():
    """Показывает справку по использованию"""
    print("""
╔═══════════════════════════════════════════════════════════╗
║  Linx2DOI Build Script (PyInstaller)                      ║
╠═══════════════════════════════════════════════════════════╣
║  Использование:                                           ║
║    python build.py              # Полная сборка           ║
║    python build.py --clean      # Только очистка          ║
║    python build.py --check      # Только проверка         ║
║    python build.py --help       # Эта справка             ║
╚═══════════════════════════════════════════════════════════╝
""")


def main():
    """Основная функция"""
    args = sys.argv[1:]

    if "--help" in args or "-h" in args:
        show_usage()
        return

    if "--check" in args:
        check_pyinstaller()
        check_dependencies()
        return

    if "--clean" in args:
        clean_build_artifacts()
        print("\n✅ Очистка завершена")
        return

    print("=" * 60)
    print(f"  {PROJECT_NAME} v{VERSION} - Build Script")
    print("=" * 60)

    if not check_pyinstaller():
        sys.exit(1)

    if not check_dependencies():
        sys.exit(1)

    clean_build_artifacts()

    if not build():
        sys.exit(1)

    print("\n" + "=" * 60)
    print("  ✅ Готово! Можете запускать приложение.")
    print("=" * 60)
    print(f"\n📁 Расположение: {os.path.abspath(OUTPUT_DIR)}\\{PROJECT_NAME}.exe")
    print("\n📋 Следующие шаги:")
    print("   1. Протестируйте приложение на чистой системе")
    print("   2. Проверьте обработку .docx файлов")
    print("   3. Убедитесь, что все функции работают (DOI, ISBN, PMID)")
    print("   4. Проверьте экспорт в RIS")


if __name__ == "__main__":
    main()