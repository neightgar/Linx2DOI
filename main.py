#!/usr/bin/env python3
"""
DOI Extractor & APA Formatter
Точка входа в приложение
"""

import sys
from PyQt6.QtWidgets import QApplication, QMessageBox
from PyQt6.QtGui import QIcon

from src.gui.main_window import MainWindow, get_icon_path


def check_dependencies():
    """Проверка наличия необходимых зависимостей"""
    try:
        import requests
        from docx import Document
        return True
    except ImportError as e:
        QMessageBox.critical(
            None,
            "Ошибка зависимостей",
            f"Установите зависимости:\n"
            f"pip install python-docx requests PyQt6\n\n"
            f"Ошибка: {str(e)}"
        )
        return False


def setup_dark_theme(app):
    """Настройка темной темы Fusion"""
    from PyQt6.QtGui import QPalette, QColor
    from PyQt6.QtCore import Qt

    # Темная палитра
    dark_palette = QPalette()

    # Основные цвета
    dark_palette.setColor(QPalette.ColorRole.Window, QColor(45, 45, 48))
    dark_palette.setColor(QPalette.ColorRole.WindowText, QColor(230, 230, 230))
    dark_palette.setColor(QPalette.ColorRole.Base, QColor(30, 30, 32))
    dark_palette.setColor(QPalette.ColorRole.AlternateBase, QColor(40, 40, 42))
    dark_palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(60, 60, 65))
    dark_palette.setColor(QPalette.ColorRole.ToolTipText, QColor(230, 230, 230))
    dark_palette.setColor(QPalette.ColorRole.Text, QColor(230, 230, 230))
    dark_palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 57))
    dark_palette.setColor(QPalette.ColorRole.ButtonText, QColor(230, 230, 230))
    dark_palette.setColor(QPalette.ColorRole.BrightText, QColor(255, 0, 0))
    dark_palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
    dark_palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
    dark_palette.setColor(QPalette.ColorRole.HighlightedText, QColor(0, 0, 0))

    # Отключенные элементы
    dark_palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.WindowText, QColor(127, 127, 127))
    dark_palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, QColor(127, 127, 127))
    dark_palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, QColor(127, 127, 127))
    dark_palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Highlight, QColor(80, 80, 80))
    dark_palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.HighlightedText, QColor(127, 127, 127))

    app.setPalette(dark_palette)


def main():
    """Главная функция приложения"""
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Установка иконки приложения
    icon_path = get_icon_path()
    if icon_path:
        app.setWindowIcon(QIcon(str(icon_path)))

    # Применяем темную тему
    setup_dark_theme(app)

    # Проверка зависимостей
    if not check_dependencies():
        sys.exit(1)

    # Создание и отображение главного окна
    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
