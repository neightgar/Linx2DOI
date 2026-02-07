"""
Делегаты для отображения специальных элементов в таблицах
"""

from PyQt6.QtWidgets import QStyledItemDelegate, QLineEdit
from PyQt6.QtCore import Qt, QUrl
from PyQt6.QtGui import QColor, QDesktopServices


class HyperlinkDelegate(QStyledItemDelegate):
    """Делегат для отображения гиперссылок в таблице"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.link_color = QColor("#1e40af")

    def paint(self, painter, option, index):
        """Отрисовка ячейки с гиперссылкой"""
        if index.data(Qt.ItemDataRole.UserRole + 1):
            option.palette.setColor(option.palette.ColorRole.Text, self.link_color)
        super().paint(painter, option, index)

    def editorEvent(self, event, model, option, index):
        """Обработка клика по гиперссылке"""
        if event.type() == event.Type.MouseButtonRelease and index.data(Qt.ItemDataRole.UserRole + 1):
            url = index.data(Qt.ItemDataRole.UserRole + 2)
            if url:
                QDesktopServices.openUrl(QUrl(url))
                return True
        return super().editorEvent(event, model, option, index)


class ManualDOIDelegate(QStyledItemDelegate):
    """Делегат для редактирования ручного DOI с выравниванием вверх"""

    def createEditor(self, parent, option, index):
        """Создает редактор с выравниванием текста вверх"""
        editor = QLineEdit(parent)
        editor.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        editor.setStyleSheet("""
            QLineEdit {
                padding-top: 2px;
                padding-left: 4px;
            }
        """)
        return editor

    def setEditorData(self, editor, index):
        """Устанавливает данные в редактор"""
        value = index.model().data(index, Qt.ItemDataRole.EditRole)
        editor.setText(str(value) if value else "")

    def setModelData(self, editor, model, index):
        """Сохраняет данные из редактора"""
        model.setData(index, editor.text(), Qt.ItemDataRole.EditRole)
