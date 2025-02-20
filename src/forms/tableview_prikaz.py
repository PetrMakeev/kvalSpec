from datetime import datetime
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtCore import Qt
from src.modules.resource_image import resource_path

class PrikazTableModel(QtCore.QAbstractTableModel):
    def __init__(self, data, columns):
        super().__init__()
        self._data = data
        self._columns = columns

    def headerData(self, section, orientation, role):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                # return f"Column {section + 1}"
                return str(self._columns[section])


    def data(self, index, role):
    
        if not index.isValid():
            return None

        value = self._data[index.row()][index.column()]
        
        # Отображение текста, исключая колонку 3
        if role == Qt.ItemDataRole.DisplayRole and index.column() != 3:
            if index.column() == 2 and isinstance(value, datetime):
                return datetime.strftime(value, "%d.%m.%Y")
            return value
        
        if role == Qt.ItemDataRole.TextAlignmentRole:
            return Qt.AlignmentFlag.AlignCenter
        
        
        if role == Qt.ItemDataRole.DecorationRole and index.column() == 3:
            # Проверяем, является ли значение именем PDF файла
            value = self._data[index.row()][index.column()]
            if isinstance(value, str) and value.endswith('.pdf'):
                return QtGui.QIcon(resource_path("icons\\pdf.png"))
        
        return None




    def rowCount(self, index):
        # The length of the outer list.
        return len(self._data)

    def columnCount(self, index):
        # The following takes the first sub-list, and returns
        # the length (only works if all rows are an equal length)
        if not self._data:
            return 0
        return len(self._data[0])