from datetime import datetime
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtCore import Qt

class SpecTableModel(QtCore.QAbstractTableModel):
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
        if role == Qt.ItemDataRole.DisplayRole:
            # Получаем исходное значение
            value = self._data[index.row()][index.column()]

            #  поля  ["id", ФИО", "Подразделение", "Должность", "Квалификация", "Дата следующей сдачи"]
             
            if index.column() == 5:
                return datetime.strftime(value, "%d.%m.%Y")
            else:
                return self._data[index.row()][index.column()]
            
        if role == QtCore.Qt.ItemDataRole.TextAlignmentRole and (index.column() != 1 ):
            return QtCore.Qt.AlignmentFlag.AlignHCenter|QtCore.Qt.AlignmentFlag.AlignVCenter
        if role == QtCore.Qt.ItemDataRole.TextAlignmentRole and (index.column() == 1 ):
            return QtCore.Qt.AlignmentFlag.AlignLeft|QtCore.Qt.AlignmentFlag.AlignVCenter

        if role == Qt.ItemDataRole.BackgroundRole and index.column() == 5:
            current_date = datetime.now()
            end_of_year = datetime(current_date.year, 12, 31)
            cell_date = self._data[index.row()][index.column()]

            if isinstance(cell_date, datetime):
                if current_date <= cell_date <= end_of_year:
                    return QtGui.QColor(144, 238, 144)
                elif cell_date < current_date:
                    return QtGui.QColor("yellow")


    def rowCount(self, index):
        # The length of the outer list.
        return len(self._data)

    def columnCount(self, index):
        # The following takes the first sub-list, and returns
        # the length (only works if all rows are an equal length)
        if not self._data:
            return 0
        return len(self._data[0])