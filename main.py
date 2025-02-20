from PyQt6 import QtWidgets
from PyQt6.QtGui import QFont
from src.forms.MainWindow import MainWindow

import sys

def main():

    app = QtWidgets.QApplication([])
    # стиль приложения
    app.setStyle('Fusion')
    app.setFont(QFont('Arial', 10))
    
    application = MainWindow()
    
    application.show()
    
    sys.exit(app.exec())

 
if __name__ == "__main__":
    main()