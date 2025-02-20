from PyQt6.QtWidgets import (QMainWindow,
                             QTableView, 
                             QMenu, QMessageBox, QFileDialog)
from PyQt6 import (QtCore, 
                   QtGui, 
                   QtWidgets)
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtGui import QCursor, QAction, QIcon
from src.ui.ui_prikaz import Ui_frmPrikaz

from src.modules.db_local import DB_Local
from src.forms.tableview_prikaz import PrikazTableModel
from datetime import datetime
from src.forms.widget_manager import set_widget_background, check_widget
from os import path, getcwd, remove
from src.modules.pdfContainer import PDFContainer
from src.modules.resource_image import resource_path


class PrikazWindow(QMainWindow, Ui_frmPrikaz):
    def __init__(self, parent=None, *args, obj=None, **kwargs):
        super(PrikazWindow, self).__init__(parent, *args, **kwargs) 
        self.setupUi(self)
        
        self.parent = parent
        
        self.prikaz_id = None
        self.prikaz_nom = None
        self.prikaz_dt = None
        self.prikaz_filename = None
        self.prikaz_path = None
        self.mode_prikaz = None
        self.mode_edit = "NONE"
        
        self.container = PDFContainer()

        ###################################################
        #  панель грида
        ###################################################
        # настраиваем и заполняем грид
        self.db = DB_Local()
        self.refresh_grid()
        
                  
        # привязываем контекстное меню к гриду
        self.table.customContextMenuRequested.connect(self.contextMenuEvent)

        
        
        ###################################################
        #  панель редактирования
        ###################################################
        # сигнал изменения полей вызывает перекраску полей по условию
        self.txt_prikaz_nom.textChanged.connect(lambda: check_widget(self.txt_prikaz_nom, len(self.txt_prikaz_nom.text().strip()) == 0))
        self.txt_prikaz_dt.dateChanged.connect(lambda: check_widget(self.txt_prikaz_dt, self.txt_prikaz_dt.date() < QDate(2000, 1, 1)))
        
        # кнопки выбора и очистки выбора файла скана приказа
        self.btn_prikaz_link_sel.clicked.connect(self.open_pdf)
        self.btn_prikaz_link_clr.clicked.connect(self.clear_pdf)
        
        # кнопка сохранить
        self.btn_save.clicked.connect(self.save_prikaz)        
        
        # кнопка отмена
        self.btn_cancel.clicked.connect(self.saveno_prikaz)
        
        # кнопка сохранения в эксель
        

        
    ##############################################################################
    # popup меню грида  с функционалом
    #############################################################################
    # настройка вызова контекстного меню
    def contextMenuEvent(self, point):
        # проверяем чтобы был клик в области строк грида
        if  (type(point) == QtCore.QPoint):
            context_menu = QMenu(self)
            
            
            pop_sel = QAction(QIcon(resource_path("icons\\sel.png")), "Выбрать приказ", self)
            context_menu.addAction(pop_sel)
            pop_sel.triggered.connect(lambda: self.select_prikaz(self.table.currentIndex().row()))

            context_menu.addSeparator()
            
            
            pop_add = QAction(QIcon(resource_path("icons\\add.png")), "Новый приказ", self)
            context_menu.addAction(pop_add)
            pop_add.triggered.connect(self.popup_add)
            
            context_menu.addSeparator()

            # настраиваем в зависимости от данных в поле скана 
            if not self.data_prikaz[self.table.currentIndex().row()][3]:
                
                pop_edit_pdf = QAction(QIcon(resource_path("icons\\import.png")), "Добавить скан приказа", self)
                context_menu.addAction(pop_edit_pdf)
                pop_edit_pdf.triggered.connect(self.popup_add_pdf)
            
            else:
                
                pop_del_pdf = QAction(QIcon(resource_path("icons\\clear.png")), "Удалить скан приказа", self)
                context_menu.addAction(pop_del_pdf)
                pop_del_pdf.triggered.connect(self.popup_del_pdf)
                
                pop_extract_pdf = QAction(QIcon(resource_path("icons\\extract.png")), "Извлечь скан приказа", self)
                context_menu.addAction(pop_extract_pdf)
                pop_extract_pdf.triggered.connect(self.popup_extract_pdf)
            
            
            context_menu.addSeparator()

            pop_del = QAction(QIcon(resource_path("icons\\del.png")), "Удалить приказ", self)
            context_menu.addAction(pop_del)
            pop_del.triggered.connect(self.popup_del)
            
            context_menu.popup(QCursor.pos())
            
        
   
        
    # пункт контекстного меню - удалить
    def popup_del(self, event):
        confirmation_dialog = QMessageBox(self)
        confirmation_dialog.setIcon(QMessageBox.Icon.Warning)
        confirmation_dialog.setWindowTitle("Подтверждение")
        confirmation_dialog.setText("Вы уверены, что хотите удалить запись?")
        confirmation_dialog.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        confirmation_dialog.setDefaultButton(QMessageBox.StandardButton.No)

        # Получение ответа пользователя
        user_response = confirmation_dialog.exec()

        if user_response == QMessageBox.StandardButton.Yes:
        # id выбранной строки
            curr_row = self.table.currentIndex().row()
            self.curr_prikaz_id = self.data_prikaz[curr_row][0] 
            # id выбранной строки после удаления
            curr_row_post_del = (curr_row - 1) if curr_row > 0 else 0 
            curr_id_post_del = self.data_prikaz[curr_row_post_del][0]  
            self.del_prikaz(self.curr_prikaz_id, curr_id_post_del, self.data_prikaz[curr_row][3])      
         
        

        
    # пункт контекстного меню - добавить
    def popup_add(self, event):
        # очищаем поля
        self.txt_prikaz_nom.setText("")
        self.txt_prikaz_dt.setDate(datetime.strptime("1900/01/01", "%Y/%m/%d"))
        self.btn_prikaz_link.setText("")
        self.prikaz_filename = None
        self.prikaz_path = None
        self.curr_prikaz_id = -1

        # подкрашиваем обязательные поля
        check_widget(self.txt_prikaz_nom, len(self.txt_prikaz_nom.text().strip()) == 0)
        check_widget(self.txt_prikaz_dt, self.txt_prikaz_dt.date() < QDate(2000, 1, 1))
          
        # переводим в режим добавления
        self.mode_edit = 'ADD'
        self.txt_prikaz_dt.setReadOnly(False)
        self.txt_prikaz_nom.setReadOnly(False)
        self.frm_grid.setVisible(False)
       
        
         
    # пункт контекстного меню - изменить скан приказа
    def popup_add_pdf(self, event):   
        curr_row = self.table.currentIndex().row()   
        
        self.curr_prikaz_id = self.data_prikaz[curr_row][0]
        data = self.db.read_prikaz_to_edit(self.curr_prikaz_id)
        prikaz_dt = datetime.strptime(data[1],"%Y-%m-%d %H:%M:%S") 
        prikaz_nom = data[0]
        file_path, _ = QFileDialog.getOpenFileName(self, "Select PDF File", "", "PDF Files (*.pdf)")
        if file_path:
            try:
                # добавляем в контейнер
                filename = self.container.add_file(file_path, datetime.strftime(prikaz_dt, "%d.%m.%Y"), prikaz_nom, )
                # сохраняем в базу
                prikaz = {
                    "id"          : self.curr_prikaz_id,
                    "prikaz_nom"  : prikaz_nom,
                    "prikaz_dt"   : prikaz_dt,
                    "pdf_link"    : filename
                }  
                self.db.prikaz_save(prikaz)
                self.refresh_grid()

            except ValueError as e:
                QMessageBox.critical(self, "Error", str(e))
                
        
        
    # пункт контекстного меню - удалить скан приказа
    def popup_del_pdf(self, event):   
        curr_row = self.table.currentIndex().row()   
        
        self.curr_prikaz_id = self.data_prikaz[curr_row][0]
        data = self.db.read_prikaz_to_edit(self.curr_prikaz_id)
        prikaz_dt = datetime.strptime(data[1],"%Y-%m-%d %H:%M:%S") 
        prikaz_nom = data[0]
        prikaz_filename = data[2]
        prikaz = {
            "id"          : self.curr_prikaz_id,
            "prikaz_nom"  : prikaz_nom,
            "prikaz_dt"   : prikaz_dt,
            "pdf_link"    : ""
        }  
        self.db.prikaz_save(prikaz)
        #  помечаем файл скана приказа на удаление и сжимаем контейнер
        self.container.mark_for_deletion(prikaz_filename)
        self.container.pack_container()
        self.refresh_grid()
        # Проходим по всем строкам модели для поиска новой позиции отредактированной строки
        for row in range(self.table.model().rowCount(QtCore.QModelIndex())):
            index = self.model.index(row, 0)  # Предполагаем, что id находится в первом столбце
            if self.model.data(index, Qt.ItemDataRole.DisplayRole) == self.curr_prikaz_id:
                self.table.setCurrentIndex(index)
                self.table.scrollTo(index)  # Прокрутить до выделенной строки

        
        
        
    # пункт контекстного меню - извлечь скан приказа
    def popup_extract_pdf(self, event):   
        curr_row = self.table.currentIndex().row()   
        
        self.curr_prikaz_id = self.data_prikaz[curr_row][0]
        data = self.db.read_prikaz_to_edit(self.curr_prikaz_id)  
        prikaz_filename = data[2]
        save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить как", prikaz_filename, "PDF файлы (*.pdf)")
        if save_path:
            try:
                self.container.extract_file(prikaz_filename, save_path)
                QMessageBox.information(self, "Успешно", f"Файл скана приказа сохранен.")
            except FileNotFoundError as e:
                QMessageBox.critical(self, "Error", str(e))


        

    ##############################################################################
    # добавление, сохранение, удаление приказов
    #############################################################################
    # сохранение сведение о приказе  
    def save_prikaz(self):
        
        prikaz_nom = False
        prikaz_dt = False
        # проверка на дубликат
        for row in self.data_prikaz:
            if str(row[1]).upper() == self.txt_prikaz_nom.text().upper():
                prikaz_nom = True
                
            if datetime(row[2].year, row[2].month, row[2].day, 0 ,0 ,0).date() == self.txt_prikaz_dt.date().toPyDate():
                prikaz_dt = True
            if prikaz_dt and prikaz_nom and self.mode_edit == "ADD":
                info_dialog = QMessageBox(self)
                info_dialog.setIcon(QMessageBox.Icon.Information)
                info_dialog.setWindowTitle("Сохранение отменено!")
                info_dialog.setText(f"Приказ от {self.txt_prikaz_dt.text()} № {self.txt_prikaz_nom.text().strip()} есть в базе данных.")
                info_dialog.setStandardButtons(QMessageBox.StandardButton.Ok)
                info_dialog.exec()
                return
        
        # проверяем заполнение полей
        if len(self.txt_prikaz_nom.text().strip()) == 0:
            self.txt_prikaz_nom.setFocus()
            return
 
        if self.txt_prikaz_dt.date() < QDate(2000, 1, 1):
            self.txt_prikaz_dt.setFocus()     
            return 

        # если указан файл PDF добавляем его в контейнер
        filename = ""
        if self.prikaz_path:
            filename = self.container.add_file(self.prikaz_path, self.txt_prikaz_dt.text(), self.txt_prikaz_nom.text()) 


        # сохраняем данные                        
        if self.mode_edit == 'ADD':
            prikaz = {
                "id"          : -1,
                "prikaz_nom"  : self.txt_prikaz_nom.text(),
                "prikaz_dt"   : self.txt_prikaz_dt.date().toPyDate(),
                "pdf_link"    : filename
            }  
            
        elif self.mode_edit == 'EDIT':
            prikaz = {
                "id"          : self.curr_prikaz_id,
                "prikaz_nom"  : self.txt_prikaz_nom.text(),
                "prikaz_dt"   : self.txt_prikaz_dt.date().toPyDate(),
                "pdf_link"    : filename
            }    
        else:
            return

        # обновляем грид
        self.db.prikaz_save(prikaz=prikaz)
        self.refresh_grid()
        
        
        # Восстанавливаем текущую позицию
        if self.mode_edit == "ADD":
            # определяем новый id приказа после добавления
            self.curr_prikaz_id = max(row[0] for row in self.data_prikaz)
        
        # Проходим по всем строкам модели для поиска новой позиции отредактированной строки
        for row in range(self.table.model().rowCount(QtCore.QModelIndex())):
            index = self.model.index(row, 0)  # Предполагаем, что id находится в первом столбце
            if self.model.data(index, Qt.ItemDataRole.DisplayRole) == self.curr_prikaz_id:
                self.table.setCurrentIndex(index)
                self.table.scrollTo(index)  # Прокрутить до выделенной строки
            
        # выключаем режим редактирования
        self.mode_edit = 'NONE'
        self.frm_grid.setVisible(True)
        self.table.setFocus()




    # удаление приказа  
    def del_prikaz(self, curr_id, curr_id_post, prikaz_filename):
        # проверяем приказы связанные с сотрудниками
        if len(self.db.check_spec_on_prikaz(curr_id))>0:
            info_dialog = QMessageBox(self)
            info_dialog.setIcon(QMessageBox.Icon.Information)
            info_dialog.setWindowTitle("Удаление отменено!")
            info_dialog.setText("Данный приказ используется.")
            info_dialog.setStandardButtons(QMessageBox.StandardButton.Ok)
            info_dialog.exec()
            return
        
        # обновляем грид
        self.db.prikaz_del(curr_id)
        self.refresh_grid()
        
        # если есть имя скана приказа то помечаем его на удаление и сжимаем контейнер
        self.container.mark_for_deletion(prikaz_filename)
        self.container.pack_container()
        
        # Проходим по всем строкам модели для поиска новой позиции отредактированной строки
        for row in range(self.table.model().rowCount(QtCore.QModelIndex())):
            index = self.model.index(row, 0)  # Предполагаем, что id находится в первом столбце
            if self.model.data(index, Qt.ItemDataRole.DisplayRole) == curr_id_post:
                #self.table.setCurrentIndex(index)
                self.table.scrollTo(index)  # Прокрутить до выделенной строки
            
        # выключаем режим редактирования
        self.mode_edit = 'NONE'
        self.frm_grid.setVisible(True)
        self.table.setFocus()
            
        
        
    # отмена  сохранения
    def saveno_prikaz(self):
        self.mode_edit = "NONE"
        self.frm_grid.setVisible(True)        
                
        
        
        
    ##############################################################################
    # выбор приказа и передача сведений о нем в основную форму  
    #############################################################################      
    def select_prikaz(self, row):
        if not (self.mode_prikaz == "FILTER"):
            # передаем сведения о приказе в основное окно
            if self.parent:
                self.parent.update_edit_from_prikaz(self.data_prikaz[row][0], 
                                                    self.data_prikaz[row][1],
                                                    self.data_prikaz[row][2])
        else:
            # передаем сведения о приказе в фильтры
            if self.parent:
                self.parent.update_filter_from_prikaz(self.data_prikaz[row][0], 
                                                      self.data_prikaz[row][1],
                                                      self.data_prikaz[row][2])
        self.hide()
        
        
    ##############################################################################
    # выбор и очистка файла скана приказа   
    #############################################################################   
    def open_pdf(self):        
        # Диалог открытия файла
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("PDF Files (*.pdf)")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        
        if file_dialog.exec():
            self.prikaz_path= file_dialog.selectedFiles()[0]
            self.prikaz_filename = path.basename(self.prikaz_path)
            self.btn_prikaz_link.setText(self.prikaz_filename)

    def clear_pdf(self):
        self.btn_prikaz_link.setText("")
        self.prikaz_filename = None
        self.prikaz_file = None



    ########################################################################################
    # событие перед закрытием приложения
    ######################################################################################   

    def closeEvent(self, event):
        

        if self.mode_edit == "NONE":
            # self.saveno_prikaz()
            event.accept()
        else:
            if self.mode_edit == "FILTER":
                self.saveno_prikaz()
            else:
                # Показываем диалог для подтверждения выхода
                reply = QMessageBox.question(
                    self,
                    "Новый приказ",
                    "Сохранить сведения о новом приказе?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                    
                if reply == QMessageBox.StandardButton.Yes:
                    # Выполняем действия перед закрытием, если нужно
                    if self.mode_edit == "ADD":
                        self.save_prikaz()
                        event.accept()
                    else:
                        # Отменяем закрытие
                        event.ignore()
                else:
                    self.saveno_prikaz()
                
                
    ########################################################################################
    # обновляем грид
    ######################################################################################   
    def refresh_grid(self):
        # обновляем данные        
        self.data_prikaz = self.db.read_prikaz()
        self.model = PrikazTableModel(self.data_prikaz, ["id", "Номер", "Дата", " "])
        self.table.setModel(self.model)
        # ширина столбцов
        self.table.setColumnWidth(0, 0)
        self.table.setColumnWidth(1, 70)
        self.table.setColumnWidth(2, 90)
        self.table.setColumnWidth(3, 20)

        
        # режим выделения = строка целиком
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)  
        self.model.layoutChanged.emit()                