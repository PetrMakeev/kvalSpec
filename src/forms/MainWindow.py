from PyQt6.QtWidgets import (QMainWindow,
                            QTableView, 
                            QMenu,
                            QComboBox, QMessageBox, QFileDialog
                             )
from PyQt6 import (QtCore, 
                   QtGui, 
                   QtWidgets)
from PyQt6.QtCore import Qt, QDate, QSize
from PyQt6.QtGui import QCursor, QFont, QAction, QIcon
from src.ui.ui_main import Ui_MainWindow
from src.modules.filters import el_spec, el_kval, el_zvan, el_podr
from src.modules.db_local import DB_Local
from src.forms.tableview_spec import SpecTableModel
from datetime import datetime
from src.forms.PrikazWindow import PrikazWindow
from dateutil.relativedelta import relativedelta
from src.modules.filters_manager import FilterManager
from src.forms.widget_manager import set_widget_background, check_widget
from src.modules.export_xls import export_xls
from src.modules.resource_image import resource_path

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, *args, obj=None, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs) 
        self.setupUi(self)
        
        self.setWindowIcon(QIcon(resource_path("icons\\database.ico")))
        
        # создаем форму приказов
        self.w_prikaz = PrikazWindow(parent = self)
        
         # Устанавливаем изображение
        icon = QIcon(resource_path("icons\\print.png"))  # Укажите путь к изображению
        self.btn_print.setIcon(icon)
        self.btn_print.setIconSize(QSize(40 , 40))
        
        ###################################################
        #  панель редактирования
        ###################################################
        # режим редактирования
        self.mode_edit = 'NONE'
        self.curr_id_spec = -1 
        self.curr_prikaz_id = -1 # для панели редактирования
        self.filter_prikaz_id = None # для панели грида
        
        
        
        # заполняем комбобоксы
        for el in el_spec:
            if not el == "все":
                self.cmb_edit_status.addItem(el)
        self.cmb_edit_status.setCurrentIndex(-1)
                
        self.cmb_edit_kval.addItems(el_kval)
        self.cmb_edit_kval.setCurrentIndex(-1)
        
        self.cmb_edit_podr.addItems(el_podr)
        self.cmb_edit_podr.setCurrentIndex(-1)
        
        self.cmb_edit_zvan.addItems(el_zvan)
        self.cmb_edit_zvan.setCurrentIndex(-1)
        
        
        # сигнал изменения полей вызывает перекраску полей по условию
        self.txt_edit_family.textChanged.connect(lambda: check_widget(self.txt_edit_family, len(self.txt_edit_family.text().strip()) == 0))
        self.txt_edit_name.textChanged.connect(lambda: check_widget(self.txt_edit_name, len(self.txt_edit_name.text().strip()) == 0))
        self.txt_edit_lastname.textChanged.connect(lambda: check_widget(self.txt_edit_lastname, len(self.txt_edit_lastname.text().strip()) == 0))
        self.cmb_edit_podr.currentIndexChanged.connect(lambda: check_widget(self.cmb_edit_podr, self.cmb_edit_podr.currentIndex() < 0))
        self.txt_edit_dolzh.textChanged.connect(lambda: check_widget(self.txt_edit_dolzh, len(self.txt_edit_dolzh.text().strip()) == 0))
        self.cmb_edit_zvan.currentIndexChanged.connect(lambda: check_widget(self.cmb_edit_zvan, self.cmb_edit_zvan.currentIndex() < 0))
        self.cmb_edit_kval.currentIndexChanged.connect(lambda: check_widget(self.cmb_edit_kval, self.cmb_edit_kval.currentIndex() < 0))
        self.txt_edit_prikaz.textChanged.connect(lambda: check_widget(self.txt_edit_prikaz, self.curr_prikaz_id == -1))
        self.cmb_edit_status.currentIndexChanged.connect(lambda: check_widget(self.cmb_edit_status, self.cmb_edit_status.currentIndex() < 0))
        
        
        # вызов формы справочника Приказов из панели редактирования
        self.btn_edit_prikaz.clicked.connect(lambda: self.open_prikaz('EDIT'))


        # кнопка сохранить
        self.btn_save.clicked.connect(self.save_spec)
        
        # кнопка отмена
        self.btn_cancel.clicked.connect(self.saveno_spec)




        ###################################################
        #  панель грида
        ###################################################
        # заполняем комбобоксы фильтров
        self.cmb_spec.addItems(el_spec)
        self.cmb_spec.setCurrentIndex(3)
        
        self.cmb_kval.addItems(el_kval)
        self.cmb_kval.setCurrentIndex(-1)
        
        self.cmb_zvan.addItems(el_zvan)
        self.cmb_zvan.setCurrentIndex(-1)
        
        self.cmb_podr.addItems(el_podr)
        self.cmb_podr.setCurrentIndex(-1)

        # выставляем флаги фильтров
        # Настройка фильтров
        self.filters_config = {
            "podr": False,
            "kval": False,
            "zvan": False,
            "prikaz": False,
            "srok_tg": False,
            "srok_end": False,
        }
        self.ui_elements = {
            "btn_podr": self.btn_podr,
            "cmb_podr": self.cmb_podr,
            "btn_kval": self.btn_kval,
            "cmb_kval": self.cmb_kval,
            "btn_zvan": self.btn_zvan,
            "cmb_zvan": self.cmb_zvan,
            "btn_prikaz": self.btn_prikaz,
            "txt_prikaz": self.txt_prikaz,
            "btn_sel_prikaz": self.btn_sel_prikaz,
            "btn_srok_tg": self.btn_srok_tg,
            "btn_srok_end": self.btn_srok_end,
        }
        self.filter_manager = FilterManager(self.filters_config, self.ui_elements)        
        
      
        # приказ выбранной строки
        self.curr_prikaz_id = -1

        # соединяем сигналы клика кнопок флагов фильтров
        # Подключение кнопок к фильтрам
        self.btn_podr.clicked.connect(lambda: self.filter_toggled_set("podr", self.btn_podr.isChecked()))
        self.btn_kval.clicked.connect(lambda: self.filter_toggled_set("kval", self.btn_kval.isChecked()))
        self.btn_zvan.clicked.connect(lambda: self.filter_toggled_set("zvan", self.btn_zvan.isChecked()))
        self.btn_prikaz.clicked.connect(lambda: self.filter_toggled_set("prikaz", self.btn_prikaz.isChecked()))
        self.btn_srok_tg.clicked.connect(lambda: self.filter_toggled_set("srok_tg", self.btn_srok_tg.isChecked()))
        self.btn_srok_end.clicked.connect(lambda: self.filter_toggled_set("srok_end", self.btn_srok_end.isChecked()))
        # изменение значений фильтров 
        self.cmb_spec.currentIndexChanged.connect(self.filter_set)
        self.cmb_podr.currentIndexChanged.connect(self.filter_set)
        self.cmb_kval.currentIndexChanged.connect(self.filter_set)
        self.cmb_zvan.currentIndexChanged.connect(self.filter_set)
        
        
        # настраиваем и заполняем грид
        self.db = DB_Local()
        self.filter_set()
        # self.data_spec = self.db.read_spec()
        
        
        # привязываем контекстное меню к гриду
        self.table.customContextMenuRequested.connect(self.contextMenuEvent)
        
        # вызов формы приказа из панели фильтров
        self.btn_sel_prikaz.clicked.connect(lambda: self.open_prikaz('FILTER'))
        
        # кнопка сохранения в эксель
        self.btn_print.clicked.connect(self.export_xls)

        

    ##############################################################################
    # экспорт в эксель 
    #############################################################################
    def export_xls(self):
        save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить как", "Сведения о", "Excel файл (*.xlsx)")
        if save_path:
            try:
                
                filter_spec = f" status = {self.cmb_spec.currentIndex()} " if self.cmb_spec.currentIndex() <3 else None
                # подразделение
                filter_podr = f" podr = '{self.cmb_podr.currentText()}' " if self.cmb_podr.currentIndex() > -1 else None
                #квалификация
                filter_kval = f" kval = '{self.cmb_kval.currentText()}' " if self.cmb_kval.currentIndex() > -1 else None
                # звание
                filter_zvan = f" zvan = '{self.cmb_zvan.currentText()}' " if self.cmb_zvan.currentIndex() > -1 else None
                # приказ
                filter_prikaz = f" idPrikaz = {self.filter_prikaz_id} " if self.filter_prikaz_id else None
                # сдача в текущем году
                filter_tg = " dtSledKval BETWEEN strftime('%Y-01-01', 'now') AND strftime('%Y-12-31', 'now') " if self.btn_srok_tg.isChecked() else None
                # сдача просрочена
                filter_end = " DATE(dtSledKval) < DATE('now') " if self.btn_srok_end.isChecked() else None

                self.data_export = self.db.read_spec_export(
                                    filter_spec=filter_spec,
                                    filter_podr=filter_podr,
                                    filter_kval=filter_kval,
                                    filter_zvan=filter_zvan,
                                    filter_prikaz=filter_prikaz,
                                    filter_tg=filter_tg,
                                    filter_end=filter_end)
                
                export_xls(self.data_export, save_path)
                
                QMessageBox.information(self, "Успешно", f"Файл {save_path} сохранен.")
            except FileNotFoundError as e:
                QMessageBox.critical(self, "Error", str(e))
        
                # сотрудники, военные, уволенные


        print("export xls")

        
        

    ##############################################################################
    # popup меню грида  с функционалом
    #############################################################################
    # настройка вызова контекстного меню
    def contextMenuEvent(self, point):
        # проверяем чтобы был клик в области строк грида
        if  (type(point) == QtCore.QPoint):
            context_menu = QMenu(self)
            
           
            pop_add = QAction(QIcon(resource_path("icons\\add.png")), "Новый специалист", self)
            context_menu.addAction(pop_add)
            pop_add.triggered.connect(self.popup_add)
            
            # если нет записей в гриде отключаем пункты меню
            if self.data_spec:
                pop_edit = QAction(QIcon(resource_path("icons\\edit.png")), "Изменить сведения", self)
                context_menu.addAction(pop_edit)
                pop_edit.triggered.connect(self.popup_edit)
                
                pop_del = QAction(QIcon(resource_path("icons\\del.png")), "Удалить запись", self)
                context_menu.addAction(pop_del)
                pop_del.triggered.connect(self.popup_del)


            context_menu.popup(QCursor.pos())
            
            
    # пункт контекстного меню - добавить
    def popup_add(self, event):
        self.txt_edit_family.clear()
        self.txt_edit_name.clear()
        self.txt_edit_lastname.clear()
        self.cmb_edit_podr.setCurrentIndex(-1)
        self.txt_edit_dolzh.clear()
        self.cmb_edit_zvan.setCurrentIndex(-1)
        self.cmb_edit_kval.setCurrentIndex(-1)
        self.chk_edit_podtv.setChecked(False)
        self.curr_prikaz_id = -1
        self.curr_prikaz_dt = datetime.strptime("1900/01/01", "%Y/%m/%d")
        self.txt_edit_prikaz.clear()
        self.cmb_edit_status.setCurrentIndex(-1)
        
        # подкрашиваем поля
        check_widget(self.txt_edit_family, len(self.txt_edit_family.text().strip()) == 0)
        check_widget(self.txt_edit_name, len(self.txt_edit_name.text().strip()) == 0)
        check_widget(self.txt_edit_lastname, len(self.txt_edit_lastname.text().strip()) == 0)
        check_widget(self.cmb_edit_podr, self.cmb_edit_podr.currentIndex() < 0)
        check_widget(self.txt_edit_dolzh, len(self.txt_edit_dolzh.text().strip()) == 0)
        check_widget(self.cmb_edit_zvan, self.cmb_edit_zvan.currentIndex() < 0)
        check_widget(self.cmb_edit_kval, self.cmb_edit_kval.currentIndex() < 0)
        check_widget(self.txt_edit_prikaz, self.curr_prikaz_id == -1)
        check_widget(self.cmb_edit_status, self.cmb_edit_status.currentIndex() < 0)
          
        # переводим в режим добавления
        self.mode_edit = 'ADD'
        self.frm_grid.setVisible(False)
        
        
    # пункт контекстного меню - изменить
    def popup_edit(self, event):    
        curr_row = self.table.currentIndex().row()   
        
        self.curr_id_spec = self.data_spec[curr_row][0]
        data = self.db.read_spec_to_edit(self.curr_id_spec)
        
        self.txt_edit_family.setText(data[1])
        self.txt_edit_name.setText(data[2])
        self.txt_edit_lastname.setText(data[3])
        self.cmb_edit_podr.setCurrentIndex(el_podr.index(data[4]))
        self.txt_edit_dolzh.setText(data[5])
        self.cmb_edit_zvan.setCurrentIndex(el_zvan.index(data[6]))
        self.cmb_edit_kval.setCurrentIndex(el_kval.index(data[7]))
        self.chk_edit_podtv.setChecked(data[8] == 1)
        self.curr_prikaz_id = data[9]
        self.cmb_edit_status.setCurrentIndex(data[11])

        self.txt_edit_prikaz.setText(f'№ {data[12]} от {datetime.strptime(data[13], "%Y-%m-%d %H:%M:%S").strftime("%d.%m.%Y")}')
        self.curr_prikaz_dt = datetime.strptime(data[13], "%Y-%m-%d %H:%M:%S")
          
        #  переводим в режим редактирования
        self.mode_edit = 'EDIT'
        self.frm_grid.setVisible(False)
        
                     
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
            self.curr_prikaz_id = self.data_spec[curr_row][0] 
            # id выбранной строки после удаления
            curr_row_post_del = (curr_row - 1) if curr_row > 0 else 0 
            curr_id_post_del = self.data_spec[curr_row_post_del][0]  
            self.del_spec(self.curr_prikaz_id, curr_id_post_del) 
            



        
    ##############################################################################
    # получаем состояние флагов фильтров и настраиваем фильтр
    #############################################################################
    # управление переключателями фильтра
    def filter_toggled_set(self, filter_name, checked):
        self.filter_manager.toggle_filter(filter_name, checked)
        if not (self.btn_prikaz.isChecked()):
            self.filter_prikaz_id = None
        self.filter_set()
        
       
    
    # устанавливаем фильтр и выставляем указатель на первую строку
    def filter_set(self):
        # сотрудники, военные, уволенные
        filter_spec = f" status = {self.cmb_spec.currentIndex()} " if self.cmb_spec.currentIndex() <3 else None
        # подразделение
        filter_podr = f" podr = '{self.cmb_podr.currentText()}' " if self.cmb_podr.currentIndex() > -1 else None
        #квалификация
        filter_kval = f" kval = '{self.cmb_kval.currentText()}' " if self.cmb_kval.currentIndex() > -1 else None
        # звание
        filter_zvan = f" zvan = '{self.cmb_zvan.currentText()}' " if self.cmb_zvan.currentIndex() > -1 else None
        # приказ
        filter_prikaz = f" idPrikaz = {self.filter_prikaz_id} " if self.filter_prikaz_id else None
        # сдача в текущем году
        filter_tg = " dtSledKval BETWEEN strftime('%Y-01-01', 'now') AND strftime('%Y-12-31', 'now') " if self.btn_srok_tg.isChecked() else None
        # сдача просрочена
        filter_end = " DATE(dtSledKval) < DATE('now') " if self.btn_srok_end.isChecked() else None

        self.data_spec = self.db.read_spec_filter(
                            filter_spec=filter_spec,
                            filter_podr=filter_podr,
                            filter_kval=filter_kval,
                            filter_zvan=filter_zvan,
                            filter_prikaz=filter_prikaz,
                            filter_tg=filter_tg,
                            filter_end=filter_end)
        kol_row = len(self.data_spec)
        self.statusbar.showMessage(f"Записей - {kol_row}")
        self.refresh_grid()
        
        
        

        






    ##############################################################################
    # добавление, сохранение, удаление специалиста
    #############################################################################
    # сохранение сведение о специалисте  
    def save_spec(self):
        

        # проверка на дубликат
        for row in self.data_spec:
            fio = (self.txt_edit_family.text().upper().strip() + " " + 
                self.txt_edit_name.text().upper().strip() + " " + 
                self.txt_edit_lastname.text().upper().strip())
            if str(row[1]).upper().strip() == fio and self.mode_edit == "ADD":
                info_dialog = QMessageBox(self)
                info_dialog.setIcon(QMessageBox.Icon.Information)
                info_dialog.setWindowTitle("Сохранение отменено!")
                info_dialog.setText(f"Специалист есть в базе данных.")
                info_dialog.setStandardButtons(QMessageBox.StandardButton.Ok)
                info_dialog.exec()
                return
        
        # проверяем заполнение полей
        if len(self.txt_edit_family.text().strip()) == 0:
            self.txt_edit_family.setFocus()
            return
        if len(self.txt_edit_name.text().strip()) == 0:
            self.txt_edit_name.setFocus()
            return
        if len(self.txt_edit_lastname.text().strip()) == 0:
            self.txt_edit_lastname.setFocus()
            return
        if self.cmb_edit_podr.currentIndex() < 0:
            self.cmb_edit_podr.showPopup()
            self.cmb_edit_podr.setFocus()
            return
        if len(self.txt_edit_dolzh.text().strip()) == 0:
            self.txt_edit_dolzh.setFocus()
            return
        if self.cmb_edit_zvan.currentIndex() < 0:
            self.cmb_edit_zvan.showPopup()
            self.cmb_edit_zvan.setFocus()
            return        
        if self.cmb_edit_kval.currentIndex() < 0:
            self.cmb_edit_kval.showPopup()
            self.cmb_edit_kval.setFocus()
            return 
        if self.curr_prikaz_id == -1:
            self.btn_edit_prikaz.setFocus()
            return
        if self.cmb_edit_status.currentIndex() < 0:
            self.cmb_edit_status.showPopup()
            self.cmb_edit_status.setFocus()
            return 
 

        # сохраняем данные                        
        if self.mode_edit == 'ADD':
            print('добавляем')
            spec = {
                "id"       : -1,
                "family"   : self.txt_edit_family.text(),
                "name"     : self.txt_edit_name.text(),
                "lastname" : self.txt_edit_lastname.text(),
                "podr"     : self.cmb_edit_podr.currentText(),
                "dolzh"    : self.txt_edit_dolzh.text(),
                "zvan"     : self.cmb_edit_zvan.currentText(),
                "kval"     : self.cmb_edit_kval.currentText(),
                "podtv"    : (1 if self.chk_edit_podtv.isChecked() else 0),
                "id_prikaz": self.curr_prikaz_id,
                "sledKval" : self.curr_prikaz_dt + relativedelta(years=3),
                "status"   : self.cmb_edit_status.currentIndex()
            }  
            
        elif self.mode_edit == 'EDIT':
            print('изменяем')
            spec = {
                "id"       : self.curr_id_spec,
                "family"   : self.txt_edit_family.text(),
                "name"     : self.txt_edit_name.text(),
                "lastname" : self.txt_edit_lastname.text(),
                "podr"     : self.cmb_edit_podr.currentText(),
                "dolzh"    : self.txt_edit_dolzh.text(),
                "zvan"     : self.cmb_edit_zvan.currentText(),
                "kval"     : self.cmb_edit_kval.currentText(),
                "podtv"    : (1 if self.chk_edit_podtv.isChecked() else 0),
                "id_prikaz": self.curr_prikaz_id,
                "sledKval" : self.curr_prikaz_dt + relativedelta(years=3),
                "status"   : self.cmb_edit_status.currentIndex()
            }    
        else:
            return

        # обновляем грид
        self.db.spec_save(spec=spec)
        self.filter_set()


        
        # Восстанавливаем текущую позицию
        if self.mode_edit == "ADD":
            # определяем новый id приказа после добавления
            self.curr_id_spec = max(row[0] for row in self.data_spec)
        
        # Проходим по всем строкам модели для поиска новой позиции отредактированной строки
        for row in range(self.table.model().rowCount(QtCore.QModelIndex())):
            index = self.model.index(row, 0)  # Предполагаем, что id находится в первом столбце
            if self.model.data(index, Qt.ItemDataRole.DisplayRole) == self.curr_id_spec:
                self.table.setCurrentIndex(index)
                self.table.scrollTo(index)  # Прокрутить до выделенной строки
                            
        # выключаем режим редактирования
        self.mode_edit = 'NONE'
        self.frm_grid.setVisible(True)
        self.table.setFocus()
        

    # удаление сотрудника  
    def del_spec(self, curr_id, curr_id_post):
        self.db.spec_del(curr_id)

        # обновляем грид
        self.filter_set()

        
        
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

    
    # отмена сохранения
    def saveno_spec(self):
        self.mode_edit = "NONE"
        self.frm_grid.setVisible(True)
        
        
        
        
    ########################################################################################
    # открытие формы с приказами
    ######################################################################################        
    def open_prikaz(self, mode):
        # настраиваем режим работы (фильтр или выбор на панели редактирования)
        self.w_prikaz.mode_prikaz = mode
        self.w_prikaz.show()    
        

    
    

    ########################################################################################
    # событие перед закрытием приложения
    ######################################################################################   

    def closeEvent(self, event):
        

        if self.mode_edit == "NONE":
            self.db.close_connection()
            event.accept()
        else:
            # Показываем диалог для подтверждения выхода
            reply = QMessageBox.question(
                self,
                "Подтверждение выхода",
                "Вы действительно хотите закрыть приложение?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
                
            if reply == QMessageBox.StandardButton.Yes:
                # Выполняем действия перед закрытием, если нужно
                #print("Сохраняем данные перед выходом...")
                # Разрешаем закрытие
                self.db.close_connection()
                event.accept()
            else:
                # Отменяем закрытие
                event.ignore()
                
                
    ########################################################################################
    # Прием сведений из формы приказов
    ###################################################################################### 
    def update_edit_from_prikaz(self, prikaz_id, prikaz_nom, prikaz_dt):
        self.curr_prikaz_id = prikaz_id
        self.curr_prikaz_dt = prikaz_dt
        self.txt_edit_prikaz.setText(f'№ {prikaz_nom} от {datetime.strftime(prikaz_dt, "%d.%m.%Y")}')
        check_widget(self.txt_edit_prikaz, prikaz_id == -1)

    def update_filter_from_prikaz(self, prikaz_id, prikaz_nom, prikaz_dt):
        if self.btn_prikaz.isChecked():
            self.filter_prikaz_id = prikaz_id
        else:
            self.filter_prikaz_id = None
        self.txt_prikaz.setText(f'№ {prikaz_nom} от {datetime.strftime(prikaz_dt, "%d.%m.%Y")}')
        self.filter_set()



    ########################################################################################
    # Обновляем грид
    ###################################################################################### 
    def refresh_grid(self):
        # настраиваем грид
        column_titles = ["id", "ФИО", "Подразделение", "Должность", "Квалификация", "Следующая"]
        self.model = SpecTableModel(self.data_spec, column_titles)

        self.table.setModel(self.model)
        
        # ширина столбцов
        self.table.setColumnWidth(0, 0)
        self.table.setColumnWidth(1, 300)
        self.table.setColumnWidth(2, 180)
        self.table.setColumnWidth(3, 400)
        self.table.setColumnWidth(4, 100)
        self.table.setColumnWidth(5, 90)  
        
        # режим выделения = строка целиком
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)  
        self.model.layoutChanged.emit()        