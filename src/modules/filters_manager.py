from PyQt6.QtGui import QFont

class FilterManager:
    def __init__(self, filters_config, ui_elements):
        """
        Инициализирует менеджер фильтров.
        :param filters_config: Словарь с состояниями фильтров.
        :param ui_elements: Словарь с привязкой фильтров к элементам UI.
        """
        self.filters = filters_config
        self.ui_elements = ui_elements

    def toggle_filter(self, filter_name, is_checked):
        """
        Включает или выключает фильтр и обновляет UI.
        :param filter_name: Название фильтра.
        :param is_checked: Состояние фильтра (True/False).
        """
        self.filters[filter_name] = is_checked

        # Обновление стиля кнопки
        font = QFont('Arial', 10)
        font.setBold(is_checked)
        button = self.ui_elements[f'btn_{filter_name}']
        button.setFont(font)

        # Включение/выключение связанного элемента (комбобокс или другая кнопка)
        linked_element = self.ui_elements.get(f'cmb_{filter_name}') or self.ui_elements.get(f'btn_sel_{filter_name}')
        if linked_element:
            linked_element.setEnabled(is_checked)
            if is_checked and hasattr(linked_element, 'showPopup'):
                linked_element.showPopup()
            elif (not is_checked) and hasattr(linked_element, 'showPopup'):
                linked_element.setCurrentIndex(-1)
            elif (not is_checked) and (not hasattr(linked_element, 'showPopup')):
                linked_element2 = self.ui_elements.get(f'txt_{filter_name}')
                if linked_element2:
                    linked_element2.setText("")

    def reset_filters(self):
        """
        Сбрасывает все фильтры.
        """
        for filter_name in self.filters.keys():
            self.toggle_filter(filter_name, False)
