from PyQt6.QtWidgets import QComboBox

########################################################################################
# закраска фона не заполненных полей
######################################################################################   
    
def set_widget_background(widget, condition):
    """Устанавливает фон виджета в зависимости от условия."""
    if isinstance(widget, QComboBox):
        style = "background: yellow; selection-background-color: gray;" if condition else ""
    else:                
        style = "background: yellow;" if condition else ""
    widget.setStyleSheet(f"QWidget{{{style}}}")
    

def check_widget( widget, condition):
    set_widget_background(widget, condition)