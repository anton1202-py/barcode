"""
21.04.2023
v1.2
Добавил создание файла pdf с наклейками для комплетовщиков.
Количество берется из файла Новый файл для производства - Производство + 1
Печатается, когда нажат чекбокс - печать qr-кодов 
Разделил печать для qr поставок на два файла - interface и qrcode_print
Вытащил функцию создания общего файла с наклейками в отдельный файл.
Теперь она импортируется в каждый файл. Экономит место и время.

22.04.2023
v1.3
Добавил чекбокс на возможность создания файла pdf с наклейками для комплетовщиков.
Он становится активным, если нажат чекбокс для печати qr-кодов
Добавил окно об ошибке, если артикула нет в базе данных.

24.04.2023
v1.4
Изменил создание excel-файла для сборки
Изменил алгоритм подсчета печати штрихкодов для QR-поставок

26.04.2023
v2.1
Добавил большой блок: Сборка FBS с формированием остатков
Исправил баги с печатью штрихкодов в файле barcode_print

27.04.2023
v2.2
Добавил часть функционала - создание актов по трем маркетплесов.
Трудное создание акта для Озона. CSV сохраняется в XLSX, форматируется
в зависимости от данных, потом переводится в PDF

04.05.2023
v2.3
Добавил визуальный контроль загружаемых файлов.
Исправил баг с количеством товара для FBS на каждом маркетплейсе.
(Поменял метод подсчета товара для таблиц)

06.05.2023
v2.4
Перенес все excel файлы из папки programm_data в dropbox.
Теперь должна работать центральная база данных для всех
пользователей приложения

10.05.2023
v2.5
Сделал сохранение пути до выбранной папки.
Добавил модуль со штрихкодом папки.
Добавил модуль с печатью только qr-кода.

11.05.2023
v3.1
Обновил интерфейс

12.05.2023
v3.2
Прибрался в коде. Повторяющиеся части выделил в отдельные функции.
Они лежит в файле helper_func.

Добавил d'n'd окно в файл barcode_print. Убрал кнопку выбора 
файла с количество товара.

Поменял название столбцов главных файлов, потому что они изменились
в выгрузке с ВБ.

Добавил возможность в мега режиме (FBS режим) выбирать возможность 
сразу печатать акт поставке в файл с этикетками.

13.05.2023
v3.3
Подправил файл с печатью qr-кодов для коробок - добавил выбор места сохранения
из сохраненной в кеше папки.
Исправлена опечатка в файле barcode_print

15.05.2023
v3.4
only_qrcode_print:
Добавил возможность тиражировать этикетки.
Исправил баг: не печатался файл, если во входном был один qrcode.
"""


import sys
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import QSize
from PyQt6.QtGui import QIcon

from barcode_print import BarcodePrint
from box_barcode import BarcodeBoxPrintToPdf
from draganddrop import DropMainWindow
import dropbox
from only_qrcode_print import OnlyQrcodePrint
from qrcode_print import QrcodePrint

version = 'v3.3'
refresh_token_db = 'kcjuYCMJ958AAAAAAAAAAbSglZkCTY50p7ksrTLt2e5d7zJM5_uf1D6URbTOwgQJ'
app_key_db = '3rvhk6f0pjdksc8'
app_secret_db = '3a1pe948esjx39d'
dbx_db = dropbox.Dropbox(oauth2_refresh_token=refresh_token_db,
                      app_key=app_key_db,
                      app_secret=app_secret_db)

class stackedExample(QWidget):
    def __init__(self):
        super(stackedExample, self).__init__()
        self.setStyleSheet("background-color: #2a2c2e;")
        self.resize(1000, 800)
        self.setWindowTitle(f"Barcode System {version}")

        self.main = QVBoxLayout()
        self.main.setContentsMargins(0, 20, 20, 20)
        self.setLayout(self.main)
        
        self.main_layout = QHBoxLayout()
        self.main_layout.addStretch()
        self.main.addLayout(self.main_layout)

        self.menu_button_widget = QHBoxLayout()
        self.main_layout.addLayout(self.menu_button_widget)
        self.menu_button_widget.addStretch(1)
        self.menu_layout = QVBoxLayout()
        self.main_layout.addLayout(self.menu_layout)

        self.main_body = QVBoxLayout()
        self.main_layout.addLayout(self.main_body)

        menu_button_fbs_icon = 'programm_data/icons/barcode-orange.svg'
        self.menu_button_fbs = QPushButton("    Сборка FBS с формированием остатков")
        self.menu_button_fbs.setIcon(QIcon(menu_button_fbs_icon))
        self.menu_button_fbs.setIconSize(QSize(40,40))
        self.menu_button_fbs.setStyleSheet(self.stylesheet_family_1())
        self.menu_layout.addWidget(self.menu_button_fbs)

        menu_button_bc_icon = 'programm_data/icons/barcode-with-border-orange.svg'
        self.menu_button_bc = QPushButton("    Печать штрихкодов для товара")
        self.menu_button_bc.setIcon(QIcon(menu_button_bc_icon))
        self.menu_button_bc.setIconSize(QSize(40,40))
        self.menu_button_bc.setStyleSheet(self.stylesheet_family_1())
        self.menu_layout.addWidget(self.menu_button_bc)


        menu_button_qc_icon = 'programm_data/icons/qr-code-orange.svg'
        self.menu_button_qc = QPushButton("    Печать qr-кодов для коробок")
        self.menu_button_qc.setIcon(QIcon(menu_button_qc_icon))
        self.menu_button_qc.setIconSize(QSize(40,40))
        self.menu_button_qc.setStyleSheet(self.stylesheet_family_1())
        self.menu_layout.addWidget(self.menu_button_qc)

        menu_button_box_icon = 'programm_data/icons/expand-orange.svg'
        self.menu_button_box = QPushButton("    Поставки на WB")
        self.menu_button_box.setIcon(QIcon(menu_button_box_icon))
        self.menu_button_box.setIconSize(QSize(40,40))
        self.menu_button_box.setStyleSheet(self.stylesheet_family_1())
        self.menu_layout.addWidget(self.menu_button_box)

        menu_button_qr_only_icon = 'programm_data/icons/qr-orange.svg'
        self.menu_button_qr_only = QPushButton("    Компановка наклеек (18 на лист)")
        self.menu_button_qr_only.setIcon(QIcon(menu_button_qr_only_icon))
        self.menu_button_qr_only.setIconSize(QSize(40,40))
        self.menu_button_qr_only.setStyleSheet(self.stylesheet_family_1())
        self.menu_layout.addWidget(self.menu_button_qr_only)
        self.menu_layout.addStretch()
        self.main_layout.addStretch(1)

        self.menu_button_fbs.clicked.connect(self.display)
        self.menu_button_bc.clicked.connect(self.display_1)
        self.menu_button_qc.clicked.connect(self.display_2)
        self.menu_button_box.clicked.connect(self.display_3)
        self.menu_button_qr_only.clicked.connect(self.display_4)

        self.stack1 = DropMainWindow()
        self.stack2 = BarcodePrint()
        self.stack3 = QrcodePrint()
        self.stack4 = BarcodeBoxPrintToPdf()
        self.stack5 = OnlyQrcodePrint()
      
        self.stacked_widget = QStackedWidget(self)
        self.stacked_widget.setStyleSheet("background-color: #2a2c2e;")
        self.stacked_widget.addWidget(self.stack1)
        self.stacked_widget.addWidget(self.stack2)
        self.stacked_widget.addWidget(self.stack3)
        self.stacked_widget.addWidget(self.stack4)
        self.stacked_widget.addWidget(self.stack5)
        self.main_body.addWidget(self.stacked_widget)

    def stylesheet_family_1(self):
        return'''
        QPushButton {
            font: 12pt; color: #f7f7f7;
            width: 380px;
            height: 80px;
            padding-left: 10px;
            text-align: left;
            }
            QPushButton::hover {
                    background-color: #f74a00;}
            QPushButton::pressed {
                    background-color: #ff1e00;}
                '''
    def display(self):
        self.stacked_widget.setCurrentIndex(0)

    def display_1(self):
        self.stacked_widget.setCurrentIndex(1)
    
    def display_2(self):
        self.stacked_widget.setCurrentIndex(2)
    
    def display_3(self):
        self.stacked_widget.setCurrentIndex(3)
    
    def display_4(self):
        self.stacked_widget.setCurrentIndex(4)
		
def main():
    app = QApplication(sys.argv)
    window = stackedExample()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
   main()