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

"""



from PyQt6.QtWidgets import QApplication, QPushButton, QWidget, QVBoxLayout

from barcode_print import BarcodePrint
from box_barcode import BarcodeBoxPrintToPdf
from draganddrop import DropMainWindow
from only_qrcode_print import OnlyQrcodePrint
from qrcode_print import QrcodePrint


version = 'v2.5'


class MainMenuWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.barcode_window = BarcodePrint()
        self.qrcode_window = QrcodePrint()
        self.fbs_window = DropMainWindow()
        self.barcode_for_box = BarcodeBoxPrintToPdf()
        self.only_qrcode_print = OnlyQrcodePrint()

        self.setStyleSheet("background-color: #0f1314;")
        self.resize(350, 400)
        self.setWindowTitle(f"Иннотрейд {version}")

        self.menu_layout = QVBoxLayout()
        self.menu_layout.setContentsMargins(5, 5, 5, 5)

        self.setLayout(self.menu_layout)
        
        self.fbs_btn = QPushButton("Сборка FBS с формированием остатков")
        self.fbs_btn.setStyleSheet(self.stylesheet_qpushButton())
        self.fbs_btn.clicked.connect(self.toggle_window3)
        self.menu_layout.addWidget(self.fbs_btn)


        self.barcode_btn = QPushButton("Печать штрихкодов для товара")
        self.barcode_btn.setStyleSheet(self.stylesheet_qpushButton())
        self.barcode_btn.clicked.connect(self.toggle_window1)
        self.menu_layout.addWidget(self.barcode_btn)

        self.btn3 = QPushButton("Печать qr-кодов для коробок")
        self.btn3.setStyleSheet(self.stylesheet_qpushButton())
        self.btn3.clicked.connect(self.toggle_window2)
        self.menu_layout.addWidget(self.btn3)

        self.btn4 = QPushButton("Печать этикеток для коробок")
        self.btn4.setStyleSheet(self.stylesheet_qpushButton())
        self.btn4.clicked.connect(self.toggle_window4)
        self.menu_layout.addWidget(self.btn4)

        self.btn5 = QPushButton("Компановка наклеек (18 на лист)")
        self.btn5.setStyleSheet(self.stylesheet_qpushButton())
        self.btn5.clicked.connect(self.toggle_window5)
        self.menu_layout.addWidget(self.btn5)
        
    
    def stylesheet_qpushButton(self):
        return '''QPushButton {
                background-color: #ff8126;
                font: 16pt; color: #f7f7f7;
                height: 80px;
                }
                QPushButton::hover { 
                background-color : #f74a00;
                }
                QPushButton::pressed {
                    background-color: #ff1e00;}
                '''

    def toggle_window1(self, checked):
        if self.barcode_window.isVisible():
            self.barcode_window.hide()

        else:
            self.barcode_window.show()

    def toggle_window2(self, checked):
        if self.qrcode_window.isVisible():
            self.qrcode_window.hide()

        else:
            self.qrcode_window.show()
    
    def toggle_window3(self, checked):
        if self.fbs_window.isVisible():
            self.fbs_window.hide()

        else:
            self.fbs_window.show()
    
    def toggle_window4(self, checked):
        if self.barcode_for_box.isVisible():
            self.barcode_for_box.hide()

        else:
            self.barcode_for_box.show()
    
    def toggle_window5(self, checked):
        if self.only_qrcode_print.isVisible():
            self.only_qrcode_print.hide()

        else:
            self.only_qrcode_print.show()


if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    w = MainMenuWindow()
    w.show()
    sys.exit(app.exec())
