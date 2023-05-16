import glob
import os
from pathlib import Path
from pdf2image import convert_from_path
import shutil
from PIL.ImageQt import ImageQt
from PyQt6.QtCore import QSettings, Qt
from PyQt6.QtGui import QPainter
from PyQt6.QtPrintSupport import QPrintDialog, QPrinter
from PyQt6.QtWidgets import (QApplication, QMainWindow, QCheckBox,
                             QGridLayout, QLabel, QLineEdit,
                             QListWidget,QPushButton,
                             QWidget, QVBoxLayout)
import tempfile

from helper_func import qrcode_print_to_file_main


class OnlyQrcodePrint(QMainWindow):
    def __init__(self):
        from main_file import version
        super().__init__()
        self.list_for_print = []
        self.settings = QSettings('settings.ini', QSettings.Format.IniFormat)
        self.settings.setFallbacksEnabled(False)
        self.setStyleSheet("background-color: #2a2c2e; font: 12pt; color: #f7f7f7;")
        self.setWindowTitle(f'Сборка FBS с формированием остатков {version}')
        self.resize(500, 300)
        self.setContentsMargins(20, 20, 20, 20)

        # Даем разрешение на Drop
        self.setAcceptDrops(True)
        self.files_list = []
        self.list_files = QListWidget()
        self.list_files.setStyleSheet("font: 10pt; color: #f7f7f7;")
        self.label_total_files = QLabel()

        main_layout = QVBoxLayout()
        main_layout.addWidget(QLabel('Перетащите PDF файл в это окно:'))
        main_layout.addWidget(self.list_files)
        main_layout.addWidget(self.label_total_files)


        check_content_layout = QGridLayout()
        main_layout.addLayout(check_content_layout)


        central_widget = QWidget()
        central_widget.setLayout(main_layout)


        self.label_assert_error = QLabel()
        main_layout.addWidget(self.label_assert_error)

        self.check_box_amount = QCheckBox('Тиражировать файл')
        main_layout.addWidget(self.check_box_amount)
        self.check_box_amount.stateChanged.connect(self.checkbox_status)


        self.label_amount = QLabel('Введите количество этикеток')
        self.label_amount.setStyleSheet(
            '''QLabel {
                font: 14pt;
                color: #f7f7f7;
                margin-top: 40px;
                }''')
        main_layout.addWidget(self.label_amount)

        self.input_amount = QLineEdit()
        self.input_amount.setStyleSheet(
            '''QLineEdit {
                font: 14pt;
                width: 350;
                height: 30;
                color: #f7f7f7;
                }
                QLineEdit:disabled {
                background-color: #303030;
                color: #707070}''')
        self.input_amount.setEnabled(False)
        self.input_amount.textChanged.connect(self.get_box_amount_text)
        main_layout.addWidget(self.input_amount)
        self.prog_button = QPushButton('Преобразовать файл')
        self.prog_button.setStyleSheet(
            '''QPushButton {
                border: 2px solid #f74a00;
                font: 12pt; color: #f7f7f7;
                border-radius: 8px;
                margin-top: 20px;
                height: 40px;
                }
                QPushButton::hover {
                background-color: #f74a00;
                font: 12pt; color: #f7f7f7;
                border-radius: 8px;
                height: 40px;
                }
                QPushButton:disabled {
                        background-color: #575757;
                        border: 0px;}
                QPushButton::pressed {
                        background-color: #ff1e00;}
                ''')
        self.prog_button.setEnabled(True)
        self.prog_button.pressed.connect(self.work_func)
        main_layout.addWidget(self.prog_button)
        
        self.setCentralWidget(central_widget)
        self._update_states()

    def _update_states(self):
        self.label_total_files.setText('Загружено файлов: {}'.format(self.list_files.count()))


    def dragEnterEvent(self, event):
        # Тут выполняются проверки и дается (или нет) разрешение на Drop
        mime = event.mimeData()
        # Если перемещаются ссылки
        if mime.hasUrls():
            # Разрешаем
            event.acceptProposedAction()
        
    def dropEvent(self, event):
        # Обработка события Drop
        for url in event.mimeData().urls():
            file_name = url.toLocalFile()
            self.list_files.addItem(file_name)
            self.files_list.append(file_name)
        self._update_states()
        return super().dropEvent(event)

    def remove(self):
        current_row = self.list_files.currentRow()
        if current_row >= 0:
            current_item = self.list_files.takeItem(current_row)
            del current_item
            self.files_list.clear()
            self._update_states()

    def func_returne_list(self):
        return self.files_list

    def keyPressEvent(self, e):
        """Удаляет строки из списка файлов при нажатии на delete"""
        if e.key() == Qt.Key.Key_Delete:
            self.remove()

    def checkbox_status(self):
        """Проверяет статус чекбокса"""
        if self.check_box_amount.isChecked():
            self.input_amount.setEnabled(True)
        else:
            self.input_amount.setEnabled(False)

    def get_box_amount_text(self):
        text = self.input_amount.text()
        try:
            return int(text)
        except:
            self.label_assert_error.setText(
                '<h3 style="color: rgb(250, 55, 55);">Введите целое число!</h3>'
                )

    def work_func(self):
        if len(self.func_returne_list()) == 0:
            self.label_assert_error.setText(
                '''<h3 style="color: rgb(250, 55, 55);">
                Выберете PDF файл с Qr-кодами</h3>''')
        else:
            self.label_assert_error.clear()
            BackendPart.qrcode_print_to_file(self)
            BackendPart.collect_final_pdf(self)
            BackendPart.printDialog(self)
            self.list_files.clear()
            self.files_list.clear()
            self._update_states()


class BackendPart(OnlyQrcodePrint):
    def __init__(self):
        super().__init__()
        self.list_for_print = []
        self.file_name = ''

    def qrcode_print_to_file(self):
        """Создает QR коды в необходимом формате"""

        files_list = self.func_returne_list()
        global filename_pdf
        for file in files_list:
            path = Path(file)
            filename_pdf = os.path.basename(path).split('.')[0]
            self.list_for_print = qrcode_print_to_file_main(file, filename_pdf)
            self.file_name = f'{path.parent}/{filename_pdf} resave.pdf'
        list_for_print_raw = []
        # --------- Тиражируем этикетки, если чекбокс нажат ---------
        if self.check_box_amount.isChecked():
            amount = self.get_box_amount_text()
            for i in self.list_for_print:
                for f in range(amount):
                    list_for_print_raw.append(i)
            self.list_for_print = list_for_print_raw

    def collect_final_pdf(self):
        """Делает новый PDF файл с QR-кодами"""
        from helper_func import print_barcode_to_pdf
        print_barcode_to_pdf(self.list_for_print, self.file_name)

        dir = f'cache_dir_3/{filename_pdf}/'
        filelist = glob.glob(os.path.join(dir, "*"))
        for f in filelist:
            try:
                os.remove(f)
            except Exception:
                print(' ')
    
        dir = 'cache_dir_2/'
        filelist = glob.glob(os.path.join(dir, "*"))
        for f in filelist:
            try:
                os.remove(f)
            except Exception:
                print(' ')
        self.list_for_print.clear()


    def printDialog(self):
        """Печать файла"""
        if self.file_name:
            #print('self.file_name', self.file_name)
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            dialog = QPrintDialog(printer, self)
            if dialog.exec():
                QPrintDialog.accepted
                with tempfile.TemporaryDirectory() as path:
                    images = convert_from_path(self.file_name, dpi=300, output_folder=path, poppler_path=r'poppler-0.68.0\bin')
                    painter = QPainter()
                    painter.begin(printer)
                    for i, image in enumerate(images):
                        if i > 0:
                            printer.newPage()
                        rect = painter.viewport()
                        qtImage = ImageQt(image)
                        qtImageScaled = qtImage.scaled(rect.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                        painter.drawImage(rect, qtImageScaled)
                    painter.end()
        else:
            pass

        folder = 'cache_dir_3/'
        for filename in glob.glob(os.path.join(folder, "*")):
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(filename) or os.path.islink(filename):
                    os.unlink(filename)
                elif os.path.isdir(filename):
                    shutil.rmtree(filename)
            except Exception as e:
                print('')


if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    w = OnlyQrcodePrint()
    w.show()
    sys.exit(app.exec())