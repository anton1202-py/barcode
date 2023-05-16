from PyQt6.QtWidgets import (QCheckBox, QApplication, QMainWindow, QLabel, QWidget, QPushButton,
                             QVBoxLayout, QListWidget, QGridLayout, QRadioButton)
from PyQt6.QtCore import Qt
from contextlib import closing
import io
import shutil
import glob
import os


class DropMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        from main_file import dbx_db, version

        self.dbx = dbx_db
        self.setStyleSheet("background-color: #2a2c2e; font: 12pt; color: #f7f7f7;")
        self.setWindowTitle(f'Сборка FBS с формированием остатков {version}')
        self.resize(500, 800)
        self.setContentsMargins(20, 20, 20, 20)

        # Даем разрешение на Drop
        self.setAcceptDrops(True)
        self.files_list = []
        self.xls_files_list = []
        self.list_files = QListWidget()
        self.list_files.setStyleSheet("font: 10pt; color: #f7f7f7;")
        self.label_total_files = QLabel()

        main_layout = QVBoxLayout()
        main_layout.addWidget(QLabel('Перетащите файлы в это окно:'))
        main_layout.addWidget(self.list_files)
        main_layout.addWidget(self.label_total_files)

        self.checkbox_add_qrcode = QCheckBox('Добавлять qr-код к наклейкам', self)
        main_layout.addWidget(self.checkbox_add_qrcode)
        self.checkbox_add_qrcode.setStyleSheet('''
                    QCheckBox {
                        margin-bottom: 20px;
                    }''')

        self.chech_checkbox = QCheckBox('Проверить наличие файлов', self)
        main_layout.addWidget(self.chech_checkbox)
        self.chech_checkbox.setStyleSheet('''
                    QCheckBox {
                        margin-bottom: 20px;
                    }''')
        self.chech_checkbox.stateChanged.connect(self.check_that_files_exists)


        check_content_layout = QGridLayout()
        main_layout.addLayout(check_content_layout)
        
        
        self.wb_qrcode = QCheckBox('WB QR-коды (pdf)', self)
        self.wb_qrcode.setEnabled(False)
        self.wb_qrcode.setStyleSheet(self.stylesheet_qcheckBox())
        check_content_layout.addWidget(self.wb_qrcode, 1, 0)

        self.wb_act = QCheckBox('WB акт поставки (pdf)', self)
        self.wb_act.setEnabled(False)
        self.wb_act.setStyleSheet(self.stylesheet_qcheckBox())
        check_content_layout.addWidget(self.wb_act, 2, 0)

        self.wb_list = QCheckBox('WB лист подбора (pdf)', self)
        self.wb_list.setEnabled(False)
        self.wb_list.setStyleSheet(self.stylesheet_qcheckBox())
        check_content_layout.addWidget(self.wb_list, 3, 0)

        self.wb_task = QCheckBox('WB задание с артикулами (xlsx)', self)
        self.wb_task.setEnabled(False)
        self.wb_task.setStyleSheet(self.stylesheet_qcheckBox())
        check_content_layout.addWidget(self.wb_task, 4, 0)

        self.ozon_act = QCheckBox('OZON акт поставки (pdf)', self)
        self.ozon_act.setEnabled(False)
        self.ozon_act.setStyleSheet(self.stylesheet_qcheckBox())
        check_content_layout.addWidget(self.ozon_act, 5, 0)

        self.ozon_tickets = QCheckBox('OZON этикетки (pdf)', self)
        self.ozon_tickets.setEnabled(False)
        self.ozon_tickets.setStyleSheet(self.stylesheet_qcheckBox())
        check_content_layout.addWidget(self.ozon_tickets, 6, 0)

        self.ozon_task = QCheckBox('OZON задание с артикулами (csv)', self)
        self.ozon_task.setEnabled(False)
        self.ozon_task.setStyleSheet(self.stylesheet_qcheckBox())
        check_content_layout.addWidget(self.ozon_task, 1, 1)

        self.yandex_act = QCheckBox('YANDEX акт поставки (pdf)', self)
        self.yandex_act.setEnabled(False)
        self.yandex_act.setStyleSheet(self.stylesheet_qcheckBox())
        check_content_layout.addWidget(self.yandex_act, 2, 1)

        self.yandex_tickets = QCheckBox('YANDEX этикетки (pdf)', self)
        self.yandex_tickets.setEnabled(False)
        self.yandex_tickets.setStyleSheet(self.stylesheet_qcheckBox())
        check_content_layout.addWidget(self.yandex_tickets, 3, 1)

        self.yandex_task = QCheckBox('YANDEX задание с артикулами (xlsx)', self)
        self.yandex_task.setEnabled(False)
        self.yandex_task.setStyleSheet(self.stylesheet_qcheckBox())
        check_content_layout.addWidget(self.yandex_task, 4, 1)

        self.yandex_task_pdf = QCheckBox('YANDEX задание с артикулами (pdf)', self)
        self.yandex_task_pdf.setEnabled(False)
        self.yandex_task_pdf.setStyleSheet(self.stylesheet_qcheckBox())
        check_content_layout.addWidget(self.yandex_task_pdf, 5, 1)


        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.prog_button = QPushButton('Преобразовать файлы')
        self.prog_button.setStyleSheet(
            '''QPushButton {
                border: 2px solid #f74a00;
                font: 12pt; color: #f7f7f7;
                border-radius: 8px;
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
        self.prog_button.setEnabled(False)
        self.prog_button.pressed.connect(self.test_func)
        main_layout.addWidget(self.prog_button)
        
        self.setCentralWidget(central_widget)
        self._update_states()

        self.label = QLabel("Выберите юр. лицо для этикетки")
        self.label.setStyleSheet(
            '''QLabel {
                font: 14pt;
                color: #f7f7f7;
                margin-top: 40px;
                }''')
        main_layout.addWidget(self.label)

        self.radiobtn1 = QRadioButton("ООО Иннотрейд", self)
        self.radiobtn1.setStyleSheet(
            '''QRadioButton {font: 12pt; color: #f7f7f7;}
               QRadioButton:disabled {color: #8c8c8c}''')
        self.radiobtn1.setEnabled(False)

        main_layout.addWidget(self.radiobtn1)

        self.radiobtn2 = QRadioButton("ИП Караваев", self)
        self.radiobtn2.setStyleSheet(
            '''QRadioButton {font: 12pt; color: #f7f7f7;}
               QRadioButton:disabled {color: #8c8c8c}''')
        self.radiobtn2.toggled.connect(self.showDetails)
        self.radiobtn2.setEnabled(False)
        main_layout.addWidget(self.radiobtn2)

        self.label1 = QLabel("", self)
        main_layout.addWidget(self.label1)

        self.btn = QPushButton("Сформировать итоговый PDF")
        self.btn.setStyleSheet(
            '''QPushButton {
                background-color: #f74a00;
                font: 12pt; color: #f7f7f7;
                border-radius: 8px;
                height: 40px;
                margin-top: 30px;
                }
                QPushButton:disabled {
                        background-color: #575757;
                        border: 0px;}
                QPushButton::pressed {
                        background-color: #ff1e00;}
                ''')
        self.btn.pressed.connect(self.check_push_radiobutton)
        self.btn.setEnabled(False)
        main_layout.addWidget(self.btn)

    def stylesheet_qcheckBox(self):
        return'''
            QCheckBox::indicator:checked{
            	background-color: #f74a00;
                color: #ffffff
            }
            QCheckBox:disabled {
            	color: #8c8c8c
            }
            '''

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

    def check_that_files_exists(self):
        from all_mkpl_analyz import CheckAllContent
        CheckAllContent.pdf_file_analyze(self)
        self.prog_button.setEnabled(True)

    def func_returne_list(self):
        return self.files_list
    
    def remove(self):
        current_row = self.list_files.currentRow()
        if current_row >= 0:
            current_item = self.list_files.takeItem(current_row)
            del current_item

    def keyPressEvent(self, e):
        """Удаляет строки из списка файлов при нажатии на delete"""
        if e.key() == Qt.Key.Key_Delete:
            self.remove()
        if self.list_files.count() == 0:
            self.prog_button.setEnabled(False)

    def stream_dropbox_file(self, path):
        _,res=self.dbx.files_download(path)
        with closing(res) as result:
            byte_data=result.content
            return io.BytesIO(byte_data)
        
    def showDetails(self):
        """Логика работы радиокнопок"""
        if self.radiobtn1.isChecked():
            path = '/DATABASE/Ночники ООО.xlsx'
            self.main_file = self.stream_dropbox_file(path)
            path_for_ticket = '/DATABASE/helper_files/Печать Иннотрейд.xlsx'
            return self.stream_dropbox_file(path_for_ticket), self.main_file
        elif self.radiobtn2.isChecked():
            path = '/DATABASE/Ночники ИП.xlsx'
            self.main_file = self.stream_dropbox_file(path)
            path_for_ticket = '/DATABASE/helper_files/Печать Караваев.xlsx'
            return self.stream_dropbox_file(path_for_ticket), self.main_file
        else:
            print('Выберите юр. лицо!')

    def check_push_radiobutton(self):
        """Обрабатывает ошибку не выбора радиокнопок"""
        # checking if it is checked
        if self.radiobtn1.isChecked() or self.radiobtn2.isChecked():
            self.print_pdf()
        else:
            # changing text of label
            self.label1.setText(
                '<h3 style="color: rgb(250, 55, 55);">Выберете юр. лицо!</h3>'
                )
    
    def test_func(self):
        from all_mkpl_analyz import RenamePdf, CreatePivoteExcel, PrintTicket, CheckAllContent
        RenamePdf.pdf_file_analyze(self)
        RenamePdf.create_ozone_selection_sheet_pdf(self)
        CreatePivoteExcel.xls_file_analyze(self)
        CreatePivoteExcel.create_pivot_xls(self)
        self.radiobtn1.setEnabled(True)
        self.radiobtn2.setEnabled(True)
        self.btn.setEnabled(True)
    
    def print_pdf(self):
        from all_mkpl_analyz import RenamePdf, CreatePivoteExcel, PrintTicket
        PrintTicket.qrcode_print_to_file(self)
        PrintTicket.create_list_barcode(self)
        PrintTicket.create_pdf_file_ticket_for_complect(self)
        PrintTicket.print_barcode_in_pdf(self)

    dir = 'cache_dir_2/'
    filelist = glob.glob(os.path.join(dir, "*"))
    for f in filelist:
        os.remove(f)
    folder = 'cache_dir_3/'
    for filename in glob.glob(os.path.join(folder, "*")):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(filename) or os.path.islink(filename):
                os.unlink(filename)
            elif os.path.isdir(filename):
                shutil.rmtree(filename)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (filename, e))
    # Удаление кеша из папки с кешем
    dir = 'cache_dir/'
    filelist = glob.glob(os.path.join(dir, "*"))
    for f in filelist:
        os.remove(f)



if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    w = DropMainWindow()
    w.show()
    sys.exit(app.exec())