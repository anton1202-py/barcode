from PyQt6.QtWidgets import (QCheckBox, QWidget, QLineEdit,
    QPushButton, QVBoxLayout, QRadioButton,
    QLabel, QGridLayout, QFileDialog
)
from PyQt6.QtCore import QSettings, Qt
from pathlib import Path
from contextlib import closing
import io


class InterfaceQrcode(QWidget):
    
    def __init__(self):
        super().__init__()
        from main_file import dbx_db, version

        self.dbx = dbx_db
        self.settings = QSettings('settings.ini', QSettings.Format.IniFormat)
        self.settings.setFallbacksEnabled(False)
        self.setStyleSheet("background-color: #2a2c2e;")
        self.resize(350, 400)
        self.setWindowTitle(f"Иннотрейд - печать QR-кодов {version}")
 
        layout = QVBoxLayout()
        layout.setContentsMargins(50, 50, 50, 50)
        self.setLayout(layout)
        self.label = QLabel("Выберите юр. лицо для этикетки")
        self.label.setStyleSheet(
            '''QLabel {
                font: 14pt;
                color: #f7f7f7;
                height: 100px;
                margin-top: 0px;
                }''')
        layout.addWidget(self.label)

        self.radiobtn1 = QRadioButton("ООО Иннотрейд", self)
        self.radiobtn1.setStyleSheet(
            '''QRadioButton {font: 12pt; color: #f7f7f7;}''')
        self.radiobtn1.toggled.connect(self.radiobutton_logic)
        layout.addWidget(self.radiobtn1)

        self.radiobtn2 = QRadioButton("ИП Караваев", self)
        self.radiobtn2.setStyleSheet(
            '''QRadioButton {font: 12pt; color: #f7f7f7;}''')
        self.radiobtn2.toggled.connect(self.radiobutton_logic)
        layout.addWidget(self.radiobtn2)

        self.label_1 = QLabel("Введите количество коробок")
        self.label_1.setStyleSheet(
            '''QLabel {
                font: 14pt;
                color: #f7f7f7;
                margin-top: 40px;
                }''')
        layout.addWidget(self.label_1)

        self.input1 = QLineEdit()
        self.input1.setStyleSheet(
            '''QLineEdit {
                font: 14pt;
                width: 350;
                height: 30;
                color: #f7f7f7;
                }''')
        
        self.input1.textChanged.connect(self.get_box_amount_text)
        layout.addWidget(self.input1, alignment= Qt.AlignmentFlag.AlignCenter)

        self.error_int_box = QLabel("", self)
        layout.addWidget(self.error_int_box)

        self.check1 = QCheckBox("Изменить количество товара в коробке")
        self.check1.setStyleSheet(
            '''QCheckBox {font: 14pt; color: #f7f7f7; margin-top: 40px;}''')
        self.check1.stateChanged.connect(self.checkbox_status)
        layout.addWidget(self.check1)

        self.input_2 = QLineEdit('20')
        self.input_2.setStyleSheet(
            '''QLineEdit {
                font: 14pt;
                width: 350;
                height: 30;
                color: #f7f7f7;
                }
                QLineEdit:disabled {
                        background-color: #2e2e2e;
                        color: #707070}''')
        self.input_2.setEnabled(False)
        self.input_2.textChanged.connect(self.amount_of_product_text)
        layout.addWidget(self.input_2, alignment= Qt.AlignmentFlag.AlignCenter)

        self.error_int_product = QLabel("", self)
        layout.addWidget(self.error_int_product)


        self.error_label = QLabel("", self)
        layout.addWidget(self.error_label)

        self.btn = QPushButton("Напечатать QR-коды для коробок")
        self.btn.setStyleSheet(
            '''QPushButton {
                background-color: #f74a00;
                font: 12pt; color: #f7f7f7;
                border-radius: 8px;
                height: 40px;
                margin-top: 30px;
                }
                QPushButton::pressed {
                        background-color: #ff1e00;}
                ''')
        self.btn.pressed.connect(self.check_all_data)
        layout.addWidget(self.btn)

    def checkbox_status(self):
        """Проверяет статус чекбокса"""
        if self.check1.isChecked():
            self.input_2.setEnabled(True)
        else:
            self.input_2.setEnabled(False)

    def get_box_amount_text(self):
        text = self.input1.text()
        try:
            return int(text)
        except:
            self.error_int_box.setText(
                '<h3 style="color: rgb(250, 55, 55);">Введите целое число!</h3>'
                )

    def amount_of_product_text(self):
        text = self.input_2.text()
        try:
            return int(text)
        except:
            self.error_int_product.setText(
                '<h3 style="color: rgb(250, 55, 55);">Введите целое число!</h3>'
                )

    def stream_dropbox_file(self, path1):
        _,res=self.dbx.files_download(path1)
        with closing(res) as result:
            byte_data=result.content
            return io.BytesIO(byte_data)

    def radiobutton_logic(self):
        """Логика работы радиокнопок"""
        if self.radiobtn1.isChecked():
            id_seller = 36555
            path = '/DATABASE/helper_files/for_number_box_innotrade.xlsx'
            counter_file_for_box_number = self.stream_dropbox_file(path)
            return 'Innotreid', id_seller, counter_file_for_box_number, path
        elif self.radiobtn2.isChecked():
            id_seller = 484915
            path = '/DATABASE/helper_files/for_number_box_karavaev.xlsx'
            counter_file_for_box_number = self.stream_dropbox_file(path)
            return 'Karavaev', id_seller, counter_file_for_box_number, path
        else:
            return('Выберите юр. лицо!')

    def check_all_data(self):
        """Обрабатывает ошибку не выбора радиокнопок"""
        # checking if it is checked
  
        if self.radiobtn1.isChecked() or self.radiobtn2.isChecked() and self.input1.text():
            self.print_qrcode_to_pdf()
        elif self.radiobtn1.isChecked() or self.radiobtn2.isChecked() and not self.input1.text():
            # changing text of label
            self.error_label.setText(
                '<h3 style="color: rgb(250, 55, 55);">Введите количество коробок</h3>'
                )
        elif self.input1.text():
            # changing text of label
            self.error_label.setText(
                '<h3 style="color: rgb(250, 55, 55);">Выберете юр. лицо</h3>'
                )
        elif (not self.radiobtn1.isChecked() or not self.radiobtn2.isChecked()) and not self.input1.text():
            # changing text of label
            self.error_label.setText(
                '<h3 style="color: rgb(250, 55, 55);">Введите количество коробок и выберете юр. лицо</h3>'
                )

    def choseFolderToSave(self):
        """Выбирает папку для сохранени итогового файла"""
        layout = QGridLayout()
        self.setLayout(layout)
        self.dir_name_edit = QLineEdit()
        layout.addWidget(QLabel('label'), 0, 0)
        layout.addWidget(self.dir_name_edit, 1, 1)
        dir_name = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для сохранения файла",
            self.settings.value('Lastfile'),
        )
        if dir_name:
            path = Path(dir_name)
            self.dir_name_edit.setText(str(path))
        return (str(dir_name) + '/')