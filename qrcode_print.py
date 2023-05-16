from PyQt6.QtWidgets import QApplication
import qrcode
from datetime import datetime
import dropbox
from dropbox.files import WriteMode
import glob
import img2pdf
import io
from PIL import Image, ImageDraw, ImageFont
import os
import time
import openpyxl
from openpyxl import load_workbook
import pandas as pd
import PySimpleGUI as sg

from helper_func import print_barcode_to_pdf
from interface_file import InterfaceQrcode


class QrcodePrint(InterfaceQrcode):
    def __init__(self):
        super().__init__()
        from main_file import dbx_db, version

        self.dbx = dbx_db

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

    """
    def stream_dropbox_file(self, path):
        _,res=self.dbx.files_download(path)
        with closing(res) as result:
            byte_data=result.content
            return io.BytesIO(byte_data)"""

    def data_for_qrcode(self, amount_box):
        output = io.BytesIO()
        global name_seller
        global id_seller
        name_seller, id_seller, counter_file_for_box_number, path = self.radiobutton_logic()
        excel_data = pd.read_excel(counter_file_for_box_number)
        data = pd.DataFrame(
                    excel_data, columns=['Номер коробки', 'Маркер'])
        excel_data2 = openpyxl.load_workbook(counter_file_for_box_number)
        sheet = excel_data2.active

        box_number_raw = data['Номер коробки'].to_list()
        marker_raw = data['Маркер'].to_list()
        global numeric_for_box
        numeric_for_box = []

        for i in range(len(box_number_raw)):
            while amount_box != 0:
                if marker_raw[i] != 1:
                    numeric_for_box.append(box_number_raw[i])
                    sheet[f"B{i+2}"].value = 1
                    amount_box -= 1
                    break
                else:
                    break
        excel_data2.save(output)
        output.seek(0)
        self.dbx.files_upload(output.getvalue(), path, mode=WriteMode.overwrite)

    def qrcode_generate(self, amount_of_box: int, amount_of_product: int = 20):
        current_date_for_save = datetime.now().strftime("%d%m%Y%H%M%S")
        qrcode_size = [img2pdf.mm_to_pt(70), img2pdf.mm_to_pt(49.5)]
        layout_function = img2pdf.get_layout_fun(qrcode_size)

        font = ImageFont.truetype("arial.ttf", size=75)

        list_with_qrcodes_files = []
        list_with_qrcodes_files.extend(
            ['programm_data/common_sign_image.pdf' for i in range(amount_of_box)])
        

        for box_number in range(amount_of_box):
            qr = qrcode.QRCode(version=2, box_size=35)
            data_for_qrcode = f'#wbbox#0002;{id_seller};{numeric_for_box[box_number]};{amount_of_product}'
            qr.add_data(data_for_qrcode)
            qrcode_img = qr.make_image()

            final_qrcode_image = Image.new('RGB', (1980, 1400), color=('#ffffff'))
            draw_text = ImageDraw.Draw(final_qrcode_image)

            # Ширина qr-кода
            qrcode_height = round(qrcode_img.height/2)

            # Вставляем qr-код в основной фон
            final_qrcode_image.paste(qrcode_img, (410, 30))
            draw_text.text(
                (510, 2*qrcode_height-90),
                f'{name_seller}\n',
                font=font,
                fill=('#000000')
                )
            draw_text.text(
                (510, 2*qrcode_height),
                f'#wbbox#0002;{id_seller};{numeric_for_box[box_number]};{amount_of_product}',
                font=font,
                fill=('#000000')
                )
            file_name = f'cache_dir/{current_date_for_save}-{box_number}.pdf'
            final_qrcode_image.save(f'cache_dir/{current_date_for_save}-{box_number}.png')
            final_qrcode_image.close()
            list_with_qrcodes_files.append(file_name)
            pdf = img2pdf.convert(
                f'cache_dir/{current_date_for_save}-{box_number}.png',
                layout_fun=layout_function)
            with open(f'cache_dir/{current_date_for_save}-{box_number}.pdf', 'wb') as f:
                f.write(pdf)
            sg.one_line_progress_meter('Создаю QR - коды', box_number+1, amount_of_box, 'Создаю QR - коды')

        dir = 'cache_dir/'
        filelist = glob.glob(os.path.join(dir, "*.png"))
        for f in filelist:
            os.remove(f)
        return list_with_qrcodes_files

    def print_qrcode_to_pdf(self):
        """Создает pdf файл для печати"""
        self.data_for_qrcode(self.get_box_amount_text())
        pdf_filenames = self.qrcode_generate(self.get_box_amount_text(), self.amount_of_product_text())
        file_name = (f'{self.choseFolderToSave()}/box-qrcode {time.strftime("%Y-%m-%d %H-%M")}.pdf')
        print_barcode_to_pdf(pdf_filenames, file_name)

        dir = 'cache_dir/'
        filelist = glob.glob(os.path.join(dir, "*"))
        for f in filelist:
            try:
                os.remove(f)
            except Exception:
                print('')
        
        print('Готово')

    dir = 'cache_dir/'
    filelist = glob.glob(os.path.join(dir, "*"))
    for f in filelist:
        os.remove(f)


if __name__ == '__main__':
    import sys 
    app = QApplication(sys.argv)
    window = QrcodePrint()
    window.show()
    sys.exit(app.exec())
