"""Объединяет в один PDF файл все файлы, входящие в список list_filenames.
   Сохраняет конечный файл в папку и названием, указанное в folder_summary_file_name"""

from barcode import Code128
from barcode.writer import ImageWriter
from contextlib import closing
from datetime import datetime
import io
from PyPDF3 import PdfFileReader, PdfFileWriter
from PyPDF3.pdf import PageObject
import PySimpleGUI as sg
import glob
import os
import pandas as pd
from pdf2image import convert_from_path
import img2pdf
from PIL import Image, ImageDraw, ImageFont
from PyQt6.QtWidgets import QMessageBox


def stream_dropbox_file(path):
    from main_file import dbx_db

    dbx = dbx_db
    _,res=dbx.files_download(path)
    with closing(res) as result:
        byte_data=result.content
        return io.BytesIO(byte_data)

def print_barcode_to_pdf2(list_filenames, folder_summary_file_name):
    """
    Создает pdf файл для печати. С возможностью удаления всего кеша.
    Входящие данные:
    list_filenames - список с полными адресами и названиями файлов для объединения,
    folder_summary_file_name - полное название файла для сохранения 
    (вместе с названием папок в пути)
    """
    with open(list_filenames[0], "rb") as f:
        input1 = PdfFileReader(f, strict=False)
        page1 = input1.getPage(0)
        total_width = max([page1.mediaBox.upperRight[0]*(3)])
        total_height = max([page1.mediaBox.upperRight[1]*(6)])
        horiz_size = page1.mediaBox.upperRight[0]
        vertic_size = page1.mediaBox.upperRight[1]
        output = PdfFileWriter()
        file_name = folder_summary_file_name
        new_page = PageObject.createBlankPage(
            file_name, total_width, total_height)
        new_page.mergeTranslatedPage(page1, 0, vertic_size*(5))
        output.addPage(new_page)
        page_amount = (len(list_filenames) // 18) + 1
        pages_names = []
        for p in range(1, page_amount):
            p = PageObject.createBlankPage(
                file_name, total_width, total_height)
            output.addPage(p)
            pages_names.append(p)
        for i in range(1, len(list_filenames)):
            with open(list_filenames[i], "rb") as bb:
                m = i // 18
                n = (i // 3) - 6*m
                k = i % 3
                if i < 18:
                    new_page.mergeTranslatedPage(
                        PdfFileReader(bb,
                        strict=False).getPage(0),
                        horiz_size*(k),
                        vertic_size*(5-n))
                elif i >= 18:
                    (pages_names[m-1]).mergeTranslatedPage(
                        PdfFileReader(bb,
                        strict=False).getPage(0),
                        horiz_size*(k),
                        vertic_size*(5-n))
                sg.one_line_progress_meter(
                    'Объединяю штрихкоды в PDF файл',
                    i+1,
                    len(list_filenames),
                    'Объединяю штрихкоды в PDF файл')
                output.write(open(file_name, "wb"))
    f.close()

def print_barcode_to_pdf(list_filenames, folder_summary_file_name):
    """
    Создает pdf файл для печати. Без возможности удаления кеша.
    Входящие данные:
    list_filenames - список с полными адресами и названиями файлов для объединения,
    folder_summary_file_name - полное название файла для сохранения 
    (вместе с названием папок в пути)
    """
    with open(list_filenames[0], "rb") as f:
        input1 = PdfFileReader(f, strict=False)
        page1 = input1.getPage(0)
        total_width = max([page1.mediaBox.upperRight[0]*(3)])
        total_height = max([page1.mediaBox.upperRight[1]*(6)])
        horiz_size = page1.mediaBox.upperRight[0]
        vertic_size = page1.mediaBox.upperRight[1]
        output = PdfFileWriter()
        file_name = folder_summary_file_name
        new_page = PageObject.createBlankPage(
            file_name, total_width, total_height)
        new_page.mergeTranslatedPage(page1, 0, vertic_size*(5))
        output.addPage(new_page)
        page_amount = (len(list_filenames) // 18) + 1
        pages_names = []
        for p in range(1, page_amount):
            p = PageObject.createBlankPage(
                file_name, total_width, total_height)
            output.addPage(p)
            pages_names.append(p)
        for i in range(1, len(list_filenames)):
            pdfreader_var = (PdfFileReader(open(list_filenames[i], "rb"),
                             strict=False)).getPage(0)
            m = i // 18
            n = (i // 3) - 6*m
            k = i % 3
            
            if i < 18:
                new_page.mergeTranslatedPage(
                    pdfreader_var,
                    horiz_size*(k),
                    vertic_size*(5-n))
            elif i >= 18:
                (pages_names[m-1]).mergeTranslatedPage(
                    pdfreader_var,
                    horiz_size*(k),
                    vertic_size*(5-n))
            sg.one_line_progress_meter(
                'Объединяю штрихкоды в PDF файл',
                i+1,
                len(list_filenames),
                'Объединяю штрихкоды в PDF файл')
        output.write(open(file_name, "wb"))
        f.close()


def qrcode_print_to_file_main(full_filename_with_qrcodes, filename):
    """
    Создает QR коды в необходимом формате. Входящие файлы:
    full_filename_with_qrcodes - полный путь и название до файла с qr-кодами,
    filename - название файла с qr-кодами. Для создания промежуточной папки.
    """
    
    inputpdf = PdfFileReader(open(full_filename_with_qrcodes, "rb"))
    for i in range(len(inputpdf.pages)):
        output = PdfFileWriter()
        output.addPage(inputpdf.pages[i])
        with open(
            f"cache_dir_2/{filename} {i}.pdf", "wb"
                ) as outputStream:
            output.write(outputStream)
        image = convert_from_path(
            f"cache_dir_2/{filename} {i}.pdf",
            dpi=718.896, poppler_path=r'poppler-0.68.0\bin')
        image[0].save(f"cache_dir_2/{filename} {i}.png")
        sg.one_line_progress_meter('Создаю QR - коды',
                                   i+1,
                                   len(inputpdf.pages),
                                   'Создаю QR - коды')
    dir = 'cache_dir_2/'
    filelist = glob.glob(os.path.join(dir, "*.png"))
    filelist.sort()
    i = 0
    if not os.path.isdir(f'cache_dir_3/{filename}/'):
        os.mkdir(f'cache_dir_3/{filename}/')
    for file in filelist:
        barcode_size = [img2pdf.in_to_pt(2.759), img2pdf.in_to_pt(1.95)]
        layout_function = img2pdf.get_layout_fun(barcode_size)
        im = Image.new('RGB', (1980, 1400), color=('#ffffff'))
        image1 = Image.open(file)
        # Длина подложки
        w_width = round(im.width/2)
        # Высота подложки
        w_height = round(im.height/2)
        w = round(image1.width/2)
        h = round(image1.height/2)
        # Вставляем qr код в основной фон
        im.paste(image1, ((w_width - w), (w_height-h)))
        im.save(f'cache_dir_3/{filename}/{filename} {i}.png')
        pdf = img2pdf.convert(
            f'cache_dir_3/{filename}/{filename} {i}.png', layout_fun=layout_function)
        with open(f'cache_dir_3/{filename}/{filename} {i}.pdf', 'wb') as f:
            f.write(pdf)
        i += 1 
    pdf_filenames_qrcode = glob.glob(f'cache_dir_3/{filename}/*.pdf')
    pdf_filenames_qrcode.sort()
    filelist.clear()
    dir = 'cache_dir_2/'
    filelist = glob.glob(os.path.join(dir, "*"))
    for f in filelist:
        try:
            os.remove(f)
        except Exception:
            print('')
    return pdf_filenames_qrcode


def design_barcodes(
        names_for_print_barcodes,
        article_list_main_file,
        amount_file_location,
        main_file_location,
        barcodelist_list_main_file,
        name_list_main_file,
        sheet):
    """
    Создает дизайн штрихкода. Входящие файлы:
    names_for_print_barcodes - список всех артикулов для печати (получаем из файла),
    article_list_main_file - список всех ариткулов из главной базы (ООО или ИП),
    amount_file_location - расположение и название файла с артикулами для печати,
    main_file_location - расположение и название файла с главной базой (ООО или ИП),
    barcodelist_list_main_file - список со всеми баркодами товаров из главной базы (ООО или ИП),
    name_list_main_file - список со всеми названиями товаров из главной базы (ООО или ИП),
    sheet - лист excel с юр. данными для печати на штрихкоде (ООО или ИП)
    """

    path_checked_file = '/DATABASE/список сопоставления.xlsx'

    checked_file = stream_dropbox_file(path_checked_file)

    checked_data = pd.read_excel(checked_file)
    data = pd.DataFrame(
        checked_data, columns=['Артикул', 'Артикул ВБ'])
    list_article_checked = data['Артикул'].to_list()
    list_article_checked = [item.capitalize() for item in list_article_checked]
    list_article_wb_checked = data['Артикул ВБ'].to_list()
    list_article_wb_checked = [item.capitalize() for item in list_article_wb_checked]
    # Подключаем шрифты чтобы писать текст
    font = ImageFont.truetype("arial.ttf", size=50)
    font2 = ImageFont.truetype("arial.ttf", size=60)
    font3 = ImageFont.truetype("arial.ttf", size=120)

    # Задаем размер штрихкода
    barcode_size = [img2pdf.in_to_pt(2.759), img2pdf.in_to_pt(1.95)]
    layout_function = img2pdf.get_layout_fun(barcode_size)

    current_date = datetime.now().strftime("%d.%m.%Y")
    eac_file = 'programm_data/eac.png'

    # Список для проверки первых букв штрихкода
    list_first_letter = []
    print('до фильтра', names_for_print_barcodes)
    # Проверка на наличие артикула в базе данных
    for article in names_for_print_barcodes:
        if article not in article_list_main_file:
            if article in list_article_checked:
                print('правильный путь')
                for i in range(len(list_article_checked)):
                    if article == list_article_checked[i]:
                        names_for_print_barcodes.remove(article)
                        names_for_print_barcodes.append(list_article_wb_checked[i])
                        break
            else:
                QMessageBox.critical(None,
                    "Error",
                    f'Артикула {article} нет в основном файле и файле сверки',
                    QMessageBox.StandardButton.Cancel)
        sg.one_line_progress_meter('Проверяю соответствие',
                                   len(article)+1, len(list_article_checked),
                                   'Проверяю соответствие')
    print('после фильтра', names_for_print_barcodes)
    for article in names_for_print_barcodes:
        print(article)
        if article not in article_list_main_file:
                QMessageBox.critical(None,
                    "Error",
                    f'Артикула {article} нет в основном файле и файле сверки',
                    QMessageBox.StandardButton.Cancel)
        else:
            list_first_letter.append(article[0])

    # Создание самого штрихкода
    for i in range(len(names_for_print_barcodes)):
        for j in range(len(article_list_main_file)):
            if names_for_print_barcodes[i] == article_list_main_file[j]:
                render_options = {
                    "module_width": 1,
                    "module_height": 35,
                    "font_size": 20,
                    "text_distance": 8,
                    "quiet_zone": 8
                }
                barcode = Code128(
                    f'{str(barcodelist_list_main_file[j])[:13]}',
                    writer=ImageWriter()
                    ).render(render_options)
                im = Image.new('RGB', (1980, 1400), color=('#000000'))
                image1 = Image.open(f'{eac_file}')
                draw_text = ImageDraw.Draw(im)
                # Длина подложки
                w_width = round(im.width/2)
                # Длина штрихкода
                w = round(barcode.width/2)
                # Расположение штрихкода по центру
                position = w_width - w
                # Вставляем штрихкод в основной фон
                im.paste(barcode, ((w_width - w), 250))
                # Вставляем EAC в основной фон
                im.paste(image1, (1505, 1028))
                draw_text.text(
                    (position, 150),
                    f'{name_list_main_file[j]}\n',
                    font=font2,
                    fill=('#ffffff'), stroke_width=1
                    )
                draw_text.text(
                    (position, barcode.height+270),
                    f'{article_list_main_file[j]}\n',
                    font=font3,
                    fill=('#ffffff'), stroke_width=1
                    )
                draw_text.text(
                    (position, barcode.height+410),
                    f'{sheet["A1"].value}\n'
                    f'{sheet["A2"].value}{current_date}\n'
                    f'{sheet["A3"].value} {sheet["A4"].value}\n'
                    f'{sheet["A5"].value}\n'
                    f'{sheet["A6"].value}',
                    font=font,
                    fill=('#ffffff')
                    )
                im.save(f'cache_dir/{article_list_main_file[j]}.png')
                im.close()
                pdf = img2pdf.convert(
                    f'cache_dir/{article_list_main_file[j]}.png',
                    layout_fun=layout_function)
                with open(f'cache_dir/{article_list_main_file[j]}.pdf', 'wb') as f:
                    f.write(pdf)
                sg.one_line_progress_meter('Создаю штрихкоды', i+1, len(names_for_print_barcodes), 'Создаю штрихкоды')
