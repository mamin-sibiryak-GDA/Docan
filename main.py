import os
import pathlib
import sys
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    os.environ["PYMORPHY2_DICT_PATH"] = str(pathlib.Path(sys._MEIPASS).joinpath('pymorphy2_dicts_ru/data'))
from natasha import (
    Segmenter,
    MorphVocab,

    NewsEmbedding,
    NewsMorphTagger,
    NewsSyntaxParser,
    NewsNERTagger,

    Doc
)
import fitz
import pytesseract
import cv2
from cv2 import dnn_superres
from tqdm import tqdm
import tkinter
from tkinter import filedialog
from openpyxl import load_workbook
import numpy as np

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
print("Выберите путь до PDF файла")
pdf_document = ""
tkinter.Tk().withdraw()
pdf_document = filedialog.askopenfilename(title="Выберите файл формата PDF", filetypes=[("PDF files", "*.pdf")])
tkinter.Tk().destroy()
if pdf_document == "":
    exit()
doc = fitz.open(pdf_document)
print("Исходный документ: ", doc)
print("\nКоличество страниц: %i\n\n------------------\n\n" % doc.page_count)
print(doc.metadata)
txt = ""
sr = dnn_superres.DnnSuperResImpl_create()
sr.readModel("./src/FSRCNN_x2.pb")
sr.setModel("fsrcnn", 2)
for current_page in doc:
    if doc.load_page(current_page.number).get_text("text") == "":
        for img in tqdm(doc.get_page_images(current_page.number),
                        desc="%i страница из %i" % (current_page.number + 1, doc.page_count)):
            xref = img[0]
            image = doc.extract_image(xref)
            pix = fitz.Pixmap(doc, xref)

        image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
        image = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        image_upscale = sr.upsample(image)

        page_text = pytesseract.image_to_string(image_upscale, lang='rus')
        txt += page_text
    else:
        page = doc.load_page(current_page.number)
        page_text = page.get_text("text")
        txt += page_text

f = open("tmp.txt", mode="w+", encoding="utf-8")
f.write(txt)
f.close()

print("\n\n------------------\n\nВыберите тип документа:\n")
print("1 - Запрос от судебных приставов\n\n------------------\n\n")

segmenter = Segmenter()
morph_vocab = MorphVocab()

emb = NewsEmbedding()
morph_tagger = NewsMorphTagger(emb)
syntax_parser = NewsSyntaxParser(emb)
ner_tagger = NewsNERTagger(emb)

wb = load_workbook('./src/blank1.xlsx')
ws = wb.active

num0 = 0
num1 = txt.find("\nот ", num0)
cnt = 0
while num1 != -1:
    ws.cell(row=4 + cnt, column=1, value=cnt + 1)

    num2 = txt.find("№ 86010/", num1)
    date = txt[num1 + 4:num2]
    ws.cell(row=4 + cnt, column=2, value=date)

    num3 = txt.find("Тел", num2)
    code1 = txt[num2 + 2:num3 - 1]
    ws.cell(row=4 + cnt, column=3, value=code1)

    num4 = txt.find("должника:", num3)
    tmp = txt.find("дол-\nжника:", num3)
    if tmp < num4 and tmp != -1:
        num4 = tmp
    tmp = txt.find("должни-\nка:", num3)
    if tmp < num4 and tmp != -1:
        num4 = tmp
    num5 = txt.find(",", num4)
    name = txt[num4 + 10:num5].replace("\n", " ").replace("- ", "")
    if name[0:3] != "ООО" and name[0:3] != "ОАО" and name[0:3] != "ЗАО" and name[0:3] != "АО" and name[0:3] != "ПАО":
        name = name.title()
    text = 'Почему у нас сегодня на работе нету ' + name
    doc = Doc(text)
    doc.segment(segmenter)
    doc.tag_morph(morph_tagger)
    doc.parse_syntax(syntax_parser)
    doc.tag_ner(ner_tagger)
    name = ''
    for span in doc.spans:
        span.normalize(morph_vocab)
        name += span.normal + ' '
    ws.cell(row=4 + cnt, column=7, value=name)

    dateinn = ""
    if name[0:3] == "ООО" or name[0:3] == "ОАО" or name[0:3] == "ЗАО" or name[0:3] == "АО" or name[0:3] == "ПАО":
        num6 = txt.find("ИНН", num5)
        tmpnum = txt.find(",", num6)
        dateinn = txt[num6 + 4:tmpnum]
        ws.cell(row=4 + cnt, column=4, value=int(dateinn))
    else:
        num6 = txt.find("года рождения", num4)
        tmp = txt.find("го-\nда рождения", num4)
        if tmp < num6 and tmp != -1:
            num6 = tmp
        tmp = txt.find("года\nрождения", num4)
        if tmp < num6 and tmp != -1:
            num6 = tmp
        tmp = txt.find("года ро-\nждения", num4)
        if tmp < num6 and tmp != -1:
            num6 = tmp
        tmp = txt.find("года рожде-\nния", num4)
        if tmp < num6 and tmp != -1:
            num6 = tmp
        dateinn = txt[num6 - 12:num6].replace("\n", "").replace(" ", "")
        ws.cell(row=4 + cnt, column=6, value=dateinn)

    num7 = txt.find("В отношении указанного должника возбуждено исполнительное производство от", num6)
    num8 = txt.find("№", num7)
    num9 = txt.find(".", num8)
    code2 = txt[num8 + 1:num9].replace("\n", "")
    ws.cell(row=4 + cnt, column=5, value=code2)

    num0 = num9
    num1 = txt.find("\nот ", num0)

    cnt += 1

print("Выберите путь и введите название файла")
output_path = ""
tkinter.Tk().withdraw()
while output_path == "":
    output_path = filedialog.asksaveasfilename(title="Назовите файл формата XLSX", filetypes=[("XLSX files", "*.xlsx")])
tkinter.Tk().destroy()

if output_path[-5:] == '.xlsx':
    wb.save(output_path)
else:
    wb.save(output_path + ".xlsx")
