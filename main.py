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
import re

sr = dnn_superres.DnnSuperResImpl_create()
sr.readModel("./src/FSRCNN_x2.pb")
sr.setModel("fsrcnn", 2)


def get_page_text(current_page, doc):
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
        return page_text
    else:
        page = doc.load_page(current_page.number)
        page_text = page.get_text("text")
        return page_text


def doctype1(doc):
    segmenter = Segmenter()
    morph_vocab = MorphVocab()

    emb = NewsEmbedding()
    morph_tagger = NewsMorphTagger(emb)
    syntax_parser = NewsSyntaxParser(emb)
    ner_tagger = NewsNERTagger(emb)

    wb = load_workbook('./src/blank1.xlsx')
    ws = wb.active
    cnt = 0
    for current_page in doc:
        ws.cell(row=4 + cnt, column=1, value=cnt + 1)
        page_text = get_page_text(current_page, doc)
        # print(page_text)
        # print(cnt + 1)
        date1 = re.search('от\s\d{2}(\s)?.(\s)?\d{2}(\s)?.(\s)?\d{4}', page_text)
        if date1:
            date1 = re.search('\d{2}(\s)?.(\s)?\d{2}(\s)?.(\s)?\d{4}', date1[0])
            ws.cell(row=4 + cnt, column=2, value=date1[0])
            # print(date1[0])
        code1 = re.search('86010(\s)?/(\s)?\d+(\s)?/(\s)?\d+', page_text)
        if code1:
            ws.cell(row=4 + cnt, column=3, value=code1[0])
            # print(code1[0])
        name = re.search('должника:[0-9a-zA-Zа-яА-ЯёЁ\s"\\\'.-]+,', page_text)
        if name:
            name = re.search(':[0-9a-zA-Zа-яА-ЯёЁ\s"\\\'.-]+', name[0])
            name = re.search('[0-9a-zA-Zа-яА-ЯёЁ"\\\'.-][0-9a-zA-Zа-яА-ЯёЁ\s"\\\'.-]+', name[0])
            name = name[0].replace("\n", " ").replace("- ", "")
            text = 'Почему у нас сегодня на работе нету ' + name
            d = Doc(text)
            d.segment(segmenter)
            d.tag_morph(morph_tagger)
            d.parse_syntax(syntax_parser)
            d.tag_ner(ner_tagger)
            name = ''
            for span in d.spans:
                span.normalize(morph_vocab)
                name += span.normal + ' '
            name = name.title()
            ws.cell(row=4 + cnt, column=7, value=name)
            # print(name)
        date2 = re.search('\d{2}(\s)?.(\s)?\d{2}(\s)?.(\s)?\d{4}\sго(-\n)?да\sро(-\n)?ж(-\n)?де(-\n)?ния', page_text)
        if date2:
            date2 = re.search('\d{2}(\s)?.(\s)?\d{2}(\s)?.(\s)?\d{4}', date2[0])
            ws.cell(row=4 + cnt, column=6, value=date2[0])
            # print(date2[0])
        inn = re.search('ИНН\s\d+', page_text)
        if inn:
            inn = re.search('\d+', inn[0])
            ws.cell(row=4 + cnt, column=4, value=int(inn[0]))
            # print(inn[0])
        code2 = re.search('\d+(\s)?/(\s)?\d+(\s)?/(\s)?\d+(\s)?-(\s)?ИП', page_text)
        if code2:
            ws.cell(row=4 + cnt, column=5, value=code2[0])
            # print(code2[0])
        cnt += 1

    print("\nВыберите путь и введите название файла")
    output_path = ""
    tkinter.Tk().withdraw()
    while output_path == "":
        output_path = filedialog.asksaveasfilename(title="Назовите файл формата XLSX",
                                                   filetypes=[("XLSX files", "*.xlsx")])

    if output_path[-5:] == '.xlsx':
        wb.save(output_path)
    else:
        wb.save(output_path + ".xlsx")


pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
print("Выберите путь до PDF файла")
pdf_document = ""
tkinter.Tk().withdraw()
pdf_document = filedialog.askopenfilename(title="Выберите файл формата PDF",
                                          filetypes=[("PDF files", "*.pdf")])
if pdf_document == "":
    exit()
doc = fitz.open(pdf_document)
print("Исходный документ: ", doc)
print("\nКоличество страниц: %i\n\n------------------\n" % doc.page_count)
print(doc.metadata)
print("\n------------------\n\nВыберите тип документа:\n")
print("1 - Запрос от судебных приставов\n\n------------------\n")
print("Номер документа: ", end='')
doctype = input()
print('')
if doctype == '1':
    doctype1(doc)
