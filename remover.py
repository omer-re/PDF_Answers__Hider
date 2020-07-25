#!/usr/bin/env python
# -*- coding: utf-8 -*-
from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.pdf import ContentStream
from PyPDF2.generic import NumberObject, TextStringObject, NameObject
from PyPDF2.utils import b_

from tkinter import Tk, Label, Button, StringVar
from tkinter.filedialog import askopenfilename, asksaveasfilename, askdirectory
from tkinter.constants import N,S,W,E, LEFT, TOP, RIGHT, BOTTOM
import sys, os
import tkinter.font as font
from reportlab.pdfgen import canvas

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class PdfEnhancedFileWriter(PdfFileWriter):

    colors_operands = {
        'rgb': {
            'black': [NumberObject(0), NumberObject(0), NumberObject(0)],
            'white': [NumberObject(1), NumberObject(1), NumberObject(1)],
        },
        'cmyk': {
            'black': [NumberObject(0), NumberObject(0), NumberObject(0), NumberObject(1)],
            'white': [NumberObject(0), NumberObject(0), NumberObject(0), NumberObject(0)],
        },
        'grayscale': {
            'black': [NumberObject(0)],
            'white': [NumberObject(1)],
        }
    }

    def _getOperatorType(self, operator):
        operator_types = {
            b_('Tj'): 'text',
            b_("'"):  'text',
            b_('"'):  'text',
            b_("TJ"): 'text',

            b_('rg'): 'rgb', # color
            b_('RG'): 'rgb', # color
            b_('k'):  'cmyk', # color
            b_('K'):  'cmyk', # color
            b_('g'):  'grayscale', # color
            b_('G'):  'grayscale', # color

            b_('re'): 'rectangle',

            b_('l'): 'line', # line
            b_('m'): 'line', # start line
            b_('S'): 'line', # stroke(paint) line
        }

        if operator in operator_types:
            return operator_types[operator]

        return None

    # get the operation type that the color affects on
    def _getColorTargetOperationType(self, color_index, operations):

        for i in range(color_index + 1, len(operations)):
            operator = operations[i][1]

            operator_type = self._getOperatorType(operator)

            if operator_type == 'text' or operator_type == 'rectangle' or operator_type == 'line':
                return operator_type

        return False


    def getMinimumRectangleWidth(self, fontSize, minimumNumberOfLetters = 1.5):
        return fontSize * minimumNumberOfLetters

    def removeWordStyle(self, ignoreByteStringObject=False):
        """
        Removes imported styles from Word - Path Constructors rectangles - from this output.

        :param bool ignoreByteStringObject: optional parameter
            to ignore ByteString Objects.
        """

        pages = self.getObject(self._pages)['/Kids']
        for j in range(len(pages)):
            page    = pages[j]
            pageRef = self.getObject(page)
            content = pageRef['/Contents'].getObject()

            if not isinstance(content, ContentStream):
                content = ContentStream(content, pageRef)

            _operations    = []
            last_font_size = 0

            for operator_index, (operands, operator) in enumerate(content.operations):

                if operator == b_('Tf') and operands[0][:2] == '/F':
                    last_font_size = operands[1].as_numeric()

                if operator == b_('Tj'):
                    text = operands[0]
                    if ignoreByteStringObject:
                        if not isinstance(text, TextStringObject):
                            operands[0] = TextStringObject()
                elif operator == b_("'"):
                    text = operands[0]
                    if ignoreByteStringObject:
                        if not isinstance(text, TextStringObject):
                            operands[0] = TextStringObject()
                elif operator == b_('"'):
                    text = operands[2]
                    if ignoreByteStringObject:
                        if not isinstance(text, TextStringObject):
                            operands[2] = TextStringObject()
                elif operator == b_("TJ"):
                    for i in range(len(operands[0])):
                        if ignoreByteStringObject:
                            if not isinstance(operands[0][i], TextStringObject):
                                operands[0][i] = TextStringObject()


                operator_type = self._getOperatorType(operator)

                # we are ignoring all grayscale colors
                # tests showed that black underlines,borders and tables are defined by grayscale and aren't using rgb/cmyk colors
                if operator_type == 'rgb' or operator_type == 'cmyk':

                    color_target_operation_type = self._getColorTargetOperationType(operator_index, content.operations)

                    new_color = None

                    # we are coloring all text in black and all rectangles in white
                    # removing all colors paints rectangles in black which gives us unwanted results
                    if color_target_operation_type == 'text':
                        new_color = 'black'
                    elif color_target_operation_type == 'rectangle':
                        new_color = 'white'

                    if new_color:
                        operands = self.colors_operands[operator_type][new_color]


                # remove styled rectangles (highlights, lines, etc.)
                # the 're' operator is a Path Construction operator, creates a rectangle()
                # presumably, that's the way word embedding all of it's graphics into a PDF when creating one
                if operator == b_('re'):

                    rectangle_width  = operands[-2].as_numeric()
                    rectangle_height = operands[-1].as_numeric()

                    minWidth  = self.getMinimumRectangleWidth(last_font_size, 1) # (length of X letters at the current size)
                    maxHeight = last_font_size + 6 # range to catch really big highlights
                    minHeight = 1.5 # so that thin lines will not be removed

                    # remove only style that:
                        # it's width are bigger than the minimum
                        # it's height is smaller than maximum and larger than minimum
                    if rectangle_width > minWidth and rectangle_height > minHeight and rectangle_height <= maxHeight:
                        continue

                _operations.append((operands, operator))

            content.operations = _operations
            pageRef.__setitem__(NameObject('/Contents'), content)

root = Tk()
root.title('Answers Hider')
#root.iconbitmap( resource_path('./icon.ico'))
pdf_list = []
filePaths = []

filename1 = StringVar()
src_pdf   = StringVar()


def createMultiPage(file_path):
    c = canvas.Canvas(file_path)

    for i in range(c.getPageNumber()):
        page_num = c.getPageNumber()
        text = "This is page %s"
        c.drawString(50, 50, text)
        # c.showPage()
    c.save()


def load_pdf(filename):
    f = open(filename,'rb')
    return PdfFileReader(f)

def load1():
    f = askopenfilename(multiple=True, filetypes=(('PDF File', '*.pdf'), ('All Files', '*.*')))
    var = root.tk.splitlist(f)
    for file in var:
        print(filePaths.append(file))
        message_var = str(len(pdf_list) + 1) + " file(s) loaded"
        filename1.set(message_var)
        # filename1.set(file.split('/')[-1])
        src_pdf = file
        print(file)
        print(src_pdf)
        pdf1 = load_pdf(file)
        pdf_list.append(pdf1)
        # pdf_list.append(file)
        print("Loaded " + file)

    print(pdf_list)

def add_to_writer(pdfsrc, writer):
    [writer.addPage(pdfsrc.getPage(i)) for i in range(pdfsrc.getNumPages())]
    writer.removeWordStyle()

def remove_images():
    writer          = PdfEnhancedFileWriter()
    # output_filename = asksaveasfilename(filetypes = (('PDF File', '*.pdf'), ('All Files','*.*')))
    output_saving_dir = askdirectory(title="Choose output folder...")
    i = 0
    for file in pdf_list:
        head, tail = os.path.split(filePaths[i])
        print(tail)
        file_path = os.path.join(output_saving_dir, "SCRAPPED_" + tail)
        outputfile = open(file_path, 'wb')
        add_to_writer(file, writer)
        writer.write(outputfile)
        outputfile.close()
        i = i + 1
        print(str(i) + " file(s) done")

    print("Job is done")
    root.quit()


##Label(root, text="Rectangles remover").grid(row=0, column=2, sticky=E)
Button(root, text="Choose file", command=load1, height=5, width=14).grid(row=1, column=0)
Label(root, textvariable=filename1, width=20).grid(row=1, column=1, sticky=(N,S,E,W))
#photo= PhotoImage(file=resource_path('./button_pic.png'))

#Button(root, text="Remove answers",image=photo, command=remove_images, width=100, height=120).grid(row=1, column=2,sticky=E)
Button(root, text="Remove answers", command=remove_images, font='Helvetica 12 bold', fg="red", height=4).grid(row=1, column=2, sticky=E)

#Label(root, text="Remove Answers^^").grid(row=2, column=2, sticky=E)
#Label(root, text="Good Luck!").grid(row=2, column=0, sticky=W)

Label(root, text='''שימו לב,\n
האפליקציה מסירה אובייקטים מעוצבים שיובאו מוורד,\n
ולכן יש סיכוי שתסיר גם טבלאות ואלמנטים עיצוביים אחרים, אם קיימים.\n
הדף לא נפתח כראוי בתוכנות מסויימות של אדובי,\n
הפתרון הפשוט לכך הוא לחצן ימני על הקובץ שנוצר,\n
לחצן ימני > פתח באמצעות > כרום, פיירפוקס, או כל תוכנה אחרת שיודעת להציג פדף.\n
\n
וזיכרו: הפתרון הטוב ביותר יהיה לשלוח מייל חביב למתרגל האחראי לאחר המבחן\nולבקש ממנו להעלות גם גרסה ללא הפתרונות למען הסמסטרים הבאים.\n
\n
בהצלחה!\n''', font='Helvetica 7', justify=RIGHT).grid(row=3, columnspan=3, sticky=E)


for child in root.winfo_children():
    child.grid_configure(padx=10, pady=10)

root.mainloop()




