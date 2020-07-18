#!/usr/bin/env python
# -*- coding: utf-8 -*-
from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.pdf import ContentStream
from PyPDF2.generic import NumberObject, TextStringObject, NameObject
from PyPDF2.utils import b_

from tkinter import Tk, Label, Button, StringVar
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.constants import N,S,W,E, LEFT, TOP, RIGHT, BOTTOM
import sys, os
import tkinter.font as font



def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class PdfEnhancedFileWriter(PdfFileWriter):

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
            seq_graphics   = False
            last_font_size = 0

            for operands, operator in content.operations:

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

                # lower q and upper Q signaling graphic sequence start & stop
                if operator == b_('q'):
                    seq_graphics = True
                if operator == b_('Q'):
                    seq_graphics = False


                # changing all (text) coloring to black (removing all text highlighting that's using a text color)
                # removing (text) color completely results in wierd coloring (blue text turns yellow, probably because of another color layers)
                if seq_graphics:

                    # Blacken RGB colors
                    if operator in [b_('rg'), b_('RG')]:
                        operands = [NumberObject(0), NumberObject(0), NumberObject(0)]

                    # Blacken CMYK colors
                    if operator in [b_('k'), b_('K')]:
                        operands = [NumberObject(0), NumberObject(0), NumberObject(0), NumberObject(1)]

                    # Blacken Gray Scale colors
                    if operator in [b_('g'), b_('G')]:
                        operands = [NumberObject(0)]


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


filename1 = StringVar()
src_pdf   = StringVar()

def load_pdf(filename):
    f = open(filename,'rb')
    return PdfFileReader(f)

def load1():
    f = askopenfilename(filetypes=(('PDF File', '*.pdf'), ('All Files','*.*')))
    filename1.set(f.split('/')[-1])
    src_pdf = f
    print(f)
    print(src_pdf)
    pdf1 = load_pdf(f)
    pdf_list.append(pdf1)
    pdf_list.append(f)

def add_to_writer(pdfsrc, writer):
    [writer.addPage(pdfsrc.getPage(i)) for i in range(pdfsrc.getNumPages())]
    writer.removeWordStyle()

def remove_images():
    print("remove rectangles")
    writer          = PdfEnhancedFileWriter()
    output_filename = asksaveasfilename(filetypes = (('PDF File', '*.pdf'), ('All Files','*.*')))
    outputfile      = open(output_filename + ".pdf",'wb')

    add_to_writer(pdf_list[0], writer)

    #pdf_src = PdfFileReader(inputStream)

    writer.write(outputfile)
    outputfile.close()
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




