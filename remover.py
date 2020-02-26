#!/usr/bin/env python
# -*- coding: utf-8 -*-
from PyPDF2 import PdfFileWriter, PdfFileReader
from tkinter import *
from tkinter.filedialog import *
import PyPDF2 as pypdf
import sys, os
import tkinter.font as font


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


root=Tk()
root.title('Answers Hider')
#root.iconbitmap( resource_path('./icon.ico'))
pdf_list = []


filename1=StringVar()
src_pdf=StringVar()

def load_pdf(filename):
    f = open(filename,'rb')
    return pypdf.PdfFileReader(f)

def load1():
    f = askopenfilename(filetypes=(('PDF File', '*.pdf'), ('All Files','*.*')))
    filename1.set(f.split('/')[-1])
    src_pdf=f
    print(f)
    print(src_pdf)
    pdf1 = load_pdf(f)
    pdf_list.append(pdf1)
    pdf_list.append(f)

def add_to_writer(pdfsrc, writer):
    [writer.addPage(pdfsrc.getPage(i)) for i in range(pdfsrc.getNumPages())]
    writer.removeImages()

def remove_images():
    print("remove rectangles and images")
    writer = PdfFileWriter()

    output_filename= asksaveasfilename(filetypes=(('PDF File', '*.pdf'), ('All Files','*.*')))
    outputfile= open(output_filename+".pdf",'wb')

    add_to_writer(pdf_list[0], writer)

    #pdf_src = PdfFileReader(inputStream)

    writer.write(outputfile)
    outputfile.close()
    root.quit()

##Label(root, text="Rectangles remover").grid(row=0, column=2, sticky=E)
button1=Button(root, text="Choose file", command=load1, height = 5, width = 14).grid(row=1, column=0)
Label(root, textvariable=filename1,width=20).grid(row=1, column=1, sticky=(N,S,E,W))
#photo= PhotoImage(file=resource_path('./button_pic.png'))

#Button(root, text="Remove answers",image=photo, command=remove_images, width=100, height=120).grid(row=1, column=2,sticky=E)
button2=Button(root, text="Remove answers", command=remove_images,font='Helvetica 12 bold', fg="red", height =4).grid(row=1, column=2,sticky=E)

#Label(root, text="Remove Answers^^").grid(row=2, column=2, sticky=E)
#Label(root, text="Good Luck!").grid(row=2, column=0, sticky=W)

Label(root, text='''שימו לב,\n
האפליקציה מסירה שכבה מסויימת של אובייקטים מהדף,\n
ולכן תסיר גם תמונות או שרטוטים מסויימים, אם קיימים.\n
הדף לא נפתח כראוי בתוכנות מסויימות של אדובי,\n
הפתרון הפשוט לכך הוא לחצן ימני על הקובץ שנוצר,\n
לחצן ימני > פתח באמצעות > כרום, פיירפוקס, או כל תוכנה אחרת שיודעת להציג פדף.\n
\n
וזיכרו: הפתרון הטוב ביותר יהיה לשלוח מייל חביב למתרגל האחראי לאחר המבחן\nולבקש ממנו להעלות גם גרסה ללא הפתרונות למען הסמסטרים הבאים.\n
\n
בהצלחה!\n''', font='Helvetica 7', justify=RIGHT).grid(row=3, columnspan=3, sticky=E)


for child in root.winfo_children():
    child.grid_configure(padx=10,pady=10)

root.mainloop()




