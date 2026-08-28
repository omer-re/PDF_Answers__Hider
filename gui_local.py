#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
gui_local.py
------------
ממשק המשתמש הגרפי הלוקאלי (tkinter) לכלי הסרת התשובות מקבצי PDF.
כל לוגיקת העיבוד נמצאת בקובץ pdf_cleaner.py והקובץ הזה רק קורא לה.
"""

import sys
import os
import threading

# tkinter - used for building the GUI application
from tkinter import Tk, Label, Button, StringVar, BooleanVar, Checkbutton
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter.constants import N, S, W, E, RIGHT

from pdf_cleaner import HAS_FITZ, load_pdf, process_files


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


# ---------------- GUI Setup ---------------- #
root = Tk()
root.title('Answers Remover 2.0')

pdf_list = []
filePaths = []

filename1 = StringVar()
status_var = StringVar(value="ממתין לקבצים...")

extract_answers_var = BooleanVar(value=True)
if not HAS_FITZ:
    extract_answers_var.set(False)
    print("Warning: PyMuPDF (fitz) is not installed.")


def load1():
    f = askopenfilename(multiple=True, filetypes=(('PDF File', '*.pdf'), ('All Files', '*.*')))
    var = root.tk.splitlist(f)
    for file in var:
        filePaths.append(file)
        message_var = str(len(pdf_list) + 1) + " file(s) loaded"
        filename1.set(message_var)
        pdf1 = load_pdf(file)
        pdf_list.append(pdf1)
        print("Loaded " + file)

    if filePaths:
        status_var.set("הקבצים נטענו בהצלחה. ממתין להפעלה.")


def process_files_logic(output_saving_dir):
    try:
        process_files(
            filePaths,
            output_saving_dir,
            extract_answers=extract_answers_var.get(),
            status_callback=status_var.set,
        )
        print("Job is done")

    except Exception as e:
        status_var.set(f"התרחשה שגיאה: {str(e)}")
        print(f"Error: {e}")

    finally:
        btn_load.config(state="normal")
        btn_target.config(state="normal")


def trigger_processing():
    if not pdf_list:
        status_var.set("שגיאה: לא נבחרו קבצים!")
        return

    output_saving_dir = askdirectory(title="Choose output folder...")
    if not output_saving_dir:
        return

    btn_load.config(state="disabled")
    btn_target.config(state="disabled")
    status_var.set("מתחיל תהליך...")

    threading.Thread(target=process_files_logic, args=(output_saving_dir,), daemon=True).start()


# ---------------- UI Layout ---------------- #
btn_load = Button(root, text="Choose one or more PDFs", command=load1, font='Helvetica 12 bold', fg="green", height=4, width=20)
btn_load.grid(row=1, column=0)

Label(root, textvariable=filename1, width=20).grid(row=1, column=1, sticky=(N, S, E, W))

Checkbutton(
    root,
    text='extract highlighted answers to a summary page',
    variable=extract_answers_var,
    state=("normal" if HAS_FITZ else "disabled"),
).grid(row=2, column=0, columnspan=3, sticky=E)

btn_target = Button(root, text="Target folder", command=trigger_processing, font='Helvetica 12 bold', fg="red", height=4)
btn_target.grid(row=1, column=2, sticky=E)

Label(root, textvariable=status_var, font='Helvetica 10 bold', fg="blue").grid(row=3, column=0, columnspan=3, pady=(5, 0))

Label(root, text='''שימו לב,\n
האפליקציה מסירה אובייקטים מעוצבים שיובאו מוורד,\n
ולכן יש סיכוי שתסיר גם טבלאות ואלמנטים עיצוביים אחרים, אם קיימים.\n
\n
בסוף כל קובץ תתווסף עמוד תיעוד עם רשימת התשובות שהודגשו, לפי סדר הופעתן,\n
כדי שיהיה ניתן לשייך כל תשובה לשאלה המתאימה.\n
\n
וזיכרו: הפתרון הטוב ביותר יהיה לשלוח מייל חביב למתרגל האחראי לאחר המבחן\nולבקש ממנו להעלות גם גרסה ללא הפתרונות למען הסמסטרים הבאים.\n
\n
בהצלחה!\n''', font='Helvetica 7', justify=RIGHT).grid(row=4, columnspan=3, sticky=E)


for child in root.winfo_children():
    child.grid_configure(padx=10, pady=10)

if __name__ == "__main__":
    root.mainloop()
