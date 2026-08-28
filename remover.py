#!/usr/bin/env python
# -*- coding: utf-8 -*-
# pypdf -  the modern, actively maintained Python library for PDF manipulation
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NumberObject, TextStringObject, NameObject, ContentStream

# tkinter - used for building the GUI application
from tkinter import Tk, Label, Button, StringVar, BooleanVar, Checkbutton
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter.constants import N, S, W, E, RIGHT
import sys, os
import threading

# reportlab - a library for generating PDFs programmatically
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
import io

# PyMuPDF (fitz) - used for rendering exact cropped images of the highlights
try:
    import fitz
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class PdfEnhancedFileWriter(PdfWriter):

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

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.removed_answers = []

    def _getOperatorType(self, operator):
        operator_types = {
            (b"Tj"): "text",
            (b"'"): "text",
            (b'"'): "text",
            (b"TJ"): "text",
            (b"rg"): "rgb",
            (b"RG"): "rgb",
            (b"k"): "cmyk",
            (b"K"): "cmyk",
            (b"g"): "grayscale",
            (b"G"): "grayscale",
            (b"re"): "rectangle",
            (b"l"): "line",
            (b"m"): "line",
            (b"S"): "line",
        }

        if operator in operator_types:
            return operator_types[operator]

        return None

    def _getColorTargetOperationType(self, color_index, operations):
        for i in range(color_index + 1, len(operations)):
            operator = operations[i][1]
            operator_type = self._getOperatorType(operator)

            if operator_type in ('text', 'rectangle', 'line'):
                return operator_type

        return False

    def getMinimumRectangleWidth(self, fontSize, minimumNumberOfLetters=1.5):
        return fontSize * minimumNumberOfLetters

    def removeWordStyle(self, ignoreByteStringObject=False):
        for page_number, page in enumerate(self.pages, start=1):
            content = page["/Contents"].get_object()

            if not isinstance(content, ContentStream):
                content = ContentStream(content, page)

            _operations = []
            last_font_size = 0
            operations = content.operations

            for operator_index, (operands, operator) in enumerate(operations):

                if operator == (b'Tf') and operands[0][:2] == '/F':
                    last_font_size = operands[1].as_numeric()

                if operator == (b'Tj'):
                    text = operands[0]
                    if ignoreByteStringObject:
                        if not isinstance(text, TextStringObject):
                            operands[0] = TextStringObject()
                elif operator == (b"'"):
                    text = operands[0]
                    if ignoreByteStringObject:
                        if not isinstance(text, TextStringObject):
                            operands[0] = TextStringObject()
                elif operator == (b'"'):
                    text = operands[2]
                    if ignoreByteStringObject:
                        if not isinstance(text, TextStringObject):
                            operands[2] = TextStringObject()
                elif operator == (b"TJ"):
                    for i in range(len(operands[0])):
                        if ignoreByteStringObject:
                            if not isinstance(operands[0][i], TextStringObject):
                                operands[0][i] = TextStringObject()

                operator_type = self._getOperatorType(operator)

                if operator_type in ('rgb', 'cmyk'):
                    color_target_operation_type = self._getColorTargetOperationType(operator_index, content.operations)
                    new_color = None

                    if color_target_operation_type == 'text':
                        new_color = 'black'
                    elif color_target_operation_type == 'rectangle':
                        new_color = 'white'

                    if new_color:
                        operands = self.colors_operands[operator_type][new_color]

                if operator == (b're'):
                    rectangle_width = operands[-2].as_numeric()
                    rectangle_height = operands[-1].as_numeric()

                    minWidth = self.getMinimumRectangleWidth(last_font_size, 1)
                    maxHeight = last_font_size + 6
                    minHeight = 1.5

                    if rectangle_width > minWidth and rectangle_height > minHeight and rectangle_height <= maxHeight:
                        rect_x = operands[0].as_numeric()
                        rect_y = operands[1].as_numeric()
                        self.removed_answers.append({
                            "order": len(self.removed_answers) + 1,
                            "page": page_number,
                            "pdf_rect": (rect_x, rect_y, rect_x + rectangle_width, rect_y + rectangle_height),
                        })
                        continue

                _operations.append((operands, operator))

            content.operations = _operations
            content_ref = self._add_object(content)
            page[NameObject('/Contents')] = content_ref


CROP_ZOOM = 3.0
CROP_PADDING = 2.0


def merge_highlight_rectangles(entries, threshold_x=40, threshold_y=15):
    if not entries:
        return []
    
    pages = {}
    for entry in entries:
        p = entry['page']
        if p not in pages:
            pages[p] = []
        pages[p].append(entry['pdf_rect'])
    
    merged_entries = []
    order = 1
    
    for p in sorted(pages.keys()):
        rects = pages[p]
        
        changed = True
        while changed:
            changed = False
            new_rects = []
            while rects:
                r1 = rects.pop(0)
                merged = False
                for i, r2 in enumerate(rects):
                    x0_1, y0_1, x1_1, y1_1 = r1
                    x0_2, y0_2, x1_2, y1_2 = r2
                    
                    gap_x = max(0, max(x0_1, x0_2) - min(x1_1, x1_2))
                    gap_y = max(0, max(y0_1, y0_2) - min(y1_1, y1_2))
                    
                    overlap_x = max(0, min(x1_1, x1_2) - max(x0_1, x0_2))
                    overlap_y = max(0, min(y1_1, y1_2) - max(y0_1, y0_2))
                    
                    h1, h2 = y1_1 - y0_1, y1_2 - y0_2
                    
                    same_line = overlap_y > 0.3 * min(h1, h2) and gap_x < threshold_x
                    adjacent_lines = gap_y < threshold_y and overlap_x > 0
                    
                    if same_line or adjacent_lines:
                        r1 = (min(x0_1, x0_2), min(y0_1, y0_2), max(x1_1, x1_2), max(y1_1, y1_2))
                        rects.pop(i)
                        rects.insert(0, r1) 
                        merged = True
                        changed = True
                        break
                if not merged:
                    new_rects.append(r1)
            rects = new_rects
            
        rects.sort(key=lambda r: (-round(r[1] / 10) * 10, -r[0]))
        
        for r in rects:
            merged_entries.append({
                "order": order,
                "page": p,
                "pdf_rect": r
            })
            order += 1
            
    return merged_entries


def crop_answer_images(source_pdf_path, entries):
    images = []
    if not entries:
        return images

    doc = fitz.open(source_pdf_path)
    try:
        for entry in entries:
            page = doc[entry["page"] - 1]
            page_height = page.rect.height

            x0, y0, x1, y1 = entry["pdf_rect"]
            crop = fitz.Rect(
                x0 - CROP_PADDING,
                page_height - y1 - CROP_PADDING,
                x1 + CROP_PADDING,
                page_height - y0 + CROP_PADDING,
            ) & page.rect

            if crop.is_empty:
                continue

            matrix = fitz.Matrix(CROP_ZOOM, CROP_ZOOM)
            pixmap = page.get_pixmap(matrix=matrix, clip=crop)

            images.append({
                "order": entry["order"],
                "page": entry["page"],
                "png_bytes": pixmap.tobytes("png"),
                "width_pts": pixmap.width / CROP_ZOOM,
                "height_pts": pixmap.height / CROP_ZOOM,
            })
    finally:
        doc.close()

    return images


def build_answers_summary_pdf(images, title="Highlighted Answers Summary"):
    """
    Builds the summary PDF directly into memory (BytesIO) 
    to avoid Windows file locking and disk I/O issues.
    """
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=A4)
    width, height = A4
    margin = 50
    max_img_width = width - 2 * margin
    y = height - margin

    c.setFont("Helvetica", 14)
    c.drawString(margin, y, title)
    y -= 30

    c.setFont("Helvetica", 10)
    for img in images:
        img_w, img_h = img["width_pts"], img["height_pts"]
        if img_w > max_img_width:
            scale = max_img_width / img_w
            img_w *= scale
            img_h *= scale

        needed_height = img_h + 20
        if y - needed_height < margin:
            c.showPage()
            c.setFont("Helvetica", 10)
            y = height - margin

        label = "Answer #%d - Page %d" % (img["order"], img["page"])
        c.drawString(margin, y, label)
        y -= 14

        c.drawImage(ImageReader(io.BytesIO(img["png_bytes"])), margin, y - img_h,
                    width=img_w, height=img_h)
        y -= img_h + 16

    c.save()
    packet.seek(0)
    return packet


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


def load_pdf(filename):
    f = open(filename, 'rb')
    return PdfReader(f)


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


def add_to_writer(pdfsrc, writer):
    for page in pdfsrc.pages:
        writer.add_page(page)
    writer.removeWordStyle()


def process_files_logic(output_saving_dir):
    # שומרים רפרנסים לאובייקטי הזיכרון כדי ש-pypdf תוכל לקרוא אותם בסוף
    active_streams = [] 
    
    try:
        for i, file in enumerate(pdf_list):
            writer = PdfEnhancedFileWriter()
            head, tail = os.path.split(filePaths[i])
            file_path = os.path.join(output_saving_dir, "SCRAPED_" + tail)

            status_var.set(f"מעבד קובץ {i + 1} מתוך {len(pdf_list)}...")

            add_to_writer(file, writer)

            if extract_answers_var.get() and HAS_FITZ and writer.removed_answers:
                status_var.set(f"מייצר עמוד תשובות לקובץ {i + 1}...")
                try:
                    merged_answers = merge_highlight_rectangles(writer.removed_answers)
                    images = crop_answer_images(filePaths[i], merged_answers)
                    
                    if images:
                        # ייצור ה-PDF לתוך הזיכרון הווירטואלי
                        summary_stream = build_answers_summary_pdf(images)
                        active_streams.append(summary_stream) 
                        summary_reader = PdfReader(summary_stream)
                        
                        for p in summary_reader.pages:
                            writer.add_page(p)
                except Exception as e:
                    print("Could not build the answers summary page: %s" % e)

            with open(file_path, 'wb') as outputfile:
                writer.write(outputfile)
                
        status_var.set("התהליך הסתיים בהצלחה!")
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

Label(root, textvariable=status_var, font='Helvetica 10 bold', fg="blue").grid(row=3, column=0, columnspan=3, pady=(5,0))

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

root.mainloop()
