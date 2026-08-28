#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
pdf_cleaner.py
--------------
כל לוגיקת עיבוד ה-PDF (הסרת הדגשות/תשובות ובניית עמוד סיכום).
הקובץ הזה לא תלוי בשום ספריית GUI (לא tkinter ולא streamlit),
כדי שאפשר יהיה לייבא אותו הן מ-gui_local.py (tkinter) והן מאפליקציית Streamlit.
"""

# pypdf -  the modern, actively maintained Python library for PDF manipulation
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NumberObject, TextStringObject, NameObject, ContentStream

import os
import io

# reportlab - a library for generating PDFs programmatically
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

# PyMuPDF (fitz) - used for rendering exact cropped images of the highlights
try:
    import fitz
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False


CROP_ZOOM = 3.0
CROP_PADDING = 2.0


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


def load_pdf(filename):
    f = open(filename, 'rb')
    return PdfReader(f)


def add_to_writer(pdfsrc, writer):
    for page in pdfsrc.pages:
        writer.add_page(page)
    writer.removeWordStyle()


def process_pdf_bytes(source_pdf_path, pdf_reader=None, extract_answers=True, status_callback=None):
    """
    מעבד קובץ PDF בודד ומחזיר את תוכן הקובץ הסופי כ-bytes.

    source_pdf_path: נתיב לקובץ המקור בדיסק (נדרש כדי ש-fitz יוכל לקרוא
                      ולחתוך תמונות של ההדגשות, גם אם pdf_reader כבר נטען מראש).
    pdf_reader:       אובייקט PdfReader שכבר נטען (אופציונלי). אם לא סופק,
                      ייטען מחדש מתוך source_pdf_path.
    extract_answers:  האם לחלץ עמוד סיכום עם תמונות התשובות שהודגשו.
    status_callback:  פונקציה אופציונלית שמקבלת מחרוזת סטטוס (למשל לעדכון GUI).
    """
    if pdf_reader is None:
        pdf_reader = load_pdf(source_pdf_path)

    writer = PdfEnhancedFileWriter()
    add_to_writer(pdf_reader, writer)

    if extract_answers and HAS_FITZ and writer.removed_answers:
        if status_callback:
            status_callback("מייצר עמוד תשובות...")
        try:
            merged_answers = merge_highlight_rectangles(writer.removed_answers)
            images = crop_answer_images(source_pdf_path, merged_answers)

            if images:
                summary_stream = build_answers_summary_pdf(images)
                summary_reader = PdfReader(summary_stream)
                for p in summary_reader.pages:
                    writer.add_page(p)
        except Exception as e:
            print("Could not build the answers summary page: %s" % e)

    output_stream = io.BytesIO()
    writer.write(output_stream)
    output_stream.seek(0)
    return output_stream.getvalue()


def process_files(file_paths, output_saving_dir, extract_answers=True, status_callback=None):
    """
    מעבד רשימת קבצי PDF ושומר את התוצאות בתיקיית יעד עם קידומת "SCRAPED_".
    מיועד לשימוש ע"י ה-GUI הלוקאלי (tkinter). מחזיר רשימת נתיבי הפלט שנוצרו.
    """
    output_paths = []

    for i, file_path in enumerate(file_paths):
        if status_callback:
            status_callback(f"מעבד קובץ {i + 1} מתוך {len(file_paths)}...")

        data = process_pdf_bytes(file_path, extract_answers=extract_answers, status_callback=status_callback)

        head, tail = os.path.split(file_path)
        out_path = os.path.join(output_saving_dir, "SCRAPED_" + tail)

        with open(out_path, 'wb') as outputfile:
            outputfile.write(data)

        output_paths.append(out_path)

    if status_callback:
        status_callback("התהליך הסתיים בהצלחה!")

    return output_paths
