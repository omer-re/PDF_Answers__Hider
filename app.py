#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
app.py
------
ממשק Streamlit לכלי הסרת התשובות מקבצי PDF.
כל לוגיקת העיבוד נמצאת בקובץ pdf_cleaner.py, והקובץ הזה רק בונה סביבה ממשק ווב.
"""

import io
import os
import tempfile
import zipfile

import streamlit as st

from pdf_cleaner import HAS_FITZ, process_pdf_bytes


st.set_page_config(page_title="Answers Remover", page_icon="📄", layout="centered")

# 1. הגדרת כיוון RTL גלובלי לכל האפליקציה (CSS)
st.markdown(
    """
    <style>
    .stApp {
        direction: rtl;
        text-align: right;
    }
    /* אילוץ כיווניות לימין עבור אלמנטים פנימיים כדי למנוע קפיצות טקסט */
    p, div, span, h1, h2, h3, h4, h5, h6, label {
        direction: rtl;
        text-align: right !important;
    }
    .stButton>button, .stDownloadButton>button {
        direction: rtl;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# 2. פונקציית עזר לבידוד דו-כיווני (bidi isolation)
def bidi_isolate(text):
    """
    עוטף טקסט דינמי בתווי יוניקוד FSI ו-PDI.
    מונע מהאלגוריתם הדו-כיווני לבלבל את סדר התווים במשפטים המערבים אנגלית ועברית.
    """
    return f"\u2068{text}\u2069"


# אתחול משתני state כדי לשמור את התוצאות לאחר הריצה
if "processed_results" not in st.session_state:
    st.session_state.processed_results = []
if "processing_done" not in st.session_state:
    st.session_state.processing_done = False

st.title("📄 Answers Remover")
st.caption("הסרת הדגשות/תשובות מקבצי PDF ובניית עמוד סיכום עם התשובות שהוסרו")

if not HAS_FITZ:
    st.warning(
        "החבילה PyMuPDF (fitz) לא מותקנת - לא ניתן יהיה לייצר עמוד סיכום עם תמונות של התשובות שהוסרו. "
        "העיבוד עצמו (הסרת ההדגשות) עדיין יעבוד כרגיל."
    )

uploaded_files = st.file_uploader(
    "בחרו קובץ PDF אחד או יותר",
    type=["pdf"],
    accept_multiple_files=True,
)

# אם המשתמש מוחק את הקבצים שהעלה, נאפס את מצב העיבוד
if not uploaded_files:
    st.session_state.processed_results = []
    st.session_state.processing_done = False

extract_answers = st.checkbox(
    "לחלץ את התשובות שהודגשו לעמוד סיכום בסוף הקובץ",
    value=True,
    disabled=not HAS_FITZ,
)

process_clicked = st.button("עבד קבצים", type="primary", disabled=not uploaded_files)

# שלב 1: עיבוד הקבצים (רק כשהמשתמש לוחץ על הכפתור)
if process_clicked and uploaded_files:
    # איפוס נתונים קודמים
    st.session_state.processed_results = []
    
    progress_bar = st.progress(0)
    status_placeholder = st.empty()

    def status_callback(message):
        status_placeholder.info(message)

    total = len(uploaded_files)

    for i, uploaded_file in enumerate(uploaded_files):
        isolated_filename = bidi_isolate(uploaded_file.name)
        status_callback(f"מעבד קובץ {i + 1} מתוך {total}: {isolated_filename}")

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        try:
            output_bytes = process_pdf_bytes(
                tmp_path,
                extract_answers=extract_answers,
                status_callback=status_callback,
            )
            output_name = "SCRAPED_" + uploaded_file.name
            
            # שומרים את התוצאה ישירות ב-session_state
            st.session_state.processed_results.append((output_name, output_bytes))
        except Exception as e:
            isolated_error = bidi_isolate(str(e))
            st.error(f"שגיאה בעיבוד הקובץ {isolated_filename}: {isolated_error}")
        finally:
            os.remove(tmp_path)

        progress_bar.progress((i + 1) / total)

    status_callback("התהליך הסתיים בהצלחה!")
    st.session_state.processing_done = True


# שלב 2: הצגת כפתורי ההורדה מתוך ה-session_state 
# (כך שהם יישארו שם גם כשהסקריפט רץ מחדש בעת לחיצה עליהם)
if st.session_state.processing_done and st.session_state.processed_results:
    total_processed = len(st.session_state.processed_results)
    st.success(f"עובדו {total_processed} קבצים בהצלחה.")
    st.subheader("הורדת הקבצים")

    # הורדה בודדת לכל קובץ
    for output_name, output_bytes in st.session_state.processed_results:
        isolated_out_name = bidi_isolate(output_name)
        st.download_button(
            label=f"⬇️ הורד {isolated_out_name}",
            data=output_bytes,
            file_name=output_name,
            mime="application/pdf",
            key=f"dl_{output_name}", # הוספת קידומת ל-key כדי למנוע התנגשויות
        )

    # אם יש יותר מקובץ אחד - מציעים גם הורדת ZIP מרוכז
    if total_processed > 1:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for output_name, output_bytes in st.session_state.processed_results:
                zip_file.writestr(output_name, output_bytes)
        zip_buffer.seek(0)

        st.download_button(
            label="⬇️ הורד הכל כ-ZIP",
            data=zip_buffer.getvalue(), # getvalue מחלץ את הביטים מתוך החוצץ בצורה בטוחה
            file_name="scraped_pdfs.zip",
            mime="application/zip",
            key="dl_zip_all"
        )

st.divider()

st.caption(
    """
שימו לב, האפליקציה מסירה אובייקטים מעוצבים שיובאו מוורד,
ולכן יש סיכוי שתסיר גם טבלאות ואלמנטים עיצוביים אחרים, אם קיימים.

בסוף כל קובץ תתווסף עמוד תיעוד עם רשימת התשובות שהודגשו, לפי סדר הופעתן,
כדי שיהיה ניתן לשייך כל תשובה לשאלה המתאימה.

וזיכרו: הפתרון הטוב ביותר יהיה לשלוח מייל חביב למתרגל האחראי לאחר המבחן
ולבקש ממנו להעלות גם גרסה ללא הפתרונות למען הסמסטרים הבאים.

בהצלחה!
"""
)
