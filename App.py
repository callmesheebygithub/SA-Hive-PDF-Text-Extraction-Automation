import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd
import os

INPUT_PDF = "SampleInvoice.pdf"
HIGHLIGHTED_PDF = "output/highlighted_invoice.pdf"
EXCEL_FILE = "output/invoice_data.xlsx"

os.makedirs("output", exist_ok=True)

st.set_page_config(page_title="Invoice Extraction", layout="wide")
st.title("📄 Invoice Coordinate Extraction & Verification")

# --------------------------------------------------
# Helper Functions
# --------------------------------------------------

def find_label(words, label_text):
    for w in words:
        if w["text"].lower() == label_text.lower():
            return w
    return None


def extract_right_value(label, words, tolerance=5):
    candidates = [
        w for w in words
        if abs(w["top"] - label["top"]) < tolerance
        and w["x0"] > label["x1"]
    ]
    text = " ".join(w["text"] for w in candidates)
    return text, candidates


def extract_below_block(label, words, height=120):
    block = [
        w for w in words
        if w["top"] > label["bottom"]
        and w["top"] < label["bottom"] + height
        and abs(w["x0"] - label["x0"]) < 100
    ]
    text = " ".join(w["text"] for w in block)
    return text, block


def extract_from_x(words, x_start, y_start, height=20):
    """
    Extract text starting from x_start at y_start and go to the right until text ends.
    height defines the vertical tolerance for this line.
    """
    block = [
        w for w in words
        if w["x0"] >= x_start and 
           w["top"] >= y_start and 
           w["top"] <= y_start + height
    ]
    text = " ".join(w["text"] for w in block)
    return text, block


def highlight_pdf(boxes):
    doc = fitz.open(INPUT_PDF)
    page = doc[0]

    for b in boxes:
        rect = fitz.Rect(b["x0"], b["top"], b["x1"], b["bottom"])
        page.add_highlight_annot(rect)

    doc.save(HIGHLIGHTED_PDF)
    doc.close()


# --------------------------------------------------
# Extraction Logic
# --------------------------------------------------

with pdfplumber.open(INPUT_PDF) as pdf:
    page = pdf.pages[0]
    words = page.extract_words()

extracted = {}
highlight_boxes = []

# -----------------------
# Label-based Extraction
# -----------------------

# Invoice Number
label_invoice = find_label(words, "Invoice")
if label_invoice:
    value, boxes = extract_right_value(label_invoice, words)
    extracted["Invoice Number"] = value
    highlight_boxes.extend(boxes)

# Invoice Date
label_date = find_label(words, "Date")
if label_date:
    value, boxes = extract_right_value(label_date, words)
    extracted["Invoice Date"] = value
    highlight_boxes.extend(boxes)

# Bill To Name
label_bill = find_label(words, "Name")
if label_bill:
    value, boxes = extract_right_value(label_bill, words)
    extracted["Bill To Name"] = value
    highlight_boxes.extend(boxes)

# Email
label_email = find_label(words, "Email")
if label_email:
    value, boxes = extract_right_value(label_email, words)
    extracted["Email"] = value
    highlight_boxes.extend(boxes)

# Phone
label_phone = find_label(words, "Phone")
if label_phone:
    value, boxes = extract_right_value(label_phone, words)
    extracted["Phone"] = value
    highlight_boxes.extend(boxes)

# Total Amount
label_total = find_label(words, "Total")
if label_total:
    value, boxes = extract_below_block(label_total, words, 40)
    extracted["Total Amount"] = value
    highlight_boxes.extend(boxes)

# -----------------------
# Coordinate-based Extraction (Dynamic Width)
# -----------------------

# Invoice Location (starting x=268, y=65, extend to right dynamically)
x_start = 268
y_start = 65
value, boxes = extract_from_x(words, x_start, y_start, height=20)
extracted["Invoice Location"] = value
highlight_boxes.extend(boxes)

# --------------------------------------------------
# Highlight PDF
# --------------------------------------------------
highlight_pdf(highlight_boxes)

# --------------------------------------------------
# Save Excel
# --------------------------------------------------
df = pd.DataFrame([extracted])
df.to_excel(EXCEL_FILE, index=False)

# --------------------------------------------------
# UI
# --------------------------------------------------

st.subheader("📊 Extracted Invoice Data")
st.dataframe(df)

col1, col2 = st.columns(2)

with col1:
    with open(HIGHLIGHTED_PDF, "rb") as f:
        st.download_button(
            "⬇️ Download Highlighted PDF",
            f,
            file_name="highlighted_invoice.pdf",
            mime="application/pdf"
        )

with col2:
    with open(EXCEL_FILE, "rb") as f:
        st.download_button(
            "⬇️ Download Excel File",
            f,
            file_name="invoice_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.success("✅ Extraction completed successfully")
