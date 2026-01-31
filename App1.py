import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd
import os

INPUT_PDF = "SampleInvoice.pdf"
HIGHLIGHTED_PDF = "output/highlighted_invoice.pdf"
EXCEL_FILE = "output/invoice_data.xlsx"

os.makedirs("output", exist_ok=True)

st.set_page_config(page_title="Invoice Coordinate Extraction", layout="wide")
st.title("📄 Invoice Coordinate-Based Extraction")

# --------------------------------------------------
# Helper Functions
# --------------------------------------------------

def extract_line(words, x_start, x_end, y_center, tolerance=6):
    block = [
        w for w in words
        if w["x0"] >= x_start and w["x0"] <= x_end and
           (y_center - tolerance) <= w["top"] <= (y_center + tolerance)
    ]
    text = " ".join(w["text"] for w in sorted(block, key=lambda w: w["x0"]))
    return text, block


def extract_block(words, x_start, x_end, y_start, y_end):
    block = [
        w for w in words
        if w["x0"] >= x_start and w["x0"] <= x_end and
           w["top"] >= y_start and w["bottom"] <= y_end
    ]
    text = " ".join(w["text"] for w in sorted(block, key=lambda w: (w["top"], w["x0"])))
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
# Bill To
# -----------------------
x_start_bill, x_end_bill = 134, 290
for field, y in {"Name":166, "Email":191, "Phone":216}.items():
    value, boxes = extract_line(words, x_start_bill, x_end_bill, y)
    extracted[f"Bill To {field}"] = value
    highlight_boxes.extend(boxes)

value, boxes = extract_block(words, x_start_bill, x_end_bill, 237, 286)
extracted["Bill To Address"] = value
highlight_boxes.extend(boxes)

# -----------------------
# Ship To
# -----------------------
x_start_ship, x_end_ship = 400, 555
for field, y in {"Name":166, "Email":191, "Phone":216}.items():
    tol = 12 if field == "Name" else 6
    value, boxes = extract_line(words, x_start_ship, x_end_ship, y, tolerance=tol)
    extracted[f"Ship To {field}"] = value
    highlight_boxes.extend(boxes)

value, boxes = extract_block(words, x_start_ship, x_end_ship, 237, 286)
extracted["Ship To Address"] = value
highlight_boxes.extend(boxes)

# -----------------------
# Shipment Details (Left Mid)
# -----------------------
x_left, x_left_end = 143, 288
for field, y in {
    "Est. Ship Date":333,
    "Est. Weight(kg)":358,
    "Transportation":385
}.items():
    value, boxes = extract_line(words, x_left, x_left_end, y)
    extracted[field] = value
    highlight_boxes.extend(boxes)

# Carrier
carrier_words = [
    w for w in words
    if w["x0"] >= 143 and w["x0"] <= 290 and 404 <= w["top"] <= 477
]
carrier_words_sorted = sorted(carrier_words, key=lambda w: (w["top"], w["x0"]))
extracted["Carrier"] = " ".join(w["text"] for w in carrier_words_sorted)
highlight_boxes.extend(carrier_words_sorted)

# -----------------------
# Invoice Info (Right Mid)
# -----------------------
x_right, x_right_end = 400, 555
for field, y in {
    "Invoice #":333,
    "Invoice Date":358,
    "Due Date":385
}.items():
    value, boxes = extract_line(words, x_right, x_right_end, y)
    extracted[field] = value
    highlight_boxes.extend(boxes)

# -----------------------
# NEW SECTION — Payment & Totals
# -----------------------

# Payment Method
value, boxes = extract_line(words, 135, 288, 495)
extracted["Payment Method"] = value
highlight_boxes.extend(boxes)

# Shipper Name
value, boxes = extract_line(words, 135, 288, 563)
extracted["Shipper Name"] = value
highlight_boxes.extend(boxes)

# Shipper Signature (multi-line block)
value, boxes = extract_block(words, 135, 288, 580, 630)
extracted["Shipper Signature"] = value
highlight_boxes.extend(boxes)

# Financial Totals (Right Bottom)
financial_fields = {
    "Subtotal": (495, 6),
    "Tax ($)": (529, 12),          # ⬅ Increased tolerance
    "Shipping ($)": (548, 6),
    "Total Amount": (576, 8)       # Slightly larger tolerance
}

for field, (y, tol) in financial_fields.items():
    value, boxes = extract_line(words, 400, 555, y_center=y, tolerance=tol)
    extracted[field] = value
    highlight_boxes.extend(boxes)


# --------------------------------------------------
# Highlight + Save
# --------------------------------------------------
highlight_pdf(highlight_boxes)

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
        st.download_button("⬇️ Download Highlighted PDF", f, file_name="highlighted_invoice.pdf")
with col2:
    with open(EXCEL_FILE, "rb") as f:
        st.download_button("⬇️ Download Excel File", f, file_name="invoice_data.xlsx")

st.success("✅ Extraction completed successfully")
# Financial Totals (Right Bottom)