import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd
import os
from datetime import datetime

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

st.set_page_config(page_title="Invoice Coordinate Extraction", layout="wide")
st.title("📄 Invoice Coordinate-Based Extraction (Bulk Supported)")

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


def highlight_pdf(input_path, boxes, output_path):
    doc = fitz.open(input_path)
    page = doc[0]
    for b in boxes:
        rect = fitz.Rect(b["x0"], b["top"], b["x1"], b["bottom"])
        page.add_highlight_annot(rect)
    doc.save(output_path)
    doc.close()


# --------------------------------------------------
# File Upload
# --------------------------------------------------

uploaded_files = st.file_uploader(
    "Upload Invoice PDF(s)",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:

    all_data = []

    for uploaded_file in uploaded_files:
        st.info(f"Processing: {uploaded_file.name}")

        temp_pdf_path = os.path.join(OUTPUT_DIR, uploaded_file.name)
        with open(temp_pdf_path, "wb") as f:
            f.write(uploaded_file.read())

        with pdfplumber.open(temp_pdf_path) as pdf:
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

        # Shipment Details
        for field, y in {
            "Est. Ship Date":333,
            "Est. Weight(kg)":358,
            "Transportation":385
        }.items():
            value, boxes = extract_line(words, 143, 288, y)
            extracted[field] = value
            highlight_boxes.extend(boxes)

        # Carrier
        carrier_words = [
            w for w in words
            if 134 <= w["x0"] <= 290 and 404 <= w["top"] <= 477
        ]
        carrier_sorted = sorted(carrier_words, key=lambda w: (w["top"], w["x0"]))
        extracted["Carrier"] = " ".join(w["text"] for w in carrier_sorted)
        highlight_boxes.extend(carrier_sorted)

        # Invoice Info
        for field, y in {
            "Invoice #":333,
            "Invoice Date":358,
            "Due Date":385
        }.items():
            value, boxes = extract_line(words, 400, 555, y)
            extracted[field] = value
            highlight_boxes.extend(boxes)

        # Payment & Totals
        value, boxes = extract_line(words, 135, 288, 495)
        extracted["Payment Method"] = value
        highlight_boxes.extend(boxes)

        value, boxes = extract_line(words, 135, 288, 563)
        extracted["Shipper Name"] = value
        highlight_boxes.extend(boxes)

        value, boxes = extract_block(words, 135, 288, 580, 630)
        extracted["Shipper Signature"] = value
        highlight_boxes.extend(boxes)

        financial_fields = {
            "Subtotal": (495, 6),
            "Tax ($)": (529, 12),
            "Shipping ($)": (548, 6),
            "Total Amount": (576, 8)
        }

        for field, (y, tol) in financial_fields.items():
            value, boxes = extract_line(words, 400, 555, y_center=y, tolerance=tol)
            extracted[field] = value
            highlight_boxes.extend(boxes)

        # Save highlighted PDF
        highlight_name = f"highlighted_{uploaded_file.name}"
        highlight_path = os.path.join(OUTPUT_DIR, highlight_name)
        highlight_pdf(temp_pdf_path, highlight_boxes, highlight_path)

        extracted["Source File"] = uploaded_file.name
        all_data.append(extracted)

    # --------------------------------------------------
    # Save Combined Excel
    # --------------------------------------------------
    df = pd.DataFrame(all_data)
    excel_path = os.path.join(OUTPUT_DIR, f"invoice_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    df.to_excel(excel_path, index=False)

    # --------------------------------------------------
    # UI Output
    # --------------------------------------------------
    st.subheader("📊 Extracted Invoice Data")
    st.dataframe(df)

    with open(excel_path, "rb") as f:
        st.download_button("⬇️ Download Combined Excel", f, file_name="invoice_data.xlsx")

    st.success("✅ All invoices processed successfully!")