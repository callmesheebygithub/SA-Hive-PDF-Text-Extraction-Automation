import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd
import os

OUTPUT_FOLDER = "output"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

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


def highlight_pdf(input_pdf_path, boxes, output_path):
    doc = fitz.open(input_pdf_path)
    page = doc[0]
    for b in boxes:
        rect = fitz.Rect(b["x0"], b["top"], b["x1"], b["bottom"])
        page.add_highlight_annot(rect)
    doc.save(output_path)
    doc.close()


# --------------------------------------------------
# Upload PDFs
# --------------------------------------------------
uploaded_files = st.file_uploader("Upload Invoice PDF(s)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    all_data = []
    highlighted_files = []

    for idx, uploaded_file in enumerate(uploaded_files):
        file_name = uploaded_file.name
        temp_pdf_path = os.path.join(OUTPUT_FOLDER, file_name)

        with open(temp_pdf_path, "wb") as f:
            f.write(uploaded_file.read())

        with pdfplumber.open(temp_pdf_path) as pdf:
            page = pdf.pages[0]
            words = page.extract_words()

        extracted = {"File Name": file_name}
        highlight_boxes = []

        # ---------------- Bill To ----------------
        for field, y in {"Name":166, "Email":191, "Phone":216}.items():
            value, boxes = extract_line(words, 134, 290, y)
            extracted[f"Bill To {field}"] = value
            highlight_boxes.extend(boxes)

        value, boxes = extract_block(words, 134, 290, 237, 286)
        extracted["Bill To Address"] = value
        highlight_boxes.extend(boxes)

        # ---------------- Ship To ----------------
        for field, y in {"Name":166, "Email":191, "Phone":216}.items():
            tol = 12 if field == "Name" else 6
            value, boxes = extract_line(words, 400, 555, y, tolerance=tol)
            extracted[f"Ship To {field}"] = value
            highlight_boxes.extend(boxes)

        value, boxes = extract_block(words, 400, 555, 237, 286)
        extracted["Ship To Address"] = value
        highlight_boxes.extend(boxes)

        # ---------------- Shipment Details ----------------
        for field, y in {
            "Est. Ship Date":333,
            "Est. Weight(kg)":358,
            "Transportation":385
        }.items():
            value, boxes = extract_line(words, 143, 288, y)
            extracted[field] = value
            highlight_boxes.extend(boxes)

        carrier_words = [
            w for w in words
            if 134 <= w["x0"] <= 290 and 404 <= w["top"] <= 477
        ]
        carrier_words_sorted = sorted(carrier_words, key=lambda w: (w["top"], w["x0"]))
        extracted["Carrier"] = " ".join(w["text"] for w in carrier_words_sorted)
        highlight_boxes.extend(carrier_words_sorted)

        # ---------------- Invoice Info ----------------
        for field, y in {
            "Invoice #":333,
            "Invoice Date":358,
            "Due Date":385
        }.items():
            value, boxes = extract_line(words, 400, 555, y)
            extracted[field] = value
            highlight_boxes.extend(boxes)

        # ---------------- Payment & Totals ----------------
        value, boxes = extract_line(words, 135, 288, 495)
        extracted["Payment Method"] = value
        highlight_boxes.extend(boxes)

        value, boxes = extract_line(words, 135, 288, 563)
        extracted["Shipper Name"] = value
        highlight_boxes.extend(boxes)

        value, boxes = extract_block(words, 135, 288, 580, 630)
        extracted["Shipper Signature"] = value
        highlight_boxes.extend(boxes)

        for field, (y, tol) in {
            "Subtotal": (495, 6),
            "Tax ($)": (529, 12),
            "Shipping ($)": (548, 6),
            "Total Amount": (576, 8)
        }.items():
            value, boxes = extract_line(words, 400, 555, y_center=y, tolerance=tol)
            extracted[field] = value
            highlight_boxes.extend(boxes)

        # Save highlighted PDF
        highlighted_path = os.path.join(OUTPUT_FOLDER, f"highlighted_{file_name}")
        highlight_pdf(temp_pdf_path, highlight_boxes, highlighted_path)
        highlighted_files.append(highlighted_path)

        all_data.append(extracted)

    # Save Excel
    df = pd.DataFrame(all_data)
    excel_path = os.path.join(OUTPUT_FOLDER, "invoice_data.xlsx")
    df.to_excel(excel_path, index=False)

    # UI
    st.subheader("📊 Extracted Data")
    st.dataframe(df)

    with open(excel_path, "rb") as f:
        st.download_button("⬇️ Download Excel File", f, file_name="invoice_data.xlsx", key="excel_dl")

    st.subheader("🖍️ Download Highlighted PDFs")
    for i, file_path in enumerate(highlighted_files):
        with open(file_path, "rb") as f:
            st.download_button(
                label=f"Download {os.path.basename(file_path)}",
                data=f,
                file_name=os.path.basename(file_path),
                key=f"pdf_dl_{i}"   # UNIQUE KEY FIX
            )

    st.success("✅ All invoices processed successfully!")
