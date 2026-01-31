---

# 📄 SA-Hive-PDF-Text-Extraction-Automation

A **coordinate-based PDF invoice text extraction system** built with **Streamlit**, **pdfplumber**, and **PyMuPDF**.
This tool allows users to upload **single or multiple invoices**, automatically extract structured data using **fixed X-Y coordinates**, highlight extracted text inside the PDF, and export all results into Excel.

---

## 🚀 Features

✅ Bulk invoice PDF upload
✅ Coordinate-based text extraction (high accuracy for fixed templates)
✅ Extracts structured invoice fields
✅ Multi-line block support (addresses, carrier, signature, etc.)
✅ Generates **highlighted PDFs** showing extracted text
✅ Exports **combined Excel file** for all invoices
✅ Simple Streamlit UI

---

## 🧠 Extracted Invoice Fields

### 🏢 Bill To

* Name
* Email
* Phone
* Address

### 🚚 Ship To

* Name
* Email
* Phone
* Address

### 📦 Shipment Details

* Estimated Ship Date
* Estimated Weight
* Transportation
* Carrier (multi-line)

### 🧾 Invoice Information

* Invoice Number
* Invoice Date
* Due Date

### 💳 Payment & Totals

* Payment Method
* Shipper Name
* Shipper Signature
* Subtotal
* Tax
* Shipping Cost
* Total Amount

---

## 🛠 Tech Stack

| Tool               | Purpose                                  |
| ------------------ | ---------------------------------------- |
| **Streamlit**      | Web UI                                   |
| **pdfplumber**     | Extract word-level text with coordinates |
| **PyMuPDF (fitz)** | Highlight extracted text in PDF          |
| **Pandas**         | Structured data → Excel export           |

---

## 📦 Installation

### 1️⃣ Clone the Repository

```bash
git clonehttps://github.com/callmesheebygithub/SA-Hive-PDF-Text-Extraction-Automation.git
cd SA-Hive-PDF-Text-Extraction-Automation
```

### 2️⃣ Create Virtual Environment (Recommended)

```bash
python -m venv venv
```

Activate it:

**Windows**

```bash
venv\Scripts\activate
```

**Mac/Linux**

```bash
source venv/bin/activate
```

### 3️⃣ Install Dependencies

```bash
pip install streamlit pdfplumber pymupdf pandas openpyxl
```

---

## ▶️ How to Run

```bash
streamlit run FinalApp.py
```

Your browser will open automatically.

---

## 🧑‍💻 How to Use

1. Open the app in your browser
2. Click **“Upload Invoice PDF(s)”**
3. Upload **one or multiple invoice PDFs**
4. The system will:

   * Extract structured data
   * Highlight extracted areas in each PDF
   * Display results in a table
5. Download:

   * 📊 **Excel file** with all invoice data
   * 🖍️ **Highlighted PDFs** for visual verification

---

## 📁 Output Files

All generated files are saved in the **`output/`** folder:

| File                         | Description                               |
| ---------------------------- | ----------------------------------------- |
| `invoice_data.xlsx`          | Combined extracted data from all invoices |
| `highlighted_<filename>.pdf` | PDF with highlighted extracted fields     |

---

## 🎯 When to Use This

This system works best when:

✔ Invoice layout is consistent
✔ Fields appear in fixed positions
✔ You want high-speed structured extraction
✔ OCR or AI-based parsing is not required

---

## ⚙️ Customization

To adapt this tool for a new invoice template, update the **coordinate values** inside `FinalApp.py`.

Main functions to adjust:

```python
extract_line(words, x_start, x_end, y_center)
extract_block(words, x_start, x_end, y_start, y_end)
```

---

## ⚠️ Limitations

* Works best with **digitally generated PDFs** (not scanned images)
* Coordinates are template-specific
* Currently processes **page 1 only**

---

## 🔮 Future Improvements

* Table/line-item extraction
* Multi-page invoice support
* Template auto-detection
* API version (FastAPI backend)
* OCR support for scanned invoices

---

## 👨‍💻 Author

**Muhammad Shoaib**
AI Engineer | Automation Builder | PDF Intelligence Enthusiast

---

## ⭐ Support

If this project helped you, consider giving it a ⭐ on GitHub!

---

