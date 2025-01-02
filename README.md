# **Property Data Extractor**

A Python-based tool for extracting and organizing property-related information from PDF documents. This project utilizes `pdfplumber` for data extraction and `pandas` for data organization, with the final output saved in a structured Excel file. It's designed for users who need to process and analyze property data efficiently.

---

## **Features**

- **PDF Data Extraction**:
  - Extracts key property information such as notice addresses, owner names, contact details, tenant information, and lease dates.
- **Regex-Based Parsing**:
  - Uses advanced regular expressions to extract structured data from unstructured PDF text.
- **Excel Output**:
  - Saves the extracted data in an Excel file with auto-adjusted column widths for readability.
- **Customizable Field Selection**:
  - Allows easy modification of the extracted fields and their order.
- **Error Handling**:
  - Handles missing fields gracefully by assigning default values.

---

## **Technologies Used**

- **Python 3.10**: Core programming language.
- **pdfplumber**: Library for extracting text from PDF documents.
- **pandas**: Used for organizing and processing extracted data.
- **openpyxl**: Saves data to Excel and adjusts column widths.
- **Regular Expressions**: Extracts specific information from text.

---

## **Installation**

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/Ikhsan121/Extract_data_pdf_to_excel.git
   cd Extract_data_pdf_to_excel
2.Create a Virtual Environment
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```
3.**Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
4.Run the Application: 
```bash
python main.py
```
