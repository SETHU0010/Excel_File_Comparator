# 📊 Excel File Comparator with Streamlit

This project is a **Streamlit-based web application** that allows users to compare two Excel files (`.xlsx` or `.xls`) and highlight:
- 🟨 Modified cells
- 🟩 New rows
- 🟥 Deleted rows

A downloadable comparison report is generated with color-coded highlights using **OpenPyXL**.

---

## 🚀 Features

- 🔍 Cell-by-cell and row-by-row comparison
- 🧠 Optional key column selection for accurate matching
- ✅ Summary statistics: new, deleted, and modified
- 🎨 User-selectable highlight types (new/deleted/modified)
- 📥 Downloadable Excel report with color highlights
- 📊 Preview of top 5 rows for each uploaded file

---

## 🛠️ Tech Stack

- [Streamlit](https://streamlit.io/)
- [Pandas](https://pandas.pydata.org/)
- [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/)

---

## 📦 Installation

Clone this repository and install the dependencies:

```bash
git clone https://github.com/YOUR_USERNAME/excel-comparator.git
cd excel-comparator
pip install -r requirements.txt
