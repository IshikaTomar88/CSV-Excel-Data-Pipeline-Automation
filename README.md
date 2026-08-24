# ⚡ Zero-Touch Data Pipeline & Executive Reporter

An enterprise-grade, high-performance ETL (Extract, Transform, Load) and data cleaning application built with **Python**, **Pandas**, and **Streamlit**. Designed to eliminate manual data wrangling for operations managers and e-commerce founders by instantly transforming messy, unformatted raw exports into executive-ready dashboards.

---

## 🚀 Key Features

* **Vectorized Processing Engine:** Leverages Pandas and NumPy to process 100k+ rows of raw data in under 2 seconds with absolute mathematical precision.
* **Universal File Ingestion:** Transparently handles `.csv`, `.xlsx`, `.xls`, and `.json` raw file formats.
* **Automated Data Hygiene:** 
  * Normalizes and fuzzy-matches disparate column names (e.g., `cust name`, `Amt`, `QTY` $\rightarrow$ canonical schema).
  * Cleans currency strings, strips whitespace, corrects text capitalization, and parses dates automatically.
  * Handles missing values via configurable strategies (`drop_row`, `fill_zero`, `fill_mean`, `fill_unknown`).
  * Automatically detects and removes duplicate records.
  * Flags numerical statistical outliers using Z-score thresholds.
* **🔒 Enterprise Privacy Guarantee:** All files are processed securely in isolated in-memory runtime sessions. Zero persistent storage or data logging ensures absolute client confidentiality.
* **🛡️ Smart Trial Metering & Admin Control:** Built-in usage gate that restricts public link access after free runs, sending an instant email alert and offering an interactive admin dashboard to grant or revoke day-based client licenses.

---

## 🛠️ Tech Stack & Dependencies

* **Python 3.10+**
* **Streamlit** (Interactive Web UI)
* **Pandas & NumPy** (High-performance vector data cleaning)
* **OpenPyXL & XlsxWriter** (Advanced multi-sheet Excel generation and reporting)

---

## 📦 Project Directory Structure

```text
├── app.py                 # Main Streamlit application & ETL pipeline
├── requirements.txt       # Project dependencies
└── README.md              # Project documentation