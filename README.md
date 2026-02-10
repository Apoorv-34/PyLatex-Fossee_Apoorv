# 📊 Structural Beam Analysis Report Generator

## 📖 Project Overview
This project is an automated Python-based solution designed to generate professional engineering reports for **Simply Supported Beams**. By bridging the gap between data analysis and professional typesetting, the script reads loading data from Excel and produces a high-fidelity PDF report using **PyLaTeX**.

The tool is specifically designed to meet structural engineering standards, featuring vector-based diagrams that remain crisp at any zoom level.

---

## ✨ Features
* **✅ Automated Data Ingestion:** Uses `pandas` to extract loading points and magnitudes from Excel spreadsheets.
* **✅ Native LaTeX Tabulars:** Reconstructs the force table as selectable, searchable LaTeX text (no image-based tables).
* **✅ TikZ/pgfplots Diagrams:** Generates vector-based **Shear Force Diagrams (SFD)** and **Bending Moment Diagrams (BMD)** directly in the PDF.
* **✅ Professional Formatting:** Automatically generates a Title Page, Table of Contents, and structured sections for Introduction, Analysis, and Summary.
* **✅ Visual Integration:** Embeds reference beam diagrams within the analysis section using PyLaTeX Figure environments.



[Image of a Shear Force and Bending Moment Diagram for a simply supported beam]


---

## 🛠 Technical Stack
* **Core Logic:** Python 3.10+
* **Data Processing:** `pandas`, `openpyxl`
* **PDF Generation:** `PyLaTeX`
* **Graphics:** `TikZ` / `pgfplots` (LaTeX-native vector graphics)
* **LaTeX Backend:** TeX Live, MiKTeX, or MacTeX

---

## 🚀 Quick Start

### 1. Prerequisites
Ensure you have a LaTeX distribution installed on your system (like **MiKTeX** or **TeX Live**) so the script can compile the `.tex` output.

### 2. Installation
Clone the repository and install the dependencies:
```bash
git clone [https://github.com/Apoorv-34/PyLatex-Fossee_Apoorv.git](https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git)
cd YOUR_REPO_NAME
pip install -r requirements.txt
