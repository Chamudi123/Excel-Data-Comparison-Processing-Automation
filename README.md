# 🔍 Excel VBA Data Comparison and Processing Automation

This project automates the comparison and processing of PreAlert and ImportManifest data using Excel VBA. Built to streamline logistics data validation and minimize manual processing, the tool performs intelligent data transformation, fuzzy matching using Levenshtein distance, and highlights discrepancies with color-coded results.

---

## 📌 Project Overview

This tool was developed to automate and enhance the comparison of flight-related logistics data between two large datasets: **PreAlert** and **ImportManifest**. The macro-driven solution reduces manual workload and increases accuracy by automating:

- Data preparation
- Weight conversion
- String matching (e.g., names, cities)
- Currency extraction
- Discrepancy detection and reporting

---

## 🧠 Key Features

- **Levenshtein Distance Matching**: Uses fuzzy logic to detect similar but non-identical strings (e.g., customer names).
- **Weight & Value Standardization**: Automatically converts weights to a standard unit (e.g., lbs to kgs) and extracts USD values.
- **Structured Discrepancy Reporting**: Results are output in a color-coded format, highlighting mismatches across multiple fields.
- **Step-by-Step Automation**:
  - First, run `ProcessData` to prepare and clean datasets.
  - Then, run `Comparison` to identify and output discrepancies.

---

## ⚙️ Requirements

- Microsoft Excel (with macro support enabled)
- PreAlert data file with expected column structure (see the sample data file)
- ImportManifest file with expected column structure (see the sample data file)

---

## 🖥 Sample Workflow

1. Paste PreAlert data into the **PreAlert** sheet.
2. Paste ImportManifest data into the **ImportManifest** sheet.
3. Run the `ProcessData` macro to clean and standardize.
4. Run the `Comparison` macro to detect and output discrepancies.

---

## 💡 Skills Demonstrated

- Excel VBA (Modular macros, loops, custom functions)
- Data Transformation & Cleaning
- String Matching Algorithms (Levenshtein Distance)
- Report Automation
- Debugging and Logic Building

---

## 👩‍💻 Author

**Chamudi**  
Excel VBA | Data Automation | Logistics Data Processing  
