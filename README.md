# 📊 Credit Card Expense Control — Streamlit Dashboard

This repository contains an interactive application developed in **Streamlit** for consolidating and visually analyzing credit card statements. The application replaces and extends the original script based on `openpyxl`, allowing statements to be dynamically loaded via a web interface instead of reading fixed local files.

## 🚀 Features

- **Dynamic File Upload:** Support for multiple simultaneous files (invoices `.csv` and `.xlsx`).
- **Real-Time Metrics (KPIs):** Instant visualization of Total Purchases, Total Payments/Credits, and Net Balance.
- **Interactive Charts (Plotly):**
- **Spending by Category:** Donut chart detailing the percentage distribution of purchases.
- **Spending by Card:** Bar chart comparing the volume of spending between submitted invoices.
- **Consolidated Transaction Table:** Interactive tabular visualization with support for sorting, searching, and native Streamlit pagination, with appropriate monetary (R$) and date formatting.
- **Automatic Categorization:** Regular Expression (RegEx) based inference engine to suggest categories based on transaction descriptions.

## 🛠️ Technologies Used

- **Python 3.10+**
- **Streamlit** (User interface and interactivity)
- **Pandas** (Data processing, filtering, and consolidation)
- **Plotly Express** (Dynamic and interactive graphs)
- **Openpyxl** (Excel file reading engine)
