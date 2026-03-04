# 🚀 Automated Debt Collection & Banking Notification System

A professional Python-based automation framework designed to generate mass debt collection notices and formal banking correspondence. This system transforms raw financial data from Excel into formatted PDF documents, ready for distribution.

## 📈 Business Impact
In banking operations, managing delinquent accounts requires precise and timely communication. This tool automates the generation of:
* **Reminders (Surat Peringatan):** For early-stage arrears.
* **Formal Notices (Somasi):** For legal escalations in debt recovery.
* **Billing Statements:** For routine payment notifications.

By automating this workflow, a process that typically takes days of manual data entry can be completed in minutes, ensuring 100% accuracy in financial figures and customer details.

## ✨ Key Features
* **Mass Notification Engine:** Generate thousands of personalized letters in a single batch run.
* **Smart Financial Formatting:** Automatically applies Indonesian Rupiah (IDR) currency standards to all numerical values while preserving sensitive formats like Account Numbers.
* **Legal-Grade Precision:** Uses Microsoft Word (`.docx`) templates to maintain official banking letterheads and legal watermarks.
* **Headless PDF Conversion:** High-speed, server-side conversion using LibreOffice CLI for industrial-scale output.

## 🛠️ Technical Stack
* **Python 3.x**
* **Pandas:** For large-scale data parsing and manipulation.
* **Docxtpl:** For high-fidelity document rendering.
* **LibreOffice CLI:** The engine for professional PDF generation.

## 🚀 Deployment & Usage
1.  **Data Input:** Place your debtor database in the `data/` folder.
2.  **Template Setup:** Design your official letter in Word using `{{ placeholder }}` tags.
3.  **Run Automation:**
    ```bash
    python main.py
    ```

## 🛡️ Security & Privacy
This project is designed with data privacy in mind. The `.gitignore` configuration ensures that sensitive debtor information and banking records are never uploaded to public repositories.

---
**Author:** Azmi Khalid  
*Vice President in State-Owned Banking | Expert in Banking Operations & Fintech Automation*
