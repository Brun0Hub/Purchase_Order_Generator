📄 Order Document Generator

🚀 Overview
This project automates the creation of Word documents (.docx) for customer orders.
It reads order data from a CSV file and replaces placeholders in a Word template with real values such as order number, customer details, product information, prices, and automatically generated dates.

The result is a professional order report generated for each entry in the CSV file.

✨ Features

📊 CSV Integration: Reads order data from a structured CSV file.

📝 Word Document Automation: Uses a .docx template with placeholders (e.g., AAAA, BBBB, PPPP) and replaces them with actual values.

📅 Automatic Date Handling: Inserts the current day, month (in text), and year into the document.

💰 Currency Formatting: Formats unit and total prices into Brazilian currency style (R$ 1.234,56).

🔄 Batch Processing: Generates multiple order documents in one run, one per CSV row.

📂 File Management: Saves each generated document with a unique name based on order number and status.

🐍 Python Concepts and Libraries Used

📄 python-docx: Opens and edits Word documents, replaces placeholders with actual values.

📑 csv: Reads structured data from CSV files using DictReader.

📁 os: Checks if a file already exists before saving, prevents overwriting.

⏰ datetime: Captures the current system date and extracts day, month, and year.

🌐 locale: Ensures month names are displayed in Portuguese (e.g., "abril").

🔤 String Formatting: Formats numeric values into currency strings with proper separators (R$ 1.234,56).

⚙️ How It Works
Prepare a Word template (Order.docx) with placeholders like AAAA, BBBB, PPPP, QQQQ, RRRR.

Prepare a CSV file (orders.csv) with columns such as Order Number, Customer, Product, Unit Price, etc.

Run the script:

bash
python main.py
The script generates one Word document per order, replacing placeholders with actual values and saving them in the project folder.

🖼️ Example
Template Placeholder:

Código
Order Number: AAAA
Customer: CCCC
Unit Price: KKKK
Total Price: LLLL
Date: PPPP / QQQQ / RRRR
Generated Output:

Código
Order Number: PED0001
Customer: Distribuidora Gama
Unit Price: R$ 102,34
Total Price: R$ 2.456,16
Date: 23 / abril / 2026
🎯 Benefits
⏱️ Saves time by automating repetitive document creation.

✅ Ensures consistency across all order reports.

🔧 Easy to adapt for different templates or languages.
