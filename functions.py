# Import required libraries
from docx import Document   # To manipulate Word (.docx) files
import csv                 # To read CSV files
import os                  # To check if files exist
from datetime import datetime  # To capture the current date
import locale              # To configure language/locale

# Set locale to Brazilian Portuguese so month names appear in Portuguese
locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")

def generate_order(order_number: str, issue_date: str, customer: str, customer_cnpj: str,
                   delivery_address: str, supplier: str, supplier_cnpj: str,
                   supplier_address: str, product: str, quantity: str,
                   unit_price: str, total_price: str, payment_terms: str,
                   order_status: str) -> None:
    """
    Generate a Word document by replacing placeholders with order data.
    Placeholders (AAAA, BBBB, etc.) must exist in the template file 'Order.docx'.
    """

    # Open the Word template containing placeholders
    order_doc = Document("Order.docx")

    # Capture the current system date
    today = datetime.today()
    day = today.strftime("%d")        # Day with two digits (e.g., 23)
    month = today.strftime("%B")      # Month in full text (e.g., April)
    year = today.strftime("%Y")       # Year with four digits (e.g., 2026)

    # Format prices as Brazilian currency (R$ 1.234,56)
    unit_price_formatted = f"R$ {float(unit_price):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    total_price_formatted = f"R$ {float(total_price):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    # Dictionary mapping placeholders to actual values
    order_info = {
        "AAAA": order_number,
        "BBBB": issue_date,
        "CCCC": customer,
        "DDDD": customer_cnpj,
        "EEEE": delivery_address,
        "FFFF": supplier,
        "GGGG": supplier_cnpj,
        "HHHH": supplier_address,
        "IIII": product,
        "JJJJ": quantity,
        "KKKK": unit_price_formatted,
        "LLLL": total_price_formatted,
        "MMMM": payment_terms,
        "OOOO": order_status,
        "PPPP": day,
        "QQQQ": month,
        "RRRR": year
    }

    # Replace placeholders in the Word document
    for paragraph in order_doc.paragraphs:
        for code in order_info:
            if code in paragraph.text:
                paragraph.text = paragraph.text.replace(code, order_info[code])

    # Define output file name
    output_file = f"Updated Order - {order_number} - {order_status}.docx"

    # Save the file only if it does not already exist
    if not os.path.exists(output_file):
        order_doc.save(output_file)
        print(f"File created: {output_file}")
    else:
        print(f"File already exists: {output_file}")

def generate_orders_batch(csv_path: str) -> None:
    """
    Read a CSV file and generate Word documents for each row.
    The CSV must contain columns with exact names:
    Order Number, Issue Date, Customer, CNPJ, Delivery Address,
    Supplier, Supplier CNPJ, Supplier Address, Product,
    Quantity, Unit Price, Total Price, Payment Terms, Order Status
    """
    with open(csv_path, newline="", encoding="utf-8") as file:
        reader = csv.DictReader(file, delimiter=";")
        for row in reader:
            generate_order(
                order_number=row["Número do Pedido"],
                issue_date=row["Data de Emissão"],
                customer=row["Cliente"],
                customer_cnpj=row["CNPJ"],
                delivery_address=row["Endereço de Entrega"],
                supplier=row["Fornecedor"],
                supplier_cnpj=row["CNPJ do Fornecedor"],
                supplier_address=row["Endereço do Fornecedor"],
                product=row["Produto"],
                quantity=row["Quantidade"],
                unit_price=row["Preço Unitário"],
                total_price=row["Preço Total"],
                payment_terms=row["Condição de Pagamento"],
                order_status=row["Status do Pedido"]
            )
