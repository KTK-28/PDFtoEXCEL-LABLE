import os
import re
import json
import pandas as pd
import fitz
import qrcode
from flask import Flask, render_template, request, send_from_directory, url_for
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
EXCEL_FOLDER = 'excels'
LABELS_FOLDER = 'labels'
ALLOWED_EXTENSIONS = {'pdf'}

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

if not os.path.exists(EXCEL_FOLDER):
    os.makedirs(EXCEL_FOLDER)

if not os.path.exists(LABELS_FOLDER):
    os.makedirs(LABELS_FOLDER)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_pdf_to_excel(pdf_path):
    customer_details_list = []  # List to store customer details for each page

    # Convert PDF to Excel
    with fitz.open(pdf_path) as pdf_document:
        num_pages = pdf_document.page_count

        # Iterate through each page
        for page_index in range(num_pages):
            # Extract text from the current page
            page = pdf_document[page_index]
            pdf_text = page.get_text()

            # Extract customer details and SKU from the page
            customer_block, sku = extract_customer_details_and_sku(pdf_text)

            # Add the page number, customer details, and SKU to the list
            customer_details_list.append({
                'Page': page_index + 1,
                'Customer Block': customer_block,
                'SKU': sku
            })

    # Save the list of customer details to a JSON file as a valid JSON array
    json_file_path = os.path.join(EXCEL_FOLDER, 'customer_details.json')
    with open(json_file_path, 'w', encoding='utf-8') as json_file:
        json.dump(customer_details_list, json_file, ensure_ascii=False, indent=2)

    # Create Excel file
    excel_file_path = convert_json_to_excel(json_file_path)

    return excel_file_path

def extract_customer_details_and_sku(pdf_text):
    lines = pdf_text.split('\n')
    customer_block = []
    sku = None

    collect_data = False
    for line in lines:
        if "Quantity Product Details Unit price Order Totals" in line:
            collect_data = True
            continue
        if collect_data:
            if "SKU:" in line:
                sku_match = re.search(r'SKU:\s*(.*?)\s*//', line)
                if sku_match:
                    sku = sku_match.group(1).strip()
                    break

    return customer_block, sku


def convert_json_to_excel(json_file_path):
    # Load customer data from the JSON file
    with open(json_file_path, 'r', encoding='utf-8') as json_file:
        customer_data_list = json.load(json_file)

    # Create a DataFrame from the list
    df = pd.DataFrame(customer_data_list[0]['Customer Block'])

    # Add a serial number column
    df.insert(0, 'Sr.No', range(1, len(df) + 1))

    # Generate the Excel file name based on the JSON file name
    excel_filename = os.path.splitext(os.path.basename(json_file_path))[0] + '_customer_details.xlsx'

    # Construct the full path to the Excel file
    excel_file_path = os.path.join(EXCEL_FOLDER, excel_filename)

    # Save the DataFrame to an Excel file
    df.to_excel(excel_file_path, index=False)

    print(f'Customer details saved to {excel_file_path}')

    return excel_file_path


def extract_customer_details(pdf_text):
    # Extract customer details from PDF text
    lines = pdf_text.split('\n')
    customer_block = []

    # Initialize variables to store customer details
    customer_data = {
        'Customer Name': None,
        'City': None,
        'State': None,
        'Pin Code': None,
        'Phone Number': None,
        'Order ID': None
    }

    # Flag to indicate when to start and stop collecting data
    collect_data = False

    # Iterate through lines to find relevant information
    for line in lines:
        # Start collecting data when the line contains "Ship to:"
        if "Ship to:" in line:
            collect_data = True

        # Stop collecting data when the line contains "Order ID:"
        if "Order ID:" in line:
            collect_data = False
            # Extract Order ID from the line
            order_id_match = re.match(r'Order ID:\s*(\S+)', line)
            if order_id_match:
                customer_data['Order ID'] = order_id_match.group(1).strip()

        # Collect data when the flag is set
        if collect_data:
            # Extract customer name from the line after "Ship to:"
            if "Ship to:" in line:
                customer_data['Customer Name'] = lines[lines.index(line) + 1].strip()

            # Extract phone number when the line contains "Phone :"
            if "Phone :" in line:
                customer_data['Phone Number'] = line.replace("Phone :", "").strip()

            # Extract city, state, and pin code from lines containing a combination of them
            address_match = re.match(r'(.*?),\s*(.+?)\s+(\d{6})$', line)
            if address_match:
                customer_data['City'] = address_match.group(1).strip()
                customer_data['State'] = address_match.group(2).strip()
                customer_data['Pin Code'] = address_match.group(3).strip()

    # Append customer details to the list
    customer_block.append(customer_data)

    return customer_block

def extract_customer_details_for_labels(pdf_path):
    customer_details_list = []  # List to store customer details for each page

    # Open the PDF file
    pdf_document = fitz.open(pdf_path)

    for page_index in range(pdf_document.page_count):
        # Extract text from a specific page
        page = pdf_document.load_page(page_index)
        page_text = page.get_text()

        # Find all occurrences of "Ship to:" and "Phone :" pairs
        matches = re.finditer(r'Ship to:(.*?)(Order ID:.*?|$)', page_text, re.DOTALL)

        for match in matches:
            # Extract the "Ship to:" line and the lines until "Order ID:"
            lines = match.group(1).strip().split('\n')

            # Check if "Order ID:" is present in the lines
            if any("Order ID:" in line for line in lines):
                # Stop extracting data if "Order ID:" is found
                break

            # Join lines with line breaks and append customer details to the list
            customer_block = [line.strip() for line in lines]
            customer_details_list.append({
                'Page': page_index + 1,
                'Customer Block': customer_block
            })

    # Save the list of customer details to a JSON file as a valid JSON array
    json_file_path = os.path.join(LABELS_FOLDER, 'customer_details_labels.json')
    with open(json_file_path, 'w', encoding='utf-8') as json_file:
        json.dump(customer_details_list, json_file, ensure_ascii=False, indent=2)

    print(f'Customer details for labels saved to {json_file_path}')

    return json_file_path

def create_labels_from_json(json_file_path):
    # Load customer data from the JSON file
    with open(json_file_path, 'r', encoding='utf-8') as json_file:
        customer_data_list = json.load(json_file)

    # Create a PDF file for all labels
    pdf_path = os.path.join(LABELS_FOLDER, 'all_labels.pdf')
    c = canvas.Canvas(pdf_path, pagesize=(650, 900))  # Increase the page size here

    # Generate labels for each customer
    for customer_data in customer_data_list:
        generate_label(c, customer_data)

    # Save the combined PDF
    c.save()
    print(f'All labels generated in: {pdf_path}')

    return pdf_path

def generate_label(c, customer_data):
    # Set font size for labels
    label_font_size = 22 # You can adjust this value based on your preference

    # Add "Book Post" at the top center
    c.setFont("Helvetica-Bold", label_font_size + 4)
    c.drawCentredString(330, 850, "REGD-Book Post")

    # Ship to (Customer details)
    c.setFont("Helvetica-Bold", label_font_size + 2)  # Set the font to bold for "Ship to:"
    
    # Draw "To:" on the left side
    c.drawString(100, 810, "To:")
    
    # Draw "Printed Books" on the right side
    printed_books_text = "Printed Books"
    printed_books_width = c.stringWidth(printed_books_text, "Helvetica-Bold", label_font_size + 2)
    c.drawString(590 - printed_books_width, 820, printed_books_text)

    # Leave one line after "Ship to:"
    c.drawString(100, 820, "")

    # Set the font to bold for customer details
    c.setFont("Helvetica-Bold", label_font_size + 10)  # Adjust font size for customer details

    # Iterate through customer details and draw each line
    i = 0
    for line in customer_data['Customer Block']:
        words = line.split()
        num_words = len(words)
        if num_words > 0:
            # Process each word to fit within the 25-character limit
            remaining_line = ""
            for word in words:
                if len(remaining_line) + len(word) + 1 <= 25:
                    remaining_line += word + " "
                else:
                    # Print the remaining line on the current line
                    c.drawString(100, 820 - (i + 2) * (label_font_size + 12), remaining_line.strip())
                    remaining_line = word + " "
                    i += 1
            # Print the remaining words on the next line
            c.drawString(100, 820 - (i + 2) * (label_font_size + 12), remaining_line.strip())
            i += 1

    # Add more empty lines before seller details
    for _ in range(3):
        c.drawString(100, 820 - (i + 4) * (label_font_size + 12), "")  # Blank line
        i += 1

    # From (Seller details)
    seller_details = [
        "From:",
        "ACHLESHWAR",
        "129A, GALI NO 5,",
        "SHRI RAM NAGAR,",
        "BEHIND P.F OFFICE, JODHPUR",
        "PHONE 7976644374."
    ]
    
    c.setFont("Helvetica", label_font_size + 1)  # Set the font to bold for "From:"
    for j, line in enumerate(seller_details):
        c.drawString(100, 820 - (i + j) * (label_font_size + 12), line)  # Adjust spacing

    # Generate and add QR code
    qr_content = "https://www.amazon.in/stores/Achleshwar/page/065EAADA-5EFB-4DA7-9694-AC2FC2D320E9?ref_=ast_bln&store_ref=bl_ast_dp_brandLogo_sto"  # Replace with your content
    qr = qrcode.make(qr_content)

    # Resize the QR code image if needed
    qr = qr.resize((140, 140))  # Adjust dimensions as needed

    # Draw the QR code image on the label canvas after the seller address
    qr_x = 90
    qr_y = 720 - (i + j + 2) * (label_font_size + 12)
    c.drawInlineImage(qr, qr_x, qr_y, width=qr.width, height=qr.height)

    qr_text = "Scan QR & Visit "
    qr_text_width = c.stringWidth(qr_text, "Helvetica-Bold", label_font_size + 4)
    text_x = qr_x + qr.width / 7 - qr_text_width / 7
    text_y = qr_y - 30
    c.drawString(text_x, text_y, qr_text)
    qr_text = "Achleshwar's Amazon Store :)"
    qr_text_width = c.stringWidth(qr_text, "Helvetica-Bold", label_font_size + 6)
    text_x = qr_x + qr.width / 8 - qr_text_width / 8
    text_y = qr_y - 55
    c.drawString(text_x, text_y, qr_text)
    # Show a new page for each label
    c.showPage()


def safe_str(value):
    return str(value) if value is not None else ''

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'

    file = request.files['file']

    if file.filename == '':
        return 'No selected file'

    if file and allowed_file(file.filename):
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        # Convert PDF to Excel
        excel_file_path = convert_pdf_to_excel(filepath)

        # Extract data for labels
        labels_json_path = extract_customer_details_for_labels(filepath)

        # Create labels
        labels_path = create_labels_from_json(labels_json_path)

        # Update the template context with the correct links
        excel_link = url_for('download_excel', filename=os.path.basename(excel_file_path))
        labels_link = url_for('download_labels', filename=os.path.basename(labels_path))

        return render_template('index.html', excel_link=excel_link, labels_link=labels_link)

    return 'Invalid file type'

@app.route('/download_excel/<filename>')
def download_excel(filename):
    return send_from_directory(EXCEL_FOLDER, filename, as_attachment=True)

@app.route('/download_labels/<filename>')
def download_labels(filename):
    return send_from_directory(LABELS_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
