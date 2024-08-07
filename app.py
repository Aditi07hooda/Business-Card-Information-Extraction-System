from flask import Flask, request, render_template, send_file
import pytesseract
from PIL import Image
import os
import pandas as pd
import re

app = Flask(__name__)
pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return render_template('index.html')

def extract_info(text):
    # Define regex patterns for phone number, email, etc.
    phone_pattern = re.compile(r'\b(?:\+?\d{1,3}[-.\s]?)?(?:\(?\d{1,4}\)?[-.\s]?)?\d{1,4}[-.\s]?\d{1,4}[-.\s]?\d{1,9}\b')
    email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b')
    address_pattern = re.compile(r'\d{1,5}\s\w+\s\w+\.?')  # Simplified pattern for addresses

    # Extracting the information
    phone_numbers = phone_pattern.findall(text)
    emails = email_pattern.findall(text)
    addresses = address_pattern.findall(text)
    
    # Assuming the first line is the name and the second line is the company name
    lines = text.split('\n')
    name = lines[0] if lines else ''
    company = lines[1] if len(lines) > 1 else ''
    
    # Returning the extracted information as a dictionary
    return {
        'Name': name.strip(),
        'Phone Numbers': ', '.join(phone_numbers).strip(),
        'Emails': ', '.join(emails).strip(),
        'Addresses': ', '.join(addresses).strip()
    }

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    if file:
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        # Run OCR on the uploaded image
        img = Image.open(file_path)
        text = pytesseract.image_to_string(img, config="--psm 11 --oem 3")

        # Extract specific information
        info = extract_info(text)

        # Save the extracted info into an Excel file
        df = pd.DataFrame([info])
        excel_path = os.path.join(UPLOAD_FOLDER, 'extracted_info.xlsx')
        df.to_excel(excel_path, index=False)

        return send_file(excel_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
