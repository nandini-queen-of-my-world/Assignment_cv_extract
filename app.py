import os
import re
import random
import string
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file, session
from docx import Document
from PyPDF2 import PdfReader
from win32com import client as win32_client
from werkzeug.utils import secure_filename
import subprocess

app = Flask(__name__)
app.secret_key = b'_5#y2L"F4Q8z\n\xec]/'

def extract_info_from_pdf(pdf_file):
    with open(pdf_file, 'rb') as file:
        reader = PdfReader(file)
        text = ''
        for page in reader.pages:
            text += page.extract_text()
    return text

def extract_info_from_docx(docx_file):
    doc = Document(docx_file)
    text = ''
    for para in doc.paragraphs:
        text += para.text
    return text

def extract_info_from_doc(doc_file):
    try:
        # Extract text content from .doc file using LibreOffice
        output = subprocess.check_output(['libreoffice', '--headless', '--convert-to', 'txt:Text', '--outdir', os.path.dirname(doc_file), doc_file])
        text_file_path = os.path.splitext(doc_file)[0] + '.txt'
        with open(text_file_path, 'r', encoding='utf-8') as text_file:
            text = text_file.read()
        return text
    except Exception as e:
        print(f"Error extracting text from {doc_file}: {e}")
        return None


def generate_random_email(name):
    random_number = ''.join(random.choices(string.digits, k=4))
    return f"{name.lower().replace(' ', '')}{random_number}@gmail.com"

def extract_email(text):    
    emails = re.findall(r'(?:\bE-Mailid-)?([\w\.-]+@[\w\.-]+(?:\.com)\b)', text, re.IGNORECASE)    
    cleaned_emails = [re.sub(r'\d$', '', email) for email in emails]
    return list(set(cleaned_emails))

def extract_phone_number(text):
    phone_numbers = re.findall(r'[\+\(]?[1-9]\d{0,2}[\)-]?\s*?\d{2,4}[\s.-]?\d{2,4}[\s.-]?\d{2,4}', text)
    unique_numbers = set()
    for number in phone_numbers:
        formatted_number = re.sub(r'[\s.-]', '', number)
        if len(formatted_number) == 10:
            unique_numbers.add(formatted_number)
        elif formatted_number.startswith('+') and len(formatted_number) > 10:
            unique_numbers.add(formatted_number)
    return ', '.join(unique_numbers)

def process_cv(cv_folder):
    data = []
    for root, dirs, files in os.walk(cv_folder):
        for filename in files:
            file_path = os.path.join(root, filename)
            try:
                if filename.endswith('.pdf'):
                    text = extract_info_from_pdf(file_path)
                elif filename.endswith('.docx'):
                    text = extract_info_from_docx(file_path)
                elif filename.endswith('.doc'):
                    text = extract_info_from_doc(file_path)  # Remove the original .doc file
                else:
                    continue
                email = extract_email(text)
                if not email:
                    name = filename.split('.')[0].split('_')[1]  # Extract name from filename
                    random_email = generate_random_email(name)
                    email.append(random_email)
                phone_number = extract_phone_number(text)
                name = filename.split('.')[0].split('_')[1]  # Replace folder name with an empty string
                data.append({'File Name': name, 'Email': email, 'Phone Number': phone_number, 'Text': text})
            except Exception as e:
                print(f"Error processing {filename}: {e}")
                continue
    return data

def save_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df['Email'] = df['Email'].apply(lambda x: ', '.join(x))
    df.to_excel(output_file, index=False)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'folder' not in request.files:
        return jsonify({'error': 'No folder part'})

    uploaded_files = request.files.getlist('folder')
    if not uploaded_files:
        return jsonify({'error': 'No files uploaded'})

    cv_folder = os.path.join("uploads")
    os.makedirs(cv_folder, exist_ok=True)

    for file in uploaded_files:
        if file.filename == '':
            continue
        file.save(os.path.join(cv_folder, secure_filename(file.filename)))

    cv_data = process_cv(cv_folder)

    output_file = os.path.join(cv_folder, "output.xlsx")
    save_to_excel(cv_data, output_file)

    session['excel_file'] = output_file

    return render_template('index.html', excel_file=output_file)

@app.route('/download')
def download():
    excel_file = session.get('excel_file')
    if not excel_file:
        return jsonify({'error': 'Excel file not found'})

    return send_file(excel_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)