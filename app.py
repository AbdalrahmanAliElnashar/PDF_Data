from flask import Flask, request, jsonify, send_file
import os
import pandas as pd
from pdf2image import convert_from_path
from img2table.ocr import PaddleOCR
from img2table.document import Image
import traceback

app = Flask(__name__)



# Function to convert the first page of a PDF to an image
def convert_pdf_to_image(pdf_path):
    try:
        images = convert_from_path(pdf_path, first_page=1, last_page=1)  # Extract only the first page
        if images:
            image_path = 'Photo_1.png'
            images[0].save(image_path, 'PNG')  # Save the first page as an image
            return image_path
        else:
            return None
    except Exception as e:
        print(f"Error in PDF to image conversion: {e}")
        return None

# Function to extract tables from the image
def extract_table_from_image(image_path):
    try:
        ocr = PaddleOCR(lang="en")
        doc = Image(image_path)
        output_path = 'result.xlsx'
        doc.to_xlsx(dest=output_path,
                    ocr=ocr,
                    implicit_rows=False,
                    implicit_columns=False,
                    borderless_tables=False,
                    min_confidence=50)
        return output_path
    except Exception as e:
        print(f"Error in table extraction: {e}")
        return None

# Function to extract specific data from the Excel table
def extract_specific_data(file_path):
    try:
        data = pd.read_excel(file_path)
        # Checking if required columns exist
        expected_columns = ['Course Code', 'Course Name', 'Details']
        if all(col in data.columns for col in expected_columns):
            extracted_data = data[expected_columns].dropna()
            extracted_data.rename(columns={'Details': 'Section'}, inplace=True)
            return extracted_data.to_dict(orient='records')
        else:
            return None
    except Exception as e:
        print(f"Error in data extraction: {e}")
        return None

# API Route to upload the PDF file
@app.route('/upload', methods=['POST'])
def upload_pdf():
    try:
        if 'pdf' not in request.files:
            return jsonify({'error': 'No file part'}), 400
        
        pdf_file = request.files['pdf']
        
        if pdf_file.filename == '':
            return jsonify({'error': 'No selected file'}), 400
        
        # Save the PDF to the uploads directory
        pdf_path = os.path.join('uploads', pdf_file.filename)
        pdf_file.save(pdf_path)

        # Convert PDF to a single image (first page)
        image_path = convert_pdf_to_image(pdf_path)
        if not image_path:
            return jsonify({'error': 'Failed to convert PDF to image'}), 500
        
        # Extract table from the image
        excel_file = extract_table_from_image(image_path)
        if not excel_file:
            return jsonify({'error': 'Failed to extract table from image'}), 500

        # Extract specific data from the Excel table
        extracted_data = extract_specific_data(excel_file)
        if extracted_data is None:
            return jsonify({'error': 'Required columns not found'}), 500

        return jsonify({'data': extracted_data})

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

# API Route to download the Excel file
@app.route('/download', methods=['GET'])
def download_excel():
    try:
        excel_file_path = 'result.xlsx'
        if os.path.exists(excel_file_path):
            return send_file(excel_file_path, as_attachment=True)
        else:
            return jsonify({'error': 'Excel file not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    app.run(debug=True)
