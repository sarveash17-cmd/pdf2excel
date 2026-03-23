from flask import Flask, request, render_template, jsonify, send_file
import os
import time
import threading
import uuid
import shutil
from validation import validate_pdf
from specific_extractor import generate_specific_json
from excel_exporter import extract_bescom_to_excel
from pdf_to_excel import convert_pdf_to_excel

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

conversion_progress = {}

def convert_format(filename, job_id, template_path, format_type):
    try:
        conversion_progress[job_id] = {
            'status': 'starting', 'progress': 0, 'message': 'Initializing...',
            'excel_url': None, 'error': None
        }

        if format_type == 'specific':
            conversion_progress[job_id]['progress'] = 20
            conversion_progress[job_id]['message'] = 'Extracting Specific Data...'
            json_path = os.path.join(OUTPUT_FOLDER, f"{job_id}.json")
            generate_specific_json(filename, json_path)
            
            conversion_progress[job_id]['progress'] = 60
            conversion_progress[job_id]['message'] = 'Populating Excel Template...'
            excel_path = os.path.join(OUTPUT_FOLDER, f"{job_id}_specific.xlsx")
            extract_bescom_to_excel(json_path, excel_path, template_path=template_path)
            download_url = f'/download_excel/specific/{job_id}'
        else:
            conversion_progress[job_id]['progress'] = 30
            conversion_progress[job_id]['message'] = 'Extracting All Tables from Document...'
            excel_path = os.path.join(OUTPUT_FOLDER, f"{job_id}_all.xlsx")
            convert_pdf_to_excel(filename, excel_path)
            download_url = f'/download_excel/all/{job_id}'
        
        # Finish
        conversion_progress[job_id]['progress'] = 100
        conversion_progress[job_id]['status'] = 'completed'
        conversion_progress[job_id]['message'] = 'File ready for download!'
        conversion_progress[job_id]['excel_url'] = download_url

    except Exception as e:
        conversion_progress[job_id]['status'] = 'error'
        conversion_progress[job_id]['error'] = str(e)
        conversion_progress[job_id]['message'] = f'Error: {str(e)}'

@app.route('/')
def upload_form():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_pdf():
    uploaded_pdf = request.files.get('pdf')
    uploaded_template = request.files.get('template')
    format_type = request.form.get('format_type')
    
    if format_type not in ['specific', 'all']:
        return jsonify({'status': 'error', 'message': 'Invalid format type'}), 400
        
    if not uploaded_pdf or uploaded_pdf.filename == '':
        return jsonify({'status': 'error', 'message': 'PDF Document is required'}), 400
        
    if format_type == 'specific' and (not uploaded_template or uploaded_template.filename == ''):
        return jsonify({'status': 'error', 'message': 'Excel Template is required for Specific Extraction'}), 400
        
    job_id = str(uuid.uuid4())
    pdf_filename = os.path.abspath(os.path.join(app.config['UPLOAD_FOLDER'], f"{job_id}.pdf"))
    uploaded_pdf.save(pdf_filename)
    
    errors = validate_pdf(pdf_filename)
    if errors:
        return jsonify({'status': 'error', 'message': ' | '.join(errors)}), 400
        
    template_filename = None
    if uploaded_template and uploaded_template.filename != '':
        template_filename = os.path.abspath(os.path.join(app.config['UPLOAD_FOLDER'], f"{job_id}_template.xlsx"))
        uploaded_template.save(template_filename)

    thread = threading.Thread(target=convert_format, args=(pdf_filename, job_id, template_filename, format_type))
    thread.daemon = True
    thread.start()
    return jsonify({'status': 'ok', 'job_id': job_id})

@app.route('/progress/<job_id>')
def get_progress(job_id):
    if job_id in conversion_progress:
        return jsonify(conversion_progress[job_id])
    return jsonify({'status': 'not_found', 'error': 'Job not found'}), 404

@app.route('/download_excel/<format_type>/<job_id>')
def download_excel(format_type, job_id):
    if format_type not in ['specific', 'all']:
        return "Invalid format type", 400
    
    filename = f"{job_id}_{format_type}.xlsx"
    excel_path = os.path.abspath(os.path.join(app.config['OUTPUT_FOLDER'], filename))
    
    download_name = "Populated_Template.xlsx" if format_type == 'specific' else "Full_Document_Data.xlsx"
    
    if os.path.exists(excel_path):
        return send_file(excel_path, as_attachment=True, download_name=download_name)
            
    return "Excel file not found. It may have been deleted or never created.", 404

if __name__ == '__main__':
    app.run(debug=True, port=5000)
