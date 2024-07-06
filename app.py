from flask import Flask, render_template, request, send_file
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import os
import zipfile
import shutil
import platform

app = Flask(__name__, template_folder='templates')  # Make sure 'templates' folder exists and contains index.html

def get_downloads_folder():
    if platform.system() == "Windows":
        return os.path.join(os.path.expanduser("~"), "Downloads")
    else:
        return os.path.join(os.path.expanduser("~"), "Downloads")

def convert_column_names(df):
    return {col: col.replace(' ', '_') for col in df.columns}

def generate_reports(data_file, template_file):
    # Determine file type and load data
    if data_file.filename.endswith('.csv'):
        df = pd.read_csv(data_file)
    elif data_file.filename.endswith('.xlsx'):
        df = pd.read_excel(data_file)
    else:
        raise ValueError('Unsupported file format!')

    # Convert spaced column names to snake_case
    df.rename(columns=convert_column_names(df), inplace=True)

    # Load DOCX template
    doc = DocxTemplate(template_file)
    
    # Today's date in mm/dd/yyyy format
    today_date = datetime.today().strftime('%m/%d/%Y')

    # Create a folder in the Downloads directory to store reports
    downloads_folder = get_downloads_folder()
    reports_folder = os.path.join(downloads_folder, 'reports')
    os.makedirs(reports_folder, exist_ok=True)

    # Iterate through data rows and generate reports
    for index, row in df.iterrows():
        data = row.to_dict()
        data['TODAY'] = today_date
        doc.render(data)
        report_filename = os.path.join(reports_folder, f"report_{index}.docx")
        doc.save(report_filename)

    # Zip the reports folder
    zip_filename = os.path.join(downloads_folder, 'reports.zip')
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for root, _, files in os.walk(reports_folder):
            for file in files:
                zipf.write(os.path.join(root, file), file)

    # Remove the reports folder after zipping
    shutil.rmtree(reports_folder)

    return zip_filename

# Route to serve the main HTML page
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle file upload and report generation
@app.route('/generate', methods=['POST'])
def generate():
    data_file = request.files['data_file']
    template_file = request.files['template_file']
    try:
        zip_filename = generate_reports(data_file, template_file)
        return send_file(zip_filename, as_attachment=True)
    except Exception as e:
        return str(e), 500

# Route to serve manifest.json
@app.route('/manifest.json')
def serve_manifest():
    try:
        return send_file('static/manifest.json', mimetype='application/json')  # Adjust path if necessary
    except Exception as e:
        app.logger.error(f"Error serving manifest.json: {e}")
        return str(e), 500  # Return error message and HTTP status 500 for internal server error

# Route to serve service worker (sw.js)
@app.route('/sw.js')
def serve_sw():
    try:
        return send_file('static/sw.js', mimetype='application/javascript')  # Adjust path if necessary
    except Exception as e:
        app.logger.error(f"Error serving sw.js: {e}")
        return str(e), 500  # Return error message and HTTP status 500 for internal server error

if __name__ == '__main__':
    app.run()
