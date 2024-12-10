import pandas as pd
from flask import Flask, request, render_template_string, send_from_directory
import os

app = Flask(__name__)

# Directory to store converted CSV files
UPLOAD_FOLDER = 'converted_files'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return '''
        <h2>Upload Excel File</h2>
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file">
            <br><br>
            <input type="submit" value="Upload">
        </form>
    '''

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file provided.", 400

    file = request.files['file']
    filename = file.filename

    if not filename.endswith(('.xls', '.xlsx')):
        return "Unsupported file type. Please upload an Excel file.", 400

    try:
        # Read all sheets from the Excel file
        file.seek(0)  # Reset file pointer
        sheet_data = pd.read_excel(file, sheet_name=None, dtype=str)

        # Convert each sheet to a separate CSV file
        csv_files = []
        for sheet_name, df in sheet_data.items():
            csv_filename = os.path.join(UPLOAD_FOLDER, f"{sheet_name}.csv")
            df.to_csv(csv_filename, index=False)
            csv_files.append(csv_filename)

        # Generate a response with download links for all CSV files
        result_html = "<h2>Converted Sheets</h2><ul>"
        for csv_file in csv_files:
            result_html += f'<li><a href="/download/{os.path.basename(csv_file)}" download>{os.path.basename(csv_file)}</a></li>'
        result_html += "</ul>"

        return render_template_string(result_html)

    except Exception as e:
        return f"Error processing file: {e}", 500

@app.route('/download/<filename>')
def download_file(filename):
    # Serve the CSV file for download
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000, host='0.0.0.0')
