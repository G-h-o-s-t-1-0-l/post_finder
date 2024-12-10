import pandas as pd
from flask import Flask, request, render_template_string
import os
import re

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
            <input type="text" name="tracking_no" placeholder="Enter Tracking No">
            <br><br>
            <input type="submit" value="Upload">
        </form>
    '''

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files or 'tracking_no' not in request.form:
        return "No file or tracking number provided.", 400

    file = request.files['file']
    tracking_no = request.form['tracking_no']
    filename = file.filename

    if not filename.endswith(('.xls', '.xlsx')):
        return "Unsupported file type. Please upload an Excel file.", 400

    try:
        # Read all sheets from the Excel file with no header
        file.seek(0)  # Reset file pointer
        sheet_data = pd.read_excel(file, sheet_name=None, header=None, dtype=str)

        result_html = "<h2>Tracking Information</h2>"

        # Flag to track if the tracking number is found
        tracking_found = False

        # Define regex pattern for Mail Bag No (e.g., HDOD-L04-2024-47 or similar formats)
        mail_bag_pattern = r"\b[A-Za-z]{4,6}-?[A-Za-z0-9]*[-]?[A-Za-z0-9]+\b"  # Adjusted pattern to capture variety of formats

        # Process each sheet
        for sheet_name, df in sheet_data.items():
            sheet_found = False  # Track if mail bag or # is found in the sheet

            # Iterate through every row and column to search for the tracking number
            for row_idx, row in df.iterrows():
                for col_idx, cell in row.items():
                    # Convert cell to string before searching
                    cell_str = str(cell)

                    # Check for tracking number match
                    if tracking_no.lower() in cell_str.lower():
                        tracking_found = True
                        adjusted_row_idx = row_idx + 1
                        adjusted_col_idx = col_idx + 1
                        result_html += f"<h3>Tracking No: {tracking_no} found in sheet '{sheet_name}'</h3>"
                        result_html += f"<p><strong>Row {adjusted_row_idx}, Column {adjusted_col_idx}:</strong> {cell_str}</p>"

                    # Check for "Mail Bag No" match using regex pattern
                    if re.search(mail_bag_pattern, cell_str):
                        mail_bag_no = cell_str.strip()
                        result_html += f"<p><strong>Mail Bag No:</strong> {mail_bag_no}</p>"
                        sheet_found = True

                    # Check for "#" match
                    if "#" in cell_str:
                        hash_value = cell_str.strip()
                        result_html += f"<p><strong>#:</strong> {hash_value}</p>"
                        sheet_found = True

            # If either "Mail Bag No" or "#" was found, display the sheet information
            if sheet_found:
                result_html += f"<p>Details found in sheet '{sheet_name}'.</p>"

            # Once the tracking number is found, no need to process further sheets
            if tracking_found:
                break

        # If no tracking number was found, display a message
        if not tracking_found:
            result_html += f"<p>No data found for Tracking No: {tracking_no}</p>"

        return render_template_string(result_html)

    except Exception as e:
        return f"Error processing file: {e}", 500


if __name__ == '__main__':
    app.run(debug=True, port=5000, host='0.0.0.0')
