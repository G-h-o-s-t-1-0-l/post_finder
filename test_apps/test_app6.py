import pandas as pd
from flask import Flask, request, render_template_string
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

        # Process each sheet
        for sheet_name, df in sheet_data.items():
            # Save each sheet as a CSV
            csv_filename = os.path.join(UPLOAD_FOLDER, f"{sheet_name}.csv")
            df.to_csv(csv_filename, index=False)

            # Display the entire sheet for debugging (without trimming)
            result_html += f"<h3>Full content of sheet '{sheet_name}':</h3>"
            result_html += f"<pre>{df}</pre>"

            # Flag to track if tracking number is found in the sheet
            tracking_found = False

            # Iterate through every row and column to search for the tracking number
            for row_idx, row in df.iterrows():
                for col_idx, cell in row.items():
                    # Convert cell to string before searching
                    cell_str = str(cell)

                    # Check if the tracking number exists in the current cell
                    if tracking_no.lower() in cell_str.lower():
                        tracking_found = True
                        result_html += f"<p><strong>Tracking No:</strong> {tracking_no} found in row {row_idx}, column {col_idx} (value: {cell_str})</p>"

            if not tracking_found:
                result_html += f"<p>No data found for Tracking No: {tracking_no} in sheet: {sheet_name}</p>"

        return render_template_string(result_html)

    except Exception as e:
        return f"Error processing file: {e}", 500


if __name__ == '__main__':
    app.run(debug=True, port=5000, host='0.0.0.0')
