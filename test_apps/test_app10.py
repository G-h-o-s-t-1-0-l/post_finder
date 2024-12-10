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

        # Flag to track if the tracking number is found
        tracking_found = False

        # Process each sheet
        for sheet_name, df in sheet_data.items():
            mail_bag_no = None

            # Iterate through every row and column to search for the tracking number
            for row_idx, row in df.iterrows():
                for col_idx, cell in row.items():
                    cell_str = str(cell)

                    # Check if the tracking number exists in the current cell
                    if tracking_no.lower() in cell_str.lower():
                        tracking_found = True
                        adjusted_row_idx = row_idx + 1
                        adjusted_col_idx = col_idx + 1

                        # Fetch Mail Bag No from rows 4 and 5, columns C and D (non-zero-based indexing)
                        try:
                            mail_bag_no = (
                                str(df.iloc[3, 2]) + " " + str(df.iloc[3, 3]) if not pd.isna(df.iloc[3, 2]) else None
                            )
                        except IndexError:
                            mail_bag_no = None

                        # Add result to the output
                        result_html += f"<h3>Tracking No: {tracking_no} found in sheet '{sheet_name}'</h3>"
                        result_html += f"<p><strong>Row {adjusted_row_idx}, Column {adjusted_col_idx}:</strong> {cell_str}</p>"

                        if mail_bag_no:
                            result_html += f"<p><strong>Mail Bag No:</strong> {mail_bag_no}</p>"
                        else:
                            result_html += "<p><strong>Mail Bag No:</strong> Not found</p>"

            # Stop searching further sheets if the tracking number is found
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
