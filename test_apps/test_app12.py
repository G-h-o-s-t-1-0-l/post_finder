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

        # Flags and variables
        tracking_found = False
        mail_bag_no = None
        hash_value = None  # To store the # value

        # Process each sheet
        for sheet_name, df in sheet_data.items():
            # Iterate through every row and column to search for the tracking number
            for row_idx, row in df.iterrows():
                for col_idx, cell in row.items():
                    # Convert cell to string before searching
                    cell_str = str(cell)

                    # Check if the tracking number exists in the current cell
                    if tracking_no.lower() in cell_str.lower():
                        tracking_found = True
                        # Adjusting the row and column for user-friendly display (starting from 1)
                        adjusted_row_idx = row_idx + 1
                        adjusted_col_idx = col_idx + 1
                        result_html += f"<h3>Tracking No: {tracking_no} found in sheet '{sheet_name}'</h3>"
                        result_html += f"<p><strong>Row {adjusted_row_idx}, Column {adjusted_col_idx}:</strong> {cell_str}</p>"

                        # Fetch Mail Bag No from rows 4 and 5, columns C and D (non-zero-based indexing)
                        try:
                            mail_bag_no_parts = [
                                str(df.iloc[3, 2]) if not pd.isna(df.iloc[3, 2]) else "",
                                str(df.iloc[3, 3]) if not pd.isna(df.iloc[3, 3]) else ""
                            ]
                            mail_bag_no = " ".join(part for part in mail_bag_no_parts if part.strip())
                        except IndexError:
                            mail_bag_no = None

                        # Fetch # value from column A in the same row as the tracking number
                        try:
                            hash_value = df.iloc[row_idx, 0]  # Column A is index 0
                            if pd.isna(hash_value):
                                hash_value = None
                        except IndexError:
                            hash_value = None

                        # Break inner loops once tracking number is found
                        break
                if tracking_found:
                    break
            if tracking_found:
                break

        # Display the Mail Bag No and # value if found
        if tracking_found and mail_bag_no:
            result_html += f"<p><strong>Mail Bag No:</strong> {mail_bag_no}</p>"
        elif tracking_found:
            result_html += f"<p><strong>Mail Bag No:</strong> Not found</p>"

        if tracking_found and hash_value:
            result_html += f"<p><strong>#:</strong> {hash_value}</p>"
        elif tracking_found:
            result_html += f"<p><strong>#:</strong> Not found</p>"

        # If no tracking number was found, display a message
        if not tracking_found:
            result_html += f"<p>No data found for Tracking No: {tracking_no}</p>"

        return render_template_string(result_html)

    except Exception as e:
        return f"Error processing file: {e}", 500


if __name__ == '__main__':
    app.run(debug=True, port=5000, host='0.0.0.0')
