import pandas as pd
from flask import Flask, request, render_template, session, redirect, url_for
import os

app = Flask(__name__)
app.secret_key = "supersecretkey"  # Required for session management

# Directory to store temporarily uploaded files
UPLOAD_FOLDER = 'uploaded_files'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    """Page for uploading the file."""
    if request.method == 'POST':
        if 'file' in request.files and request.files['file'].filename:
            file = request.files['file']
            filename = file.filename

            if not filename.endswith(('.xls', '.xlsx')):
                error = "Unsupported file type. Please upload an Excel file."
                return render_template('upload.html', error=error)

            # Save the uploaded file
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(file_path)
            session['uploaded_file'] = file_path  # Save the file path to the session
            session['uploaded_file_name'] = filename  # Save the file name to the session

            return redirect(url_for('search_tracking'))

    return render_template('upload.html')


@app.route('/search', methods=['GET', 'POST'])
def search_tracking():
    """Page for searching the tracking number."""
    file_path = session.get('uploaded_file')
    uploaded_file_name = session.get('uploaded_file_name', "No file selected")
    result_html = ""

    if request.method == 'POST' and file_path:
        tracking_no = request.form.get('tracking_no')
        try:
            # Read all sheets from the Excel file
            sheet_data = pd.read_excel(file_path, sheet_name=None, header=None, dtype=str)

            # Flags and variables
            tracking_found = False
            mail_bag_no = None
            hash_value = None  # To store the # value

            # Process each sheet
            for sheet_name, df in sheet_data.items():
                for row_idx, row in df.iterrows():
                    for col_idx, cell in row.items():
                        cell_str = str(cell)

                        if tracking_no.lower() in cell_str.lower():
                            tracking_found = True
                            adjusted_row_idx = row_idx + 1
                            adjusted_col_idx = col_idx + 1
                            result_html += f"<h3>Tracking No: {tracking_no} found in sheet '{sheet_name}'</h3>"
                            ##result_html += f"<p><strong>Row {adjusted_row_idx}, Column {adjusted_col_idx}:</strong> {cell_str}</p>"##

                            try:
                                mail_bag_no_parts = [
                                    str(df.iloc[3, 2]) if not pd.isna(df.iloc[3, 2]) else "",
                                    str(df.iloc[3, 3]) if not pd.isna(df.iloc[3, 3]) else ""
                                ]
                                mail_bag_no = " ".join(part for part in mail_bag_no_parts if part.strip())
                            except IndexError:
                                mail_bag_no = None

                            try:
                                hash_value = df.iloc[row_idx, 0]
                                if pd.isna(hash_value):
                                    hash_value = None
                            except IndexError:
                                hash_value = None

                            break
                    if tracking_found:
                        break
                if tracking_found:
                    break

            if tracking_found and mail_bag_no:
                result_html += f"<p ><strong>Mail Bag No:</strong> { mail_bag_no}</p>"
            elif tracking_found:
                result_html += f"<p><strong>Mail Bag No:</strong> Not found</p>"

            if tracking_found and hash_value:
                result_html += f"<p><strong>Number:</strong> { hash_value}</p>"
            elif tracking_found:
                result_html += f"<p><strong>#:</strong> Not found</p>"

            if not tracking_found:
                result_html += f"<p>No data found for Tracking No: {tracking_no}</p>"

        except Exception as e:
            result_html = f"<p style='color: red;'>Error processing file: {e}</p>"

    elif request.method == 'POST' and not file_path:
        result_html = "<p style='color: red;'>Please upload a file before searching.</p>"

    return render_template('search.html', result_html=result_html, uploaded_file_name=uploaded_file_name)


if __name__ == '__main__':
    app.run(debug=True, port=5000, host='0.0.0.0')
