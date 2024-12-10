import pandas as pd
from flask import Flask, request, render_template_string, session
import os

app = Flask(__name__)
app.secret_key = "supersecretkey"  # Required for session management

# Directory to store temporarily uploaded files
UPLOAD_FOLDER = 'uploaded_files'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    result_html = ""
    file_path = session.get('uploaded_file')  # Retrieve the path of the uploaded file from the session
    uploaded_file_name = session.get('uploaded_file_name', "")  # Get the uploaded file name

    if request.method == 'POST':
        # Handle file upload if provided
        if 'file' in request.files and request.files['file'].filename:
            file = request.files['file']
            filename = file.filename

            if not filename.endswith(('.xls', '.xlsx')):
                result_html = "<p style='color: red;'>Unsupported file type. Please upload an Excel file.</p>"
            else:
                # Save the uploaded file
                file_path = os.path.join(UPLOAD_FOLDER, filename)
                file.save(file_path)
                session['uploaded_file'] = file_path  # Save the file path to the session
                session['uploaded_file_name'] = filename  # Save the file name to the session
                uploaded_file_name = filename  # Update the displayed file name
                result_html = "<p style='color: green;'>File uploaded successfully.</p>"

        # Process tracking number search if a file is already uploaded
        if file_path and 'tracking_no' in request.form:
            tracking_no = request.form['tracking_no']
            try:
                # Read all sheets from the Excel file
                sheet_data = pd.read_excel(file_path, sheet_name=None, header=None, dtype=str)

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

            except Exception as e:
                result_html = f"<p style='color: red;'>Error processing file: {e}</p>"

        elif not file_path:
            result_html = "<p style='color: red;'>Please upload a file before searching.</p>"

    # Render the combined page
    return render_template_string('''
        <h2>Upload Excel File</h2>
        <form id="uploadForm" action="/" method="post" enctype="multipart/form-data">
            <label for="fileInput" style="cursor: pointer; padding: 5px; background-color: #f0f0f0; border: 1px solid #ccc;">Browse...</label>
            <input type="file" name="file" id="fileInput" style="display:none;" onchange="updateFileName()">
            <span id="fileName">{{ uploaded_file_name }}</span>
            <br><br>
            <input type="text" name="tracking_no" placeholder="Enter Tracking No">
            <br><br>
            <input type="submit" value="Search">
        </form>
        <hr>
        {{ result_html|safe }

        <script>
            function updateFileName() {
                const fileInput = document.getElementById('fileInput');
                const fileName = document.getElementById('fileName');
                fileName.textContent = fileInput.files.length > 0 ? fileInput.files[0].name : "";
            }
        </script>
    ''', result_html=result_html, uploaded_file_name=uploaded_file_name)


if __name__ == '__main__':
    app.run(debug=True, port=5000, host='0.0.0.0')
