import pandas as pd
from flask import Flask, request, render_template_string
import io

app = Flask(__name__)

@app.route('/')
def index():
    return '''
        <h2>Upload File</h2>
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

    try:
        # Determine file type and read accordingly
        if filename.endswith('.csv'):
            df = pd.read_csv(file, header=None, dtype=str)
        elif filename.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(file, header=None, dtype=str)
        else:
            return "Unsupported file type. Please upload a CSV or Excel file.", 400

        # Find the "Mail Bag No:" label and retrieve its value from column C (index 2)
        mail_bag_no = None
        for i, row in df.iterrows():
            if row[0] == "Mail Bag No:":
                mail_bag_no = row[2]
                break

        if mail_bag_no is None:
            mail_bag_no = "N/A"

        # Locate header row by looking for "Tracking No" in any row
        header_row_index = None
        for i, row in df.iterrows():
            if row.str.contains("Tracking No", case=False, na=False).any():
                header_row_index = i
                break

        if header_row_index is None:
            return "'Tracking No' column not found in the file."

        # Reload the file with the identified header row
        file.seek(0)  # Reset file pointer to start of file for re-read
        if filename.endswith('.csv'):
            df = pd.read_csv(file, skiprows=header_row_index)
        else:
            df = pd.read_excel(file, skiprows=header_row_index)

    except Exception as e:
        return f"Error processing file: {e}", 500

    # Confirm the 'Tracking No' column exists in this adjusted DataFrame
    if 'Tracking No' not in df.columns:
        return f"'Tracking No' column not found. Available columns after adjustment: {', '.join(df.columns)}"

    # Filter the data based on the tracking number
    filtered_data = df[df['Tracking No'] == tracking_no]

    if filtered_data.empty:
        return f"No data found for Tracking No: {tracking_no}"

    # Collect data for each matching entry
    result_html = "<h2>Tracking Information</h2>"
    for _, row in filtered_data.iterrows():
        hash_column = row.get('#', "N/A")
        result_html += f"""
            <p><strong>Tracking No:</strong> {row['Tracking No']}</p>
            <p><strong>Mail Bag No:</strong> {mail_bag_no}</p>
            <p><strong>#:</strong> {hash_column}</p>
            <hr>
        """
    
    return render_template_string(result_html)

if __name__ == '__main__':
    app.run(debug=True, port=5000, host='0.0.0.0')
