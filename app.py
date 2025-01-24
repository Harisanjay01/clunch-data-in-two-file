from flask import Flask, request, render_template, send_file
import pandas as pd
import os

app = Flask(__name__)

# Directories to store uploaded and processed files
UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return '''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Excel Column Extractor</title>
    </head>
    <body>
        <h1>Upload Excel Files</h1>
        <form action="/process" method="post" enctype="multipart/form-data">
            <label for="file1">File 1:</label>
            <input type="file" name="file1" required><br><br>
            <label for="file2">File 2:</label>
            <input type="file" name="file2" required><br><br>
            <button type="submit">Process Files</button>
        </form>
    </body>
    </html>
    '''

@app.route('/process', methods=['POST'])
def process_files():
    if 'file1' not in request.files or 'file2' not in request.files:
        return "Please upload both files.", 400

    # Get the uploaded files
    file1 = request.files['file1']
    file2 = request.files['file2']

    # Save the files temporarily
    file1_path = os.path.join(UPLOAD_FOLDER, file1.filename)
    file2_path = os.path.join(UPLOAD_FOLDER, file2.filename)
    file1.save(file1_path)
    file2.save(file2_path)

    try:
        # Process File 1
        df1 = pd.read_excel(file1_path)
        df1.columns = df1.columns.str.strip().str.lower()  # Normalize column names
        df1_processed = df1[['awb number', 'weight']]
        file1_processed_path = os.path.join(PROCESSED_FOLDER, f"processed_{file1.filename}")
        df1_processed.to_excel(file1_processed_path, index=False)
    except Exception as e:
        return f"Error processing File 1: {e}", 400

    try:
        # Process File 2
        df2 = pd.read_excel(file2_path)
        df2.columns = df2.columns.str.strip().str.lower()  # Normalize column names
        df2_processed = df2[['awb number', 'weight']]
        file2_processed_path = os.path.join(PROCESSED_FOLDER, f"processed_{file2.filename}")
        df2_processed.to_excel(file2_processed_path, index=False)
    except Exception as e:
        return f"Error processing File 2: {e}", 400

    # Return download links
    return f'''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Download Processed Files</title>
    </head>
    <body>
        <h1>Download Processed Files</h1>
        <ul>
            <li><a href="/download/{os.path.basename(file1_processed_path)}">Download Processed File 1</a></li>
            <li><a href="/download/{os.path.basename(file2_processed_path)}">Download Processed File 2</a></li>
        </ul>
    </body>
    </html>
    '''


@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(PROCESSED_FOLDER, filename)
    return send_file(file_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
