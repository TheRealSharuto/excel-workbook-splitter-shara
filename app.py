from flask import Flask
from flask import render_template, request, send_file
import pandas as pd
import os
from werkzeug.utils import secure_filename
import zipfile

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads/'
OUTPUT_FOLDER = 'output/'

# Ensure upload and output directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        # Get the file from the form
        excel_file = request.files["file"]
        rows_per_sheet = int(request.form["rows"])
        sheet_name = request.form["sheet_name"]

        # Save the uploaded file
        file_path = os.path.join(UPLOAD_FOLDER, secure_filename(excel_file.filename))
        excel_file.save(file_path)

        # Split the Excel sheet
        df = pd.read_excel(file_path)
        sheets = [df[i:i + rows_per_sheet] for i in range(0, df.shape[0], rows_per_sheet)]
        zip_filename = os.path.join(OUTPUT_FOLDER, f"{sheet_name}.zip")

        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for i, sheet in enumerate(sheets):
                output_path = os.path.join(OUTPUT_FOLDER, f'{sheet_name}{i+1}.xlsx')
                sheet.to_excel(output_path, index=False, header=True)
                zipf.write(output_path, os.path.basename(output_path))

        return send_file(zip_filename, as_attachment=True)
    
    return render_template("index.html")

if __name__ == '__main__':
    app.run()