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

@app.route('/excel-data-extractor', methods=["GET", "POST"])
def extractor():
    if request.method == "POST":

        # Ensure output directory exists
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        ex_excel_file = request.files["file"]
        if not ex_excel_file.filename.endswith('.xlsx'):
            return "Invalid file type", 400
        col_name = request.form["col_name"]
        col_value = request.form["col_value"]
        workbook_name = request.form["ext_workbook_name"]

        # Save the uploaded file
        file_path = os.path.join(UPLOAD_FOLDER, secure_filename(ex_excel_file.filename))
        ex_excel_file.save(file_path)

        # Extract from excel sheet
        df = pd.read_excel(file_path)
        if col_value != "0":
            # Filter rows where the column matches the given value
            extracted_df = df[df[col_name] == col_value]
        elif col_value == "0":
            # Filter and make workbooks for all unique values in column
            # Extract rows for each unique value in the column and save to separate workbooks
            unique_values = df[col_name].dropna().unique()
            zip_filename = os.path.join(OUTPUT_FOLDER, f"{col_name}_extracted.zip")

            with zipfile.ZipFile(zip_filename, 'w') as zipf:
                for value in unique_values:
                    filtered_df = df[df[col_name] == value]
                    output_path = os.path.join(OUTPUT_FOLDER, f'{col_name}_{value}.xlsx')
                    filtered_df.to_excel(output_path, index=False, header=True, engine='openpyxl')
                    zipf.write(output_path, os.path.basename(output_path))
            return send_file(zip_filename, as_attachment=True)
        else:
            # Filter rows where the column is blank
            extracted_df = df[df[col_name].isna()]
        
        # Save the extracted data to a new Excel File
        extracted_file_path = os.path.join(OUTPUT_FOLDER, f"{workbook_name}.xlsx")
        extracted_df.to_excel(extracted_file_path, index=False, header=True, engine='openpyxl')

        return send_file(extracted_file_path, as_attachment=True)
    
    return render_template('excel-data-extractor.html')

@app.route('/excel-column-puller', methods=["GET", "POST"])
def excel_column_puller():
    if request.method == "POST":
        files = request.files.getlist("files")
        col_name = request.form["col_name"]
        new_col_name = request.form["new_col_name"]

        # Create a zip file to store all the new workbooks
        zip_filename = os.path.join(OUTPUT_FOLDER, f"{col_name}_extracted.zip")

        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            # Process each uploaded file
            for file in files:
                file_path = os.path.join(UPLOAD_FOLDER, secure_filename(file.filename))
                file.save(file_path)

                # Read the Excel file and extract the specified column
                df = pd.read_excel(file_path)
                if col_name in df.columns:
                    if new_col_name:
                        extracted_df = df[[col_name]].rename(columns={col_name: new_col_name})
                    else:
                        extracted_df = df[[col_name]]

                    # Create a new workbook name based on the column name and original workbook name
                    original_path_name = secure_filename(file.filename)
                    original_filename = os.path.splitext(original_path_name)
                    original_filename[0].replace('.xlsx', '')
                    print(f"Original filename without extension: {original_filename[0]}")  # Debugging line
                    new_workbook_name = f"{new_col_name}-{original_filename[0]}.xlsx"
                    new_workbook_name = new_workbook_name.strip(",.')")
                    output_path = os.path.join(OUTPUT_FOLDER, new_workbook_name)

                    # Save the extracted column to a new workbook
                    extracted_df.to_excel(output_path, index=False, engine='openpyxl')

                    # Add the new workbook to the zip file
                    zipf.write(output_path, os.path.basename(output_path))
                else:
                    return f"Column '{col_name}' not found in file '{file.filename}'", 400

        return send_file(zip_filename, as_attachment=True)
    
    return render_template('excel-column-puller.html')

if __name__ == '__main__':
    app.run(debug=True)