from flask import Flask, request, send_file
import os
from main import process_schedule  # Import the function from main.py

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_file():
    # Check if a file is uploaded
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    # Save the uploaded file to a temporary location
    input_path = os.path.join("uploads", file.filename)
    file.save(input_path)

    # Call process_schedule from main.py
    workbook = process_schedule(input_path)  # This runs your entire processing, including saving the file

    # Assuming process_schedule saves the file and returns the filename
    output_filename = f"{CompanyName} Schedule {MondayDate2} to {SundayDate2}.xlsx"
    output_path = os.path.join("output", output_filename)

    # Send the output file to the user
    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
