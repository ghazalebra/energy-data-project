from flask import Flask, render_template, request, send_file
import os
from utils import Data
import datetime as datetime

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file' not in request.files:
            return render_template('index.html', message='No file part')
        
        file = request.files['file']
        
        # If the user does not select a file, the browser may also
        # submit an empty part without a filename.
        if file.filename == '':
            return render_template('index.html', message='No selected file')
        
        if file:
            start_date = request.form['start_date']
            end_date = request.form['end_date']

            start_time = request.form['start_time_h'] + ':' + request.form['start_time_m'] + ":00"
            end_time = request.form['end_time_h'] + ':' + request.form['end_time_m'] + ":00"
            data = Data(file, start_date, start_time, end_date, end_time)

            data.create_output_tables()
            data.write_output_to_file()
            return render_template('index.html', message='Submitted successfully!', data=data, output_file=data.output_file_name)
    
    return render_template('index.html')

@app.route('/download/<filename>')
def download(filename):
    # Provide download link for the output file
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
