from flask import Flask, render_template, request, send_file
import os
from utils import Data

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    return "Hi!"
    # if request.method == 'POST':
    #     # Check if the post request has the file part
    #     if 'file' not in request.files:
    #         return render_template('index.html', message='No file part')
        
    #     file = request.files['file']
        
    #     # If the user does not select a file, the browser may also
    #     # submit an empty part without a filename.
    #     if file.filename == '':
    #         return render_template('index.html', message='No selected file')
        
    #     if file:
    #         start_date = request.form['start_date']
    #         end_date = request.form['end_date']
    #         if start_date and end_date:
    #             data = Data(file, start_date, end_date)
    #         else:
    #             data = Data(file)
    #         data.create_output_tables()
    #         data.write_output_to_file()
    #         return render_template('index.html', message='Submitted successfully!', data=data, output_file=data.output_file_name)
    
    # return render_template('index.html')

# @app.route('/download/<filename>')
# def download(filename):
#     # Provide download link for the output file
#     return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
