from flask import Flask, request, redirect, url_for, render_template, send_file, after_this_request
from excel_convert import Converter
from image_watermarking import apply_watermark

import os
import zipfile
import io
import pathlib
from shutil import rmtree

app = Flask(__name__)


@app.route('/slsp-convert', methods=['GET', 'POST'])
def slsp_convert():
    path = ""
    if request.method == 'POST':
        file = request.files['file']
        converter = Converter(file)
        path = converter.slsp()
        if converter.has_error == False:
            return send_file('result.DAT', as_attachment=True, attachment_filename=path)
    return render_template('projects/excel-convert.html', BIR_form='2550Q SLSP', version='RELIEF version: 2.3', template='PURCHASES_TEMPLATE.xlsm', error=path)


@app.route('/qap-convert', methods=['GET', 'POST'])
def qap_convert():
    path = ""
    if request.method == 'POST':
        file = request.files['file']
        converter = Converter(file)
        path = converter.qap()
        if converter.has_error == False:
            return send_file('result.DAT', as_attachment=True, attachment_filename=path)
    return render_template('projects/excel-convert.html', BIR_form='1601EQ QAP', version='Alphalist version: 7.0', template='QAP_TEMPLATE.xlsm', error=path)


@app.route('/download-template/<path>')
def download_template(path):
    return send_file(f"excel_templates/{path}", as_attachment=True, attachment_filename=path)


@app.route('/tools', methods=['GET', 'POST'])
def tools():
    return render_template('tools.html')


@app.route('/projects/<project>', methods=['GET', 'POST'])
def projects(project):
    if request.method == 'POST':
        try:
            files = request.files.getlist('files[]')
            try:
                rmtree('UPLOADED_FILES')
            except:
                print("Folder not found.")
            try:
                os.makedirs('UPLOADED_FILES')
            except:
                print("Folder already exists.")
            for file in files:
                path = os.path.join('UPLOADED_FILES/', file.filename)
                file.save(path)
                apply_watermark(path, 'watermark.png')

            base_path = pathlib.Path('UPLOADED_FILES')
            data = io.BytesIO()
            with zipfile.ZipFile(data, mode='w') as z:
                for f_name in base_path.iterdir():
                    z.write(f_name)
            data.seek(0)

            return send_file(data, mimetype='application/zip', as_attachment=True, attachment_filename='data.zip')

        except Exception as e:
            return render_template('error.html', error=e)

    return render_template(f'projects/{project}.html', project_name=project)


@app.route('/')
def home():
    return render_template('index.html')


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)
