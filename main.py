from flask import Flask, jsonify, request, redirect, url_for, render_template, send_file, after_this_request
from projects.excel_convert.excel_convert import Converter
from projects.image_watermarking.image_watermarking import convert_images
from projects.pcg_gas.semaphore import send_message
import json

app = Flask(__name__)


@app.route('/slsp-convert', methods=['GET', 'POST'])
def slsp_convert():
    path = ""
    if request.method == 'POST':
        file = request.files['file']
        converter = Converter(file)
        path = converter.slsp()
        if converter.has_error == False:
            return send_file('projects/excel_convert/result.DAT', as_attachment=True, attachment_filename=path)
    return render_template('projects/excel-convert.html', BIR_form='2550Q SLSP', version='RELIEF version: 2.3', template='PURCHASES_TEMPLATE.xlsm', error=path)


@app.route('/qap-convert', methods=['GET', 'POST'])
def qap_convert():
    path = ""
    if request.method == 'POST':
        file = request.files['file']
        converter = Converter(file)
        path = converter.qap()
        if converter.has_error == False:
            return send_file('projects/excel_convert/result.DAT', as_attachment=True, attachment_filename=path)
    return render_template('projects/excel-convert.html', BIR_form='1601EQ QAP', version='Alphalist version: 7.0', template='QAP_TEMPLATE.xlsm', error=path)


@app.route('/compensation_convert', methods=['GET', 'POST'])
def compensation_convert():
    path = ""
    if request.method == 'POST':
        file = request.files['file']
        converter = Converter(file)
        path = converter.qap()
        if converter.has_error == False:
            return send_file('projects/excel_convert/result.DAT', as_attachment=True, attachment_filename=path)
    return render_template('projects/excel-convert.html', BIR_form='1604C', version='Alphalist version: 7.0', template='1604C_TEMPLATE.xlsm', error=path)


@app.route('/download-template/<path>')
def download_template(path):
    return send_file(f"excel_templates/{path}", as_attachment=True, attachment_filename=path)


@app.route('/tools', methods=['GET', 'POST'])
def tools():
    return render_template('tools.html')


@app.route('/projects/image-watermarking', methods=['GET', 'POST'])
def image_watermarking():
    if request.method == 'POST':
        try:
            files = request.files.getlist('files[]')
            data = convert_images(files)
            return send_file(data, mimetype='application/zip', as_attachment=True, attachment_filename='data.zip')
        except Exception as e:
            return render_template('error.html', error=e)
    return render_template(f'projects/image-watermarking.html', project_name='image-watermarking')


@app.route('/projects/pcg-gas', methods=['GET', 'POST'])
def pcg_message():
    if request.method == 'POST':
        try:
            data = jsonify(request.get_json())
            message = data['message']
            number = data['number']
            return send_message(message, number).json()
        except Exception as e:
            return jsonify(e)

@app.route('/')
def home():
    return render_template('index.html')


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)
