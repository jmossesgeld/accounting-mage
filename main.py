from flask import Flask, request, redirect, url_for, render_template, send_file
from excel_convert import Converter

app = Flask(__name__)


@app.route('/slsp-convert', methods=['GET', 'POST'])
def slsp_convert():
    path = ""
    if request.method == 'POST':
        file = request.files['file']
        converter = Converter(file)
        path = converter.slsp()
        if path[:5] != '<stro':
            return send_file('result.DAT', as_attachment=True, attachment_filename=path)
    return render_template('excel-convert.html', BIR_form='2550Q SLSP', version='RELIEF version: 2.3', template='PURCHASES_TEMPLATE.xlsm', error=path)


@app.route('/qap-convert', methods=['GET', 'POST'])
def qap_convert():
    path = ""
    if request.method == 'POST':
        file = request.files['file']
        converter = Converter(file)
        path = converter.qap()
        if path[:5] != '<stro':
            return send_file('result.DAT', as_attachment=True, attachment_filename=path)
    return render_template('excel-convert.html', BIR_form='1601EQ QAP', version='Alphalist version: 7.0', template='QAP_TEMPLATE.xlsm', error=path)


@app.route('/download-template/<path>')
def download_template(path):
    return send_file(path, as_attachment=True, attachment_filename=path)


@app.route('/tools', methods=['GET', 'POST'])
def tools():
    return render_template('tools.html')


@app.route('/')
def home():
    return render_template('index.html')


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)
