from flask import Flask, request, redirect, url_for, render_template, send_file
from excel_convert import convert
import time


app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        # while(True):
        #     try:
        #         file = request.files['file']
        #         break
        #     except KeyError:
        #         time.sleep(1)
        file = request.files['file']
        return send_file('result.dat', download_name=convert(file))

    return render_template('index.html')


@app.route('/slsp/purchases/template')
def purchase_template():
    path = 'PURCHASES_TEMPLATE.xlsm'
    return send_file(path, download_name=path)


if __name__ == '__main__':
    app.run(debug=True)
