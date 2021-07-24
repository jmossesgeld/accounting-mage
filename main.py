from os import error
from sys import exc_info
import traceback
from types import TracebackType
from werkzeug.exceptions import HTTPException
from flask import json
from flask import Flask, request, redirect, url_for, render_template, send_file
from excel_convert import convert


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
        path = convert(file)
        print(path[:5])
        if path[:5] == '<stro':
            return render_template('index.html',error=path)
        else:
            return send_file('result.DAT', as_attachment=True, attachment_filename=path)
    return render_template('index.html')


@app.route('/slsp/purchases/template')
def purchase_template():
    path = 'PURCHASES_TEMPLATE.xlsm'
    return send_file(path, download_name=path)


if __name__ == '__main__':
    app.run()
