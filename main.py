import os
from flask import send_from_directory
from flask import Flask, flash, request, redirect, url_for, render_template, send_file
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from excel_convert import convert


app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        file = request.files['file']
        return send_file('result.dat', download_name=convert(file))

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
