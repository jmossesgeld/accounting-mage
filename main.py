from flask import Flask, request, redirect, url_for, render_template, send_file
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
 