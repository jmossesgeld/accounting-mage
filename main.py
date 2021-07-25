from flask import Flask, request, redirect, url_for, render_template, send_file
from excel_convert import convert


app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def tools():
    if request.method == 'POST':
        file = request.files['file']
        path = convert(file)
        print(path[:5])
        if path[:5] == '<stro':
            return render_template('tools.html',error=path)
        else:
            return send_file('result.DAT', as_attachment=True, attachment_filename=path)
    return render_template('tools.html')


@app.route('/slsp/purchases/template')
def purchase_template():
    path = 'PURCHASES_TEMPLATE.xlsm'
    return send_file(path, download_name=path)

@app.route('/hero')
def home():
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
