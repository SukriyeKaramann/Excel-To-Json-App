import os
import pandas as pd
from flask import Flask, request, render_template, redirect, url_for, send_file, flash, session
import json

app = Flask(__name__)
app.secret_key = 'excel2json_secret'
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('excel_file')
        if not file:
            flash('Lütfen bir dosya seçin.')
            return redirect(request.url)

        if not file.filename.endswith(('.xlsx', '.xls')):
            flash('Yalnızca .xlsx veya .xls uzantılı dosyalar kabul edilir.')
            return redirect(request.url)

        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)

        try:
            df = pd.read_excel(filepath)
        except Exception as e:
            flash(f'Dosya okunamadı: {str(e)}')
            return redirect(request.url)

        json_data = df.to_dict(orient='records')
        preview_data = json_data[:10]  # İlk 10 satır önizleme
        json_filename = file.filename.rsplit('.', 1)[0] + '.json'
        json_path = os.path.join(app.config['OUTPUT_FOLDER'], json_filename)

        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, ensure_ascii=False, indent=4)

        # Geçici session ile path'i gönder
        session['json_filename'] = json_filename
        session['preview_data'] = preview_data
        return redirect(url_for('result'))

    return render_template('index.html')


@app.route('/result')
def result():
    json_filename = session.get('json_filename')
    preview_data = session.get('preview_data')

    if not json_filename or not preview_data:
        flash("Veri bulunamadı.")
        return redirect(url_for('index'))

    # JSON verisini düzgün bir şekilde string olarak dönüştür
    preview_data_json = json.dumps(preview_data, ensure_ascii=False, indent=4)

    return render_template('result.html', preview_data=preview_data_json, json_file=json_filename)


@app.route('/download/<filename>')
def download(filename):
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    else:
        flash('Dosya bulunamadı.')
        return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)
