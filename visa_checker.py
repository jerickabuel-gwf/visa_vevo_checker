from flask import Flask, request, render_template, redirect, url_for
import pandas as pd
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        data = pd.read_csv(filepath)
        analysis = data.describe().to_html()
        return render_template('analysis.html', tables=[analysis], titles=[''])
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host = '0.0.0.0', debug=True)
