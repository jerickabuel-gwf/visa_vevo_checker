from flask import Flask, render_template, request, send_file, after_this_request
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


@app.route('/', methods=['GET', 'POST'])
def index():
  if request.method == 'POST':
    file = request.files['file']
    if file:
      file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
      file.save(file_path)
      df = pd.read_csv(file_path)
      # Process the file using pandas
      processed_df = df  # Add your processing logic here
      processed_file_path = os.path.join(app.config['UPLOAD_FOLDER'],
                                         'processed_' + file.filename)
      processed_df.to_csv(processed_file_path, index=False)
      return render_template('index.html',
                             filename='processed_' + file.filename)
  return render_template('index.html')


@app.route('/download/<filename>')
def download_file(filename):
  file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)

  @after_this_request
  def remove_file(response):
    try:
      os.remove(file_path)
      original_file_path = file_path.replace('processed_', '')
      if os.path.exists(original_file_path):
        os.remove(original_file_path)
    except Exception as error:
      app.logger.error("Error removing or closing downloaded file handle",
                       error)
    return response

  return send_file(file_path, as_attachment=True)


if __name__ == '__main__':
  if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
  app.run(host="0.0.0.0", port=8080, debug=True)
