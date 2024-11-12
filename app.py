from flask import Flask, render_template, request, send_file, redirect, flash
import time
import os
import pandas as pd
import shutil
from server_monitor import process_servers

app = Flask(__name__)

app.secret_key = os.environ.get('SECRET_KEY')
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['DOWNLOAD_FOLDER'] = 'downloads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB upload limit



if os.path.exists(app.config['UPLOAD_FOLDER']):
    shutil.rmtree(app.config['UPLOAD_FOLDER'])
if os.path.exists(app.config['DOWNLOAD_FOLDER']):
    shutil.rmtree(app.config['DOWNLOAD_FOLDER'])


os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part in the request.')
            return redirect(request.url)
        
        file = request.files['file']
        
        if file.filename == '':
            flash('No file selected for uploading.')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            filename = 'uploaded_input.xlsx'
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            try:
                servers_df = pd.read_excel(filepath)
            except Exception as e:
                flash(f'Error reading Excel file: {e}')
                return redirect(request.url)

            template_path = 'server_template.xlsx'
            output_filename = f'Server_Capacity_{time.strftime("%Y-%m-%d")}.xlsx'
            output_filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)

            messages = process_servers(servers_df, template_path, output_filepath)
            all_failed = "Failed to connect to all servers. No file generated." in messages

            for msg in messages:
                flash(msg)

            if not all_failed and os.path.exists(output_filepath):
                return render_template('index.html', download_filename=output_filename)
            else:
                if os.path.exists(output_filepath):
                    os.remove(output_filepath)
                return redirect(request.url)
        else:
            flash('Allowed file types are .xlsx and .xls')
            return redirect(request.url)
    
    return render_template('index.html')


@app.route('/download/<filename>')
def download_file(filename):
    return send_file(
        os.path.join(app.config['DOWNLOAD_FOLDER'], filename),
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    app.run(debug=True)

