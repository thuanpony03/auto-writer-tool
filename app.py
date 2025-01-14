from flask import Flask, render_template, request, send_file, flash, jsonify
import os
from werkzeug.utils import secure_filename
import zipfile
from io import BytesIO
from main import fill_word_template_with_table

app = Flask(__name__)
app.secret_key = os.urandom(24)  # Cần thiết để sử dụng flash messages

# Configure upload folder
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output_docs'
ALLOWED_EXTENSIONS = {'xlsx', 'docx'}

# Create directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_documents():
    if 'excel_file' not in request.files or 'word_template' not in request.files:
        return jsonify({'status': 'error', 'message': 'Vui lòng tải lên cả file Excel và Word template'})

    excel_file = request.files['excel_file']
    word_template = request.files['word_template']

    if excel_file.filename == '' or word_template.filename == '':
        return jsonify({'status': 'error', 'message': 'Chưa chọn file'})

    if not (allowed_file(excel_file.filename) and allowed_file(word_template.filename)):
        return jsonify({'status': 'error', 'message': 'Định dạng file không hợp lệ'})

    try:
        # Save uploaded files
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(excel_file.filename))
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(word_template.filename))
        
        excel_file.save(excel_path)
        word_template.save(template_path)

        # Generate documents using your existing function
        fill_word_template_with_table(excel_path, template_path, OUTPUT_FOLDER)

        # Create zip file in memory
        memory_file = BytesIO()
        with zipfile.ZipFile(memory_file, 'w') as zf:
            for root, _, files in os.walk(OUTPUT_FOLDER):
                for file in files:
                    file_path = os.path.join(root, file)
                    zf.write(file_path, os.path.relpath(file_path, OUTPUT_FOLDER))

        #Delete all files in output folder
        for root, _, files in os.walk(OUTPUT_FOLDER):
            for file in files:
                file_path = os.path.join(root, file)
                os.remove(file_path)

        memory_file.seek(0)
        
        # Clean up uploaded files
        os.remove(excel_path)
        os.remove(template_path)

        return send_file(
            memory_file,
            mimetype='application/zip',
            as_attachment=True,
            download_name='generated_documents.zip'
        )

    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Có lỗi xảy ra: {str(e)}'})

if __name__ == '__main__':
    app.run(debug=True)