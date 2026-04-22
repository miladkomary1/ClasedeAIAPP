import os
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename
from run_takeoff import run_takeoff

app = Flask(__name__)

# Config for Vercel (Read-only filesystem requires /tmp)
UPLOAD_FOLDER = '/tmp/uploads'
OUTPUT_FOLDER = '/tmp/outputs'
ALLOWED_EXTENSIONS = {'dxf'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
        
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            # Process the file
            out_path, rows = run_takeoff(filepath, OUTPUT_FOLDER)
            
            if out_path is None or not os.path.exists(out_path):
                return jsonify({'error': 'No rebar annotations found in the provided DXF file.'}), 400
                
            # Send info and file download link to user
            output_filename = os.path.basename(out_path)
            return jsonify({
                'success': True,
                'download_url': f'/api/download/{output_filename}',
                'rows': rows
            })
            
        except Exception as e:
            return jsonify({'error': str(e)}), 500
            
    return jsonify({'error': 'Invalid file type. Only .dxf files are allowed.'}), 400

@app.route('/api/download/<filename>')
def download_file(filename):
    safe_name = secure_filename(filename)
    return send_file(os.path.join(OUTPUT_FOLDER, safe_name), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
