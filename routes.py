from flask import Blueprint, render_template, request, send_file
from utils import run_script
import os

from flask import current_app
main = Blueprint('main', __name__)

@main.route('/')
def index():
    return render_template('index.html')

@main.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    file_path = os.path.join(current_app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)
    output_file = run_script(file_path)
    return send_file(output_file, as_attachment=True)
