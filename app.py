from flask import Flask, render_template, request, send_file
from processing import extract_to_doc
import zipfile
import os
# from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024
app.config['UPLOAD_EXTENSIONS'] = ['.zip']
app.config['UPLOAD_PATH'] = 'uploads'


@app.route('/', methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files['file']
        file_like_object = file.stream._file  
        # translation_option = request.form["translation_option"]
        url = request.form["url"]
        response = extract_to_doc(file_like_object,url)
        return response
    return render_template('index.html')



if __name__=="__main__":
    app.run(debug=True)