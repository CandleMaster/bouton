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
        translation_option = request.form["translation_option"]
        url = request.form["url"]
        # zipfile_ob = zipfile.ZipFile(file_like_object)
        # file_names = zipfile_ob.namelist()
        # if uploaded_file.filename != '':
        #     uploaded_file.save(uploaded_file.filename)
        #     filename = secure_filename(uploaded_file.filename)
        #     input_data=uploaded_file.save(os.path.join(app.config['UPLOAD_PATH'], filename))
        response = extract_to_doc(file_like_object,url,translation_option)
        return response
    return render_template('index.html')

            # return redirect(url_for(hello_world))

        # if filename != '':
        #     file_ext = os.path.splitext(filename)[1]
        #     if file_ext not in app.config['UPLOAD_EXTENSIONS']:
        #         abort(400)
        #     f.save(os.path.join(app.config['UPLOAD_PATH'], filename))



if __name__=="__main__":
    app.run(debug=True)