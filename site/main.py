"""
Created by: Clement

<Name of project>
<Description of project>
<Use case>
"""

import glob
import os
from datetime import datetime
import shutil

from app import app
from flask import flash, request, redirect, render_template, send_from_directory
from werkzeug.utils import secure_filename

import shopify_to_dhl_format

ALLOWED_EXTENSIONS = {'csv'}
PATH = 'site\\output'


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[
        1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def upload_form():
    return render_template('uploads.html')


@app.route('/', methods=['POST'])
def upload_file():

    upload_path = 'site/uploads/'
    output_path = 'site/output/'

    filelist = [f for f in os.listdir(upload_path) if f.endswith(".csv")]
    for f in filelist:
        os.remove(os.path.join(upload_path, f))

    filelist = [f for f in os.listdir(output_path) if f.endswith(".xlsx")]
    for f in filelist:
        os.remove(os.path.join(output_path, f))

    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No file selected for uploading')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            dt = datetime.now().strftime('%Y-%m-%d_%H%M%S_')
            new_filename = dt + filename
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], new_filename))
            flash('File successfully uploaded as {}'.format(new_filename))
            flash('Please wait while we convert your file')
            run_script(filename)
            return redirect('/getCSV')
        else:
            flash('Allowed file types are csv')
            return redirect(request.url)


@app.route('/output')
def output():
    return '''
            <html><body style="text-align: center;">
            Processing...</br><a href="/getCSV">Click me.</a>
            </body></html>
            '''


@app.route("/getCSV")
def getCSV():
    files = [x for x in os.listdir(PATH) if x.endswith(".xlsx")]
    lst_fullpath = [glob.glob(PATH)[0]+'\\'+filename for filename in files]
    newest = max(lst_fullpath, key=os.path.getmtime)
    newest = newest.replace('site\\output\\', '')
    try:
        return send_from_directory(app.config["OUTPUT_FOLDER"], filename=newest, as_attachment=True)
    except Exception as e:
        print(e)
        return '''
                <html><body style="text-align: center;">
                Something Went Wrong.
                </br><a href="/">Click here to go back to index.</a>
                </body></html>
                '''

    # return '''
    #         <html><body style="text-align: center;">
    #         {}
    #         </br><a href="/">Click here to go back to index.</a>
    #         </body></html>
    #         '''.format(newest)


def run_script(filename):
    shopify_to_dhl_format.main(filename)


if __name__ == "__main__":
    app.run(host= '0.0.0.0')
