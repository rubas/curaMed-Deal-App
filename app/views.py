from flask import *
import requests
import os

from app import app
from app.hubspot import *

select = None 

app_data = {
    "name":         "Hubspot Deal App",
    "description":  "Hubspot Deal App to import curaMed data",
    "author":       "ProfileMedia",
    "html_title":   "curaMed Deal App",
    "project_name": "curaMed Deal App",
    "keywords":     "flask, webapp, Hubspot, curaMed"
}

ALLOWED_EXTENSIONS = {'xlsx'} 

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        input_file = request.files['file']
        if input_file.filename == '':
            return render_template('home.html', app_data=app_data)
        if input_file and allowed_file(input_file.filename):
            input_file.save(os.path.join('upload'))
            return redirect(url_for('select'))
            #return render_template('file_uploaded.html', app_data=app_data)
    else:
        return render_template('home.html', app_data=app_data)


@app.route('/select',methods=['GET','POST'])
def select():
    global select
    if request.method == 'POST':
        select = request.form.get('select') 
        return render_template("upload.html",app_data= app_data)
        #if select == "Ärzte Import":
        #    pass
        #elif select == "Kliniken/Gruppenpraxen":
        #    pass
        #elif select == "Zuweisungsart":
        #    #test(path="upload")
        #    #return redirect(url_for('upload'))
        #    return render_template("upload.html",app_data= app_data)
    return render_template('select.html', app_data=app_data)

@app.route('/upload',methods=['GET','POST'])
def upload():
    #return render_template("upload.html",app_data= app_data)
    if select == "Ärzte Import":
        return Response(aerzte_import(file_path="upload"), mimetype= 'text/event-stream')
    elif select == "Kliniken/Gruppenpraxen":
        return Response(gruppenpraxen_import(file_path="upload"), mimetype= 'text/event-stream')
    elif select == "Zuweisungsart": 
        return Response(zuweisungen(file_path="upload"), mimetype= 'text/event-stream')

    