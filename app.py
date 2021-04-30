from flask import Flask, render_template, request, send_from_directory, send_file
from pandas import read_excel, DataFrame
import os
import numpy as np

UPLOAD_FOLDER = 'uploads/'
app = Flask("excel-app")
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route("/")
def homepage():
    return render_template("home.html")


@app.route("/output", methods=["POST"])
def  outputpage():

    if request.method == "POST":
        if request.form.get("openfile"):
            file = request.files["excel_file"]
            filename = file.filename
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            #Reading excel file
            db = read_excel("./"+UPLOAD_FOLDER+"/"+filename)
            #Extracting column names and storing in headings
            headings = db.columns
            headings=headings.to_list()
            data = db.values
            return render_template("table.html", headings=headings, data=data, filename=filename)
        

        elif request.form.get("newrow"):
            filename = request.form.get("filename")
            #Reading excel file
            db = read_excel("./"+UPLOAD_FOLDER+"/"+filename)
            #Extracting column names and storing in headings
            headings = db.columns
            headings=headings.to_list()
            data = db.values
            row = np.array([])

            for field in headings:
                row = np.hstack([row, request.form.get(field)])

            data = np.vstack([data,row])
            newdb = DataFrame(data=data, columns=headings)
            newdb.to_excel("./"+UPLOAD_FOLDER+"/"+filename, sheet_name="Updated", index=False)
            return render_template("table.html", headings=headings, data=data, filename=filename)
       

    
@app.route('/download/<filename>')
def download(filename):
    file_path = UPLOAD_FOLDER + filename
    print("helle"+file_path)
    return send_file(file_path, as_attachment=True, attachment_filename='')    





app.run(port=8080, debug=True)