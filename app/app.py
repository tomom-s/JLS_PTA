import os
from flask import Flask, flash, redirect, render_template, request, redirect, url_for, send_from_directory, jsonify, send_file
from io import BytesIO
from glob import glob
#import io
from zipfile import ZipFile
#import openpyxl
from openpyxl.styles import PatternFill, Font
#import csv
#import sys
import pandas as pd
#import datetime
#pip freeze > requirements.txt
#pip install -r requirements.txt
# Configure application
app = Flask(__name__)
app.debug = os.environ.get('FLASK_DEBUG', False)

# Ensure templates are auto-reloaded
app.config["TEMPLATES_AUTO_RELOAD"] = True

DOWNLOAD_FOLDER = 'static/downloads/'
UPLOAD_FOLDER = 'static/uploads/'

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

if not os.path.exists(DOWNLOAD_FOLDER):
    os.makedirs(DOWNLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

APP_ROOT = os.path.dirname(os.path.abspath(__file__))

ALLOWED_EXTENSIONS = set(['csv'])

TASKS = {1:"台帳作成",2:"台帳から名簿の作成"}


def allowed_file(filename):
	return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET", "POST"])
def index():
#Show tasks to perform
    print(TASKS)
    if request.method == "POST":
        id = int(request.form.get("tasks"))

        return render_template('task.html', id=id, desc=TASKS[id])
    else:
        return render_template("index.html", tasks=TASKS)

@app.route('/task', methods=['GET', 'POST'])
def upload():
    if request.method == "POST":
        try:
            id = int(request.form.get('id'))
        except TypeError:
            message = 'タスクを選択してください'
            return render_template('apology.html', message=message)


        print(id)
        files = request.files.getlist('file[]')
        print(files)
        target = os.path.join(APP_ROOT)
        print(target)
        print(os.path.join(app.config['UPLOAD_FOLDER']))

        file_set = []

        for file in files:
            if allowed_file(file.filename):
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], file.filename))
            else:
                message = 'Allowed files are CSV only. Please try again.'
                for f in os.listdir(UPLOAD_FOLDER):
                    path = os.path.join(UPLOAD_FOLDER, f)
                    if os.path.isfile(path):
                        file_set.append(path)

                for f in file_set:
                    os.remove(f)
                return render_template('apology.html', message=message)

        for f in os.listdir(UPLOAD_FOLDER):
            path = os.path.join(UPLOAD_FOLDER, f)
            if os.path.isfile(path):
                file_set.append(path)
        try:
            if id == 1:
                flist = master_creation(file_set)

            elif id == 2:
                flist = split_class(file_set)


            else:
                message = "Something went wrong. Please go back and try again."
                return render_template('apology.html', message=message)


            print(flist)
            for f in file_set:
                os.remove(f)
            return render_template('download.html', fileslist=flist)
        except Exception as e:

            message = e
#            message = "The files uploaded and tasks selected do not match. Please try again. "
            return render_template('apology.html', message=message)





@app.route("/download", methods=['GET', 'POST'])
def download_file():
    target = DOWNLOAD_FOLDER

    stream = BytesIO()
    with ZipFile(stream, 'w') as zf:
        for file in glob(os.path.join(target, '*.*')):
            zf.write(file)
            os.remove(file)
            print('ダウンロード準備が整いました')
    stream.seek(0)
    return send_file(stream, as_attachment=True, download_name='archive.zip')

def master_creation(file_set):
    xlsx_list = []
    for f in file_set:
        out = f.replace(UPLOAD_FOLDER, DOWNLOAD_FOLDER)
        out = out.replace('.csv', '.xlsx')

    if len(file_set) > 1:
        message = "ファイルは複数選択しないでください。やり直してください。"
        return render_template('apology.html', message=message)

#       print('Multiple csv files in folder. Please keep one only')

    else:
        excel_name = file_set[0].replace('.csv', '.xlsx')
        excel_name = excel_name.replace(UPLOAD_FOLDER, DOWNLOAD_FOLDER)
        df = pd.read_csv(file_set[0], skiprows=1)
        print(df.columns.to_list())
        #df.drop(df.index[df['連番'] == '連番'], inplace=True)
        #df = df.dropna(subset=['連番'])
        df.drop(df.index[df['役員'] == '役員'], inplace=True)
        df = df.dropna(subset=['役員'])

#        print(df)
        df = df.replace('幼稚部2年', '幼稚部年長', regex=True)
        df = df.replace('幼稚部1年', '幼稚部年中', regex=True)
        df = df.replace('Japanese Division1年Japanese1組', '日本語部1', regex=True)
        df = df.replace('Japanese Division2年Japanese2組', '日本語部2', regex=True)
        df = df.replace('Japanese Division3年Japanese3組', '日本語部3', regex=True)
        df = df.replace('Japanese Division4年Japanese4組', '日本語部4', regex=True)
        df = df.replace('Japanese Division5年Japanese5組', '日本語部5', regex=True)
        df = df.replace('高等部1年1組', '高等部1年', regex=True)
        df = df.replace('高等部2年1組', '高等部2年', regex=True)
        df = df.replace('配付', '長子', regex=True)

        df.to_excel(excel_name, columns=['生徒番号','配付','学年組','生徒漢字名', '生徒ローマ字名', '性別', '兄弟姉妹のクラス',
                                                        '兄弟姉妹名', '保護者１漢字名', '保護者１電話','保護者１email',
                                                        '保護者２漢字名', '保護者２email'], index=False)
        xlsx_list.append(excel_name)
        return xlsx_list

def split_class(file_set):
    xlsx_list = []

    if len(file_set) > 1:
        message = "ファイルは複数選択しないでください。やり直してください。"
        return render_template('apology.html', message=message)

    else:
        df = pd.read_csv(file_set[0])
        df = df.rename(columns={'保護者１漢字名':'保護者',
                                              '保護者１電話':'保護者電話',
                                              '保護者１email':'保護者email'})
        print(df)
        out = file_set[0].replace(UPLOAD_FOLDER, DOWNLOAD_FOLDER)
        out = os.path.dirname(out)

        print(out)
        for p in df['学年組'].unique():
            df.loc[df['学年組'] == p].to_excel(f'{out}/{p}.xlsx',columns=['生徒漢字名','生徒ローマ字名','性別','兄弟姉妹のクラス','兄弟姉妹名','保護者','保護者電話','保護者email','免除/減免'], index=False)
            print(f'{file_set}/{p}.xlsx')
            xlsx_list.append(f'{out}/{p}.xlsx')
        return xlsx_list

if __name__ == '__main__':
    app.run(host='0.0.0.0')
