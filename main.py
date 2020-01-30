from flask import Flask, request, render_template
import flask_excel as excel
import openpyxl
from helpers import prediction
import time

app = Flask(__name__)


@app.route("/", methods=['GET', 'POST'])
def main():
    return render_template('main.html')


@app.route("/upload", methods=['GET', 'POST'])
def upload_file():
    subs = {'alg': 'Математика','ru': 'Русский язык','phis': 'Физика','chem':  'Химия',
     'inf': 'Информатика и ИКТ','bio': 'Биология',
     'geo': 'География', 'hist': 'История',
     'social': 'Обществознание',
     'lit': 'Литература','eng': 'Английский язык' }
    s1 = time.time()
    b = openpyxl.load_workbook(request.files['file'])
    print('start', time.time() - s1)
    s = time.time()
    a = prediction(b)
    print('finish', time.time() - s)
    h = []
    a = sorted(a, key = lambda x: x[0])
    for j in a:
        l = []
        b = j[1]
        l.append({'value': j[0], 'color': 2})
        for i in subs:
            r = {'value': ['-', 0.0], 'color': 0}
            for k in b:
                if k[1] == i:
                    k[1] = subs[i]
                    k[0] = int(k[0]*10000)/100
                    if int(k[2]) == 5:
                        r ={'value':  [k[2], k[0]], 'color': 1}
                    elif int(k[2]) == 4 :
                        r = {'value': [k[2], k[0]], 'color': 3}
                    elif int(k[2]) == 2:
                        r = {'value': [k[2], k[0]], 'color': 4}
                    elif int(k[2]) == 3:
                        r = {'value': [k[2], k[0]], 'color': 5}
            l.append(r)
        h.append(l)
    print(h)
    print(time.time() - s1)

    return render_template('result.html', posts = h)

# insert database related code here
if __name__ == "__main__":
    excel.init_excel(app)
    app.run(debug=True)