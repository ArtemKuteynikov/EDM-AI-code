y = []
import openpyxl
def to_dicts(wb, sheet_name):
    mid, num, exam, sub, class_nom, fin, absent, num_5, num_4, num_3, num_2 = execute(wb, sheet_name)
    d1 = dict()
    for i in mid:
        d1.update({i: [mid[i], num[i], fin[i], absent[i], num_5[i], num_4[i], num_3[i], num_2[i], exam[i]]})
    return d1, sub,  class_nom
from helpers import get_sub, is_examing, execute
subs = {'Алгебра': 'alg', 'Геометрия': 'geom', 'Русский язык': 'ru', 'Физика': 'phis', 'Химия': 'chem',
        'Информатика и ИКТ': 'inf', 'Биология': 'bio',
        'География': 'geo', 'История': 'hist', 'Обществознание': 'social',
        'Литература': 'lit', 'Английский язык': 'eng'}
res = dict()
work_books_9_19 = [openpyxl.load_workbook('journals-9-19-(1).xlsx')]
for i in subs:
    res.update({subs[i]: dict()})
for wb in work_books_9_19:
    sheets = wb.get_sheet_names()
    for j in sheets:
        a = get_sub(j, wb)[1]
        if a!=0:
            if is_examing(a, subs)[0]:
                d, sub, _ = to_dicts(wb, j)
                print(is_examing(a, subs)[1], sub)
                res[is_examing(a, subs)[1]].update(d)
exam = ['exam_phis', 'exam_chem', 'exam_inf', 'exam_bio', 'exam_geo', 'exam_hist', 'exam_social', 'exam_lit',
        'exam_eng']
d = dict()
for i in res['alg']:
    d.update({i: list()})

for i in res:
    for j in res[i]:
        print(res[i][j][-1], i, j)
        if res[i][j][-1] != False:
            a = d[j]
            print(res[i][j], i, j)
            a.append(str(i)+' '+str(res[i][j][-1]))
            d.update({j: a})
a = []
h = []
for i in d:
    a.append([i, d[i]])

a = sorted(a, key = lambda x: x[0])
print(a)
book = openpyxl.Workbook()
sheet = book.active
rows = []
for i in a:
    b = []
    b.append(i[0])
    b.extend(i[1])
    rows.append(b)

for row in rows:
    sheet.append(row)
    print(row)

book.save('appending.xlsx')