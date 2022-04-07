from tkinter import *
import os
import pandas as pd
from datetime import datetime
from textwrap import wrap
import sys


def read_excel(path_file):
    questions = pd.read_excel(path_file, index_col=0, dtype={
        'Name': str, 'Weight': int})
    return questions


a = 0

file = read_excel('questions.xlsx')
values = file.values
questions = []
answers = []
wrong_answers = []
files = []


def create_file_name(time):
    dt_string = time.strftime("%d-%m-%Y (%H-%M-%S)")
    return dt_string


for i in range(0, len(values)):
    questions.append((values[i][0], values[i][1]))


def create_question_list():
    global a
    if a < len(questions):
        text = str(a + 1) + "/" + str(len(questions)) + questions[a][0]
        if len(text) > 150:
            text = '\n'.join(wrap(text, 150))
        label['text'] = text

        button1.config(text='Да')
        button2.config(text='Нет')

        button1.config(command=bt1)
        button2.config(command=bt2)
        a = a + 1
    else:
        write_in_file()


def create_question_list_qual():
    global a
    if a < len(questions):
        text = str(a + 1) + "/" + str(len(questions)) + questions[a][0]
        if len(text) > 150:
            text = '\n'.join(wrap(text, 150))
        label['text'] = text

        button1.config(text='Да')
        button2.config(text='Нет')

        button1.config(command=bt11)
        button2.config(command=bt22)
        a = a + 1
    else:
        write_in_file_qual()


def bt1():
    answers.append([questions[a - 1][0], questions[a - 1][1], "+"])
    create_question_list()


def bt2():
    answers.append([questions[a - 1][0], questions[a - 1][1], "-"])
    create_question_list()


def bt11():
    answers.append((questions[a - 1][0], "+"))
    create_question_list_qual()


def bt22():
    answers.append((questions[a - 1][0], "-"))
    create_question_list_qual()


def write_in_file():
    path = os.getcwd() + r"/data/quan/"
    file_name = create_file_name(datetime.now())
    quan_file = open(path + str(file_name + '.txt'), 'w+', encoding='utf-8')
    for x in answers:
        quan_file.write(str(x[0]) + ' ' + str(x[1]) + ' ' + str(x[2]) + '\n')
    risk = result(answers)
    for x in risk:
        quan_file.write(str(x))
        quan_file.write('\n')
    quan_file.write('\n')
    quan_file.close()
    get_residual_risk(risk)


def get_residual_risk(risk):
    label['text'] = 'Вы прошли тестирование'
    button1.config(text='Сравнить с предыдущим тестированием')
    button2.config(text='Выйти')

    button1.config(command=verify)
    button2.config(command=exit)


def verify():
    actual_risk = result(answers)
    file_q = get_last_quan_file()
    # show_selected_file(file_q)
    prev_persent = get_prev_risk(file_q)

    actual_persent = int(actual_risk[1].split(':')[1].replace(' ', '').replace('%', ''))
    complete_residual_risk(prev_persent, actual_persent)


def show_wrong_answers():
    T.insert(END, 'Моменты, на которые стоит обратить внимание: ')
    T.insert(END, '\n')
    for x in wrong_answers:
        T.insert(END, x)
        T.insert(END, '\n')


def complete_residual_risk(prev, actual):
    T.delete(0.0, END)
    if prev == actual:
        T.insert(END, '''Риск невыполнения требований не изменился.\nОн так же составляет: ''' + str(
            actual) + '%, предыдущий риск: ' + str(prev) + '%')
        T.insert(END, '\n')
        show_wrong_answers()
    elif prev > actual:
        T.insert(END, '''Риск невыполнения требований стал меньше.\nОн составляет: ''' + str(
            actual) + '%, предыдущий риск: ' + str(prev) + '%')
        T.insert(END, '\n')
        show_wrong_answers()
    elif prev < actual:
        T.insert(END, '''Риск невыполнения требований стал больше.\nОн составляет: ''' + str(
            actual) + '%, предыдущий риск: ' + str(prev) + '%')
        T.insert(END, '\n')
        show_wrong_answers()


def exit():
    sys.exit(0)


def complete_risk(max, company):
    result = ['Максимальный риск невыполнения требований ISO 17799: ' + str(max),
              'Риск невыполнения требований ISO 17799 в компании: ' + str(round((max - company) / max * 100)) + '%']

    T.insert(END, 'Максимальный риск невыполнения требований ISO 17799: ' + str(max))
    T.insert(END, '\n')
    T.insert(END,
             'Риск невыполнения требований ISO 17799 в компании: ' + str(round((max - company) / max * 100)) + '%')
    T.insert(END, '\n')
    return result


# подсчёт результатов для количественной оценки
def result(answers):
    r_max = 0
    r_company = 0
    for answer in answers:
        r_max += answer[1]
        if answer[2] == '+':
            r_company += answer[1]
        else:
            wrong_answers.append(answer[0])
    return complete_risk(r_max, r_company)


def read_excel(path_file):
    questions = pd.read_excel(path_file, index_col=0, dtype={
        'Name': str, 'Weight': int})
    return questions


def complete_test():
    label.config(text='Выберите тип оценки:')
    button1.config(text='Качественная оценка')
    button2.config(text='Количественная оценка')
    button1.config(command=create_question_list)
    button2.config(command=create_question_list_qual)


def get_last_qual_file():
    path = os.getcwd() + r"/data/qual/"
    files = os.listdir(path)
    files = [os.path.join(path, file_q) for file_q in files]
    files = [file_q for file_q in files if os.path.isfile(file_q)]
    files = sorted(files, key=os.path.getctime)
    files.pop()
    return max(files, key=os.path.getctime)


def get_last_quan_file():
    path = os.getcwd() + r"/data/quan/"
    files = os.listdir(path)
    files = [os.path.join(path, file_q) for file_q in files]
    files = [file_q for file_q in files if os.path.isfile(file_q)]
    files = sorted(files, key=os.path.getctime)
    files.pop()
    return max(files, key=os.path.getctime)


def show_selected_file(file_q):
    file_q = open(file_q, 'r', encoding='utf-8')
    lines = file_q.readlines()
    file_q.close


def get_prev_risk(file_q):
    f = open(file_q, 'r', encoding='utf-8')
    lines = f.readlines()
    return int(lines[len(lines) - 2].split(':')[1].replace(' ', '').replace('%', ''))


def write_in_file_qual():
    path = os.getcwd() + r"/data/qual/"
    file_name = create_file_name(datetime.now())
    qual_file = open(path + str(file_name + '.txt'), 'w+', encoding='utf-8')
    for x in answers:
        qual_file.write(str(x[0]) + ' ' + str(x[1]) + '\n')

    res, risk = complete_risk_qual(answers)
    qual_file.write(res + '\n')
    qual_file.write(risk + '\n')
    qual_file.write('\n')
    qual_file.write('\n')
    qual_file.close()
    get_residual_risk_q(res, risk)


def get_residual_risk_q(res, risk):
    label['text'] = 'Вы прошли тестирование'
    button1.config(text='Сравнить с предыдущим тестированием')
    button2.config(text='Выйти')

    button1.config(command=verify_q)
    button2.config(command=exit)


def verify_q():
    # получить последний файл с результатом тестирования
    file_q = get_last_qual_file()
    # show_selected_file(file_q)
    prev_p, level = get_prev_risk_q(file_q)
    T.delete(0.0, END)
    res, risk = complete_risk_qual(answers)
    complete_residual_risk_q(prev_p, level, float(res.split(':')[1].replace(' ', '')), risk.replace(' ', ''))


def complete_residual_risk_q(prev_p, prev_level, actual_p, actual_level):
    T.delete(0.0, END)
    if prev_p > actual_p:
        T.insert(END, '''Риск невыполнения требований стал меньше.\nОн составляет: ''' + str(actual_p))
        T.insert(END, '\n')
        T.insert(END,
                 '''Уровень изменился.\nУровень: ''' + str(round(actual_p, 2) * 100) + '%, предыдущий риск: ' + str(
                     round(prev_p, 2) * 100) + '%')
        T.insert(END, '\n')
        show_wrong_answers()
    elif prev_p < actual_p:
        T.insert(END, '''Риск невыполнения требований стал больше.\nОн составляет: ''' + str(actual_p))
        T.insert(END, '\n')
        T.insert(END,
                 '''Уровень изменился.\nУровень: ''' + str(round(actual_p, 2) * 100) + '%, предыдущий риск: ' + str(
                     round(prev_p, 2) * 100) + '%')
        T.insert(END, '\n')
        show_wrong_answers()
    elif prev_p == actual_p:
        T.insert(END, '''Риск невыполнения требований не изменился.\nОн так же составляет: ''' + str(actual_p))
        T.insert(END, '\n')
        T.insert(END,
                 '''Уровень не изменился.\nУровень: ''' + str(round(actual_p, 2) * 100) + '%, предыдущий риск: ' + str(
                     round(prev_p, 2) * 100) + '%')
        T.insert(END, '\n')
        show_wrong_answers()


def get_prev_risk_q(file_q):
    f = open(file_q, 'r', encoding='utf-8')
    lines = f.readlines()
    p = float(lines[-4].split(':')[1].replace(' ', ''))
    level = lines[-3]
    return p, level


def complete_risk_qual(answers):
    count_all_questions = len(answers)
    count_pos_question = 0
    for x in answers:
        if x[1] == '+':
            count_pos_question += 1
        elif x[1] == '-':
            wrong_answers.append(x[0])
    return result_qual(count_all_questions, count_pos_question)


def result_qual(all_q, pos_q):
    res = pos_q / all_q
    if res < (1 / 3):
        T.insert(END, '\n')
        T.insert(END, 'Количественная оценка: ' + str(res))
        T.insert(END, '\n')
        T.insert(END, "Низкий" + '\n')
        return 'Количественная оценка: ' + str(res), "Низкий"
    elif (1 / 3) < res < (2 / 3):
        T.insert(END, '\n')
        T.insert(END, 'Количественная оценка: ' + str(res))
        T.insert(END, '\n')
        T.insert(END, "Средний" + '\n')
        return 'Количественная оценка: ' + str(res), "Средний"
    else:
        T.insert(END, '\n')
        T.insert(END, 'Количественная оценка: ' + str(res))
        T.insert(END, '\n')
        T.insert(END, "Высокий" + '\n')
        return 'Количественная оценка: ' + str(res), "Высокий"


def show_answers():
    label.config(text='Выберите тип оценки:')
    button1.config(text='Качественная оценка')
    button2.config(text='Количественная оценка')

    button1.config(command=b_quan)
    button2.config(command=b_qual)


def b_quan():
    path = os.getcwd() + r"/data/quan/"
    txt_files = filter(lambda x: x.endswith('.txt'), os.listdir(path))
    for i, x in enumerate(txt_files):
        files.append((i, x))
    button2.grid_forget()
    entry.grid(column=1, row=2)
    window.update()
    select_file_qn(path, files)


def b_qual():
    path = os.getcwd() + r"/data/qual/"
    txt_files = filter(lambda x: x.endswith('.txt'), os.listdir(path))
    for i, x in enumerate(txt_files):
        files.append((i, x))

    button2.grid_forget()
    entry.grid(column=1, row=2)
    window.update()
    select_file_ql(path, files)


def select_file_ql(path, files):
    for x in files:
        T.insert(END, str(x[0]) + ' ' + str(x[1]))
        T.insert(END, '\n')
    label.config(text='Выберите файл:')
    button1.config(command=open_chose_file_ql)
    button1.config(text='Посмотерть')


def select_file_qn(path, files):
    for x in files:
        T.insert(END, str(x[0]) + ' ' + str(x[1]))
        T.insert(END, '\n')
    label.config(text='Выберите файл:')
    button1.config(command=open_chose_file_qn)
    button1.config(text='Посмотерть')


def open_chose_file_ql():
    path = os.getcwd() + r"/data/qual/"
    file_q = entry.get()
    for x in files:
        if int(file_q) == x[0]:
            T.delete(0.0, END)
            T.insert(END, '\n')
            T.insert(END, "Вы выбрали файл " + x[1])
            T.insert(END, '\n')
            show_selected_file(path + x[1])


def open_chose_file_qn():
    path = os.getcwd() + r"/data/quan/"
    file_q = entry.get()
    for x in files:
        if int(file_q) == x[0]:
            T.delete(0.0, END)
            T.insert(END, '\n')
            T.insert(END, "Вы выбрали файл " + x[1])
            T.insert(END, '\n')
            show_selected_file(path + x[1])


def show_selected_file(file_q):
    file_q = open(file_q, 'r', encoding='utf-8')
    lines = file_q.readlines()
    for line in lines:
        T.insert(END, line.strip())
        T.insert(END, '\n')
    file_q.close


# if choise == 1:
# 	path = os.getcwd() + r"/data/quan/"
# 	txt_files = filter(lambda x: x.endswith('.txt'), os.listdir(path))
# 	for i,x in enumerate(txt_files):
# 		files.append((i, x))
# 	select_file(path,files)
# elif choise == 2:
# 	path = os.getcwd() + r"/data/qual/"
# 	txt_files = filter(lambda x: x.endswith('.txt'), os.listdir(path))
# 	for i,x in enumerate(txt_files):
# 		files.append((i, x))
# 	select_file(path, files)

window = Tk()
w = window.winfo_screenwidth()
h = window.winfo_screenheight()
w = int(w // 1.1)
h = int(h // 1.5)
window.geometry(f'{w}x{h}')
window.configure(background='grey')
label = Label(window,
              text="Выберите вариант ответа:",
              fg="white",
              background='grey',
              font=("Arial Bold", 11),
              )
label.grid(column=1, row=0, columnspan=3, rowspan=2)

button1 = Button(window,
                 text="Пройти тест",
                 fg="white",
                 background='grey',
                 width=w//35,
                 command=complete_test,
                 )
button1.grid(column=0, row=2,  columnspan=2)

button2 = Button(window,
                 text="Смотреть ответы",
                 fg="white",
                 background='grey',
                 width=w//35,
                 command=show_answers,
                 )
button2.grid(column=3, row=2, columnspan=2)

entry = Entry(window)
entry.grid(column=2, row=2)
entry.grid_forget()

T = Text(window, height=30, font=("Arial Bold", 12))
T.grid(column=0, row=3, columnspan=5, sticky='we')
window.grid_columnconfigure(0, weight=1)
window.grid_columnconfigure(1, weight=1)
window.grid_columnconfigure(3, weight=1)
window.grid_columnconfigure(4, weight=1)
window.mainloop()
