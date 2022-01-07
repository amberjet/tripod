import sys
import xlsxwriter

from PyQt5.QtWidgets import QWidget, QApplication, QMainWindow, QLabel, QDialog
from PyQt5.QtWidgets import QPushButton, QHBoxLayout, QVBoxLayout, QLineEdit
from PyQt5.QtGui import QIntValidator
from PyQt5.QtCore import Qt
from PyQt5 import uic

f = open("files/opros.txt")
QUESTIONS = f.readlines()  # список вопросов
f.close()
REVERSE_NUMBERS = [13, 19, 22, 24, 26]  # номера вопросов от обратного
CARE = [6, 18, 23, 28, 29]  # вопросы поддержки
CAPTIVATE = [1, 20, 10, 22, 9, 17]  # вопросы интереса
CMANAGE = [5, 11, 13, 16, 19, 21, 26]  # вопросы управления поведением
CONFER = [2, 3, 25, 27, 30, 12]  # вопросы уважения
CLARIFY = [4, 8, 24, 14, 7, 15]  # вопросы объяснений
d = open("files/data.txt")
data = d.readlines()
d.close()
if len(data) == 0:
    answers = [0] * len(QUESTIONS)
    total_stud = 0
else:
    total_stud = int(data[0])
    answers = [int(x) for x in data[1].strip().split()]
res_dict = {}


class Tripod(QWidget):  # основное окно
    def __init__(self):
        super(Tripod, self).__init__()
        self.setWindowTitle('Трипод')

        self.main_layout = QVBoxLayout(self)
        self.buttons_layout = QVBoxLayout(self)
        self.k_ans = QLabel(self)
        self.k_ans.setText('Количество анкет, внесенных в систему: ' + str(total_stud))
        self.add_tripod = QPushButton(self)
        self.add_tripod.setText('Добавить ответы')
        self.add_tripod.clicked.connect(self.make_question)
        self.get_result = QPushButton(self)
        self.get_result.setText('Сформировать отчет')
        self.get_result.clicked.connect(make_report)

        self.buttons_layout.addWidget(self.add_tripod)
        self.buttons_layout.addWidget(self.get_result)
        self.main_layout.addWidget(self.k_ans)
        self.main_layout.addLayout(self.buttons_layout)
        self.setLayout(self.main_layout)

    def make_question(self):
        q = Question(1)
        q.show()
        q.exec()
        self.k_ans.setText('Ответов в системе: ' + str(total_stud))


class No_Anketes(QDialog):  # если ответов в системе нет
    def __init__(self):
        super(No_Anketes, self).__init__()
        self.setWindowTitle('Ошибка')
        self.main_layout = QVBoxLayout(self)
        self.end_text = QLabel(self)
        self.end_text.setText('В системе нет анкет')
        self.ok_button = QPushButton(self)
        self.ok_button.setText('ОК')
        self.ok_button.clicked.connect(self.accept)
        self.main_layout.addWidget(self.end_text)
        self.main_layout.addWidget(self.ok_button)
        self.setLayout(self.main_layout)


class Question(QDialog):  # окно ввода ответов
    def __init__(self, pos, parent=Tripod):
        self.pos = pos
        self.parent = parent
        self.new_ans = [0] * len(QUESTIONS)
        super(Question, self).__init__()
        self.setWhatsThis('Окно для ввода ответов')
        self.setWindowTitle('Вопросы анкеты')
        self.main_layout = QVBoxLayout(self)
        self.buttons_layout = QHBoxLayout(self)
        self.question_text = QLabel(self)
        self.question_text.setText(QUESTIONS[self.pos])
        self.rule1 = QLabel(self)
        self.rule1.setText('1 - Неверно')
        self.rule2 = QLabel(self)
        self.rule2.setText('2 - Скорее неверно')
        self.rule3 = QLabel(self)
        self.rule3.setText('3 - Затрудняюсь ответить')
        self.rule4 = QLabel(self)
        self.rule4.setText('4 - Скорее верно')
        self.rule5 = QLabel(self)
        self.rule5.setText('5 - Верно')
        self.current_answer = QLineEdit(self)
        # потому что SpinBox замедляет ввод с клавиатуры
        self.current_answer.setFocus()
        self.current_answer.setValidator(QIntValidator(1, 5))
        self.inst = QLabel(self)
        self.inst.setText('Введите ответ и нажмите Enter')
        self.stop_test = QPushButton(self)
        self.stop_test.setText('Закончить ввод')
        self.stop_test.setDisabled(False)
        self.stop_test.clicked.connect(self.areyoureallysure)

        self.buttons_layout.addWidget(self.stop_test)
        self.main_layout.addWidget(self.question_text)
        self.main_layout.addWidget(self.rule1)
        self.main_layout.addWidget(self.rule2)
        self.main_layout.addWidget(self.rule3)
        self.main_layout.addWidget(self.rule4)
        self.main_layout.addWidget(self.rule5)
        self.main_layout.addWidget(self.current_answer)
        self.main_layout.addWidget(self.inst)
        self.main_layout.addLayout(self.buttons_layout)
        self.setLayout(self.main_layout)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Enter or event.key() == Qt.Key_Return:
            res = self.check_current_answer()
            if res:
                if self.pos < len(QUESTIONS) - 1:
                    self.pos += 1
                    self.question_text.setText(QUESTIONS[self.pos])
                elif self.pos == len(QUESTIONS) - 1:
                    append_answers(self.new_ans)
                    self.checkEOQ()
                    self.close()

    def checkEOQ(self):
        q = EndOfQuest()
        result = q.exec_()
        if result == QDialog.Accepted:
            self.hide()
            q = Question(1)
            q.show()
            q.exec()

    def areyoureallysure(self):
        sure_dialog = AreYouSure()
        result = sure_dialog.exec_()
        if result == QDialog.Rejected:
            self.close()

    def check_current_answer(self):
        if self.current_answer.text() not in '12345':
            warning = WrongAnswer()
            result = warning.exec_()
            if result == QDialog.Accepted:
                self.current_answer.clear()
                self.current_answer.setFocus()
            return False
        if self.current_answer.text() == '':
            return False
        else:
            if self.pos in REVERSE_NUMBERS:
                if self.current_answer.text() == '1' or self.current_answer.text() == '2':
                    self.new_ans[self.pos] += 1
            else:
                if self.current_answer.text() == '4' or self.current_answer.text() == '5':
                    self.new_ans[self.pos] += 1
        self.current_answer.clear()
        self.current_answer.setFocus()
        return True


class ReadyReport(QDialog):  # когда отчет готов
    def __init__(self):
        super(ReadyReport, self).__init__()
        self.setWindowTitle('Готово')
        self.setWhatsThis('')
        self.main_layout = QVBoxLayout(self)
        self.end_text = QLabel(self)
        self.end_text.setText('Отчёт "Результаты.xlxs" сформирован и лежит в папке с программой')
        self.ok_button = QPushButton(self)
        self.ok_button.setText('ОК')
        self.ok_button.clicked.connect(self.accept)
        self.main_layout.addWidget(self.end_text)
        self.main_layout.addWidget(self.ok_button)
        self.setLayout(self.main_layout)


class EndOfQuest(QDialog):  # окно, когда одна анкета закончена
    def __init__(self, parent=Question):
        super(EndOfQuest, self).__init__()
        self.setWindowTitle('Вопросы анкеты')
        self.main_layout = QVBoxLayout(self)
        self.setWhatsThis('')
        self.buttons_layout = QHBoxLayout(self)
        self.end_text = QLabel(self)
        self.end_text.setText('Анкета заполнена. Хотите продолжить?')
        self.add_test = QPushButton(self)
        self.add_test.setText('Еще одна анкета')
        self.add_test.clicked.connect(self.accept)
        self.stop_test = QPushButton(self)
        self.stop_test.setText('Закончить ввод')
        self.stop_test.clicked.connect(self.close)

        self.buttons_layout.addWidget(self.add_test)
        self.buttons_layout.addWidget(self.stop_test)
        self.main_layout.addWidget(self.end_text)
        self.main_layout.addLayout(self.buttons_layout)
        self.setLayout(self.main_layout)


class AreYouSure(QDialog):  # окно, если анкета прерывается на середине
    def __init__(self, parent=Question):
        super(AreYouSure, self).__init__()
        self.setWindowTitle('Анкета не закончена!')
        self.setWhatsThis('')
        self.main_layout = QVBoxLayout(self)
        self.buttons_layout = QHBoxLayout(self)
        self.end_text = QLabel(self)
        self.end_text.setText('Вы не внесли все ответы на вопросы.')
        self.end_text2 = QLabel(self)
        self.end_text2.setText('Внесенные ответы текущей анкеты не сохранятся')
        self.end_text3 = QLabel(self)
        self.end_text3.setText('Вы уверены?')
        self.add_test = QPushButton(self)
        self.add_test.setText('Да, закончить')
        self.add_test.clicked.connect(self.reject)
        self.stop_test = QPushButton(self)
        self.stop_test.setText('Нет, продолжить ввод ответов')
        self.stop_test.clicked.connect(self.accept)

        self.buttons_layout.addWidget(self.add_test)
        self.buttons_layout.addWidget(self.stop_test)
        self.main_layout.addWidget(self.end_text)
        self.main_layout.addWidget(self.end_text2)
        self.main_layout.addWidget(self.end_text3)
        self.main_layout.addLayout(self.buttons_layout)
        self.setLayout(self.main_layout)


class WrongAnswer(QDialog):  # если ответ некорректный
    def __init__(self):
        super(WrongAnswer, self).__init__()
        uic.loadUi('report.ui', self)
        self.setWindowTitle('Ошибка')
        self.setWhatsThis('')
        self.end_text.setText('Допустимые ответы: 1, 2, 3, 4, 5')
        self.ok_button.setText('ОК')
        self.ok_button.clicked.connect(self.accept)


class Window(QMainWindow):
    def __init__(self):
        super(Window, self).__init__()
        self.setWindowTitle('Трипод')
        self.setCentralWidget(Tripod())

def append_answers(a):
    global answers  # да-да, это плохо, это нужно убрать
    global total_stud
    total_stud += 1
    for i in range(len(answers)):
        answers[i] += a[i]

def save_to_file():
    f = open("files/data.txt", 'w')
    print(total_stud, file=f)
    print(' '.join(str(x) for x in answers), file=f)
    f.close()

def obrabotka():
    ans = list(enumerate(answers))
    max_care = total_stud * len(CARE)
    max_clar = total_stud * len(CLARIFY)
    max_capt = total_stud * len(CAPTIVATE)
    max_conf = total_stud * len(CONFER)
    max_cm = total_stud * len(CMANAGE)
    res_dict['Поддержка'] = round((sum(map(lambda x: x[1], filter(lambda x: x[0] in CARE, ans))) / max_care) * 100, 2)
    res_dict['Объяснение'] = round((sum(map(lambda x: x[1], filter(lambda x: x[0] in CLARIFY, ans))) / max_clar) *
                                   100, 2)
    res_dict['Интерес'] = round((sum(map(lambda x: x[1], filter(lambda x: x[0] in CAPTIVATE, ans))) / max_capt)
                                * 100, 2)
    res_dict['Уважение'] = round((sum(map(lambda x: x[1], filter(lambda x: x[0] in CONFER, ans))) / max_conf) * 100, 2)
    res_dict['Управление'] = round((sum(map(lambda x: x[1], filter(lambda x: x[0] in CMANAGE, ans))) / max_cm) * 100, 2)
    return res_dict


def make_report():
    global res_dict
    global total_stud
    if total_stud == 0:
        warning = No_Anketes()
        result = warning.exec_()
    else:
        obrabotka()
        workbook = xlsxwriter.Workbook('Результаты.xlsx')
        worksheet = workbook.add_worksheet()
        data = [(key, res_dict[key]) for key in res_dict]

        for row, (pok, res) in enumerate(data):
            worksheet.write(row, 0, pok)
            worksheet.write(row, 1, res)

        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({'values': '=Sheet1!B1:B5'})
        worksheet.insert_chart('D6', chart)
        workbook.close()
        warning = ReadyReport()
        result = warning.exec_()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    wnd = Window()
    wnd.show()
    app.exec()
    if total_stud > 0:
        save_to_file()
