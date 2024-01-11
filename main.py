# -*- coding: utf-8 -*-
# https://blog.csdn.net/weixin_46554689
# Created by: GMCY
# 2021/2/26

from re import match
from os import system
from sys import exit, argv
from time import localtime
from PyQt5 import QtCore, QtGui, QtWidgets
from win32con import CF_UNICODETEXT
from pyautogui import click, moveTo, hotkey, press, sleep, position
from win32clipboard import OpenClipboard, EmptyClipboard, SetClipboardData, CloseClipboard
from threading import Thread


def mouse_set_text(String):
    OpenClipboard()
    EmptyClipboard()
    SetClipboardData(CF_UNICODETEXT, String)
    CloseClipboard()


def mouse_press(x, y):
    moveTo(x, y)
    click(button='left')


def mouse_search(x, y, String):
    mouse_set_text(String=String)
    sleep(0.2)
    moveTo(x, y)
    click(button='left')
    press('psp')
    hotkey('ctrl', 'a')
    press('psp')
    hotkey('ctrl', 'v')
    press('enter')


def history_write(String):
    times = localtime()
    daytime = str(times[0]) + '-' + str(times[1]) + '-' + str(times[2]) + '\t' + str(times[3]) + ':' + str(
        times[4]) + ':' + str(times[5]) + '\t'
    with open('History.txt', 'a', encoding='utf-8') as f:
        f.write(daytime + String + '\n')


def history_main():
    system('start /B History.txt')
    # system('History.txt')


def help_main():
    system('start /B Help.txt')
    # system('Help.txt')


class UiCleverMouse(object):
    def __init__(self):
        self.lineEdit_show = None
        self.label_command = None
        self.lineEdit_command = None
        self.label_times = None
        self.lineEdit_times = None
        self.pushButton_start = None
        self.pushButton_stop = None
        self.lineEdit_mouse_pos = None
        self.widget_use = None
        self.pushButton_help = None
        self.pushButton_history = None
        self.widget_set = None
        self.centralWidget = None
        self.error = 0
        self.error_start = 0
        self.error_history = 0
        self.error_help = 0

    def setupUi(self, Clever_mouse):
        Clever_mouse.setObjectName("Clever_mouse")
        Clever_mouse.resize(551, 181)
        Clever_mouse.setMinimumSize(QtCore.QSize(551, 181))
        Clever_mouse.setMaximumSize(QtCore.QSize(551, 181))

        Clever_mouse.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)

        self.centralWidget = QtWidgets.QWidget(Clever_mouse)
        self.centralWidget.setObjectName("centralwidget")
        self.widget_set = QtWidgets.QWidget(self.centralWidget)
        self.widget_set.setGeometry(QtCore.QRect(0, 0, 151, 181))
        self.widget_set.setStyleSheet("background-color: rgb(245, 245, 245);")
        self.widget_set.setObjectName("widget_set")
        self.pushButton_history = QtWidgets.QPushButton(self.widget_set)
        self.pushButton_history.setGeometry(QtCore.QRect(20, 70, 111, 41))
        font = QtGui.QFont()
        font.setPointSize(15)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_history.setFont(font)
        self.pushButton_history.setStyleSheet("background-color: rgb(160, 160, 160);")
        self.pushButton_history.setObjectName("pushButton_history")
        self.pushButton_help = QtWidgets.QPushButton(self.widget_set)
        self.pushButton_help.setGeometry(QtCore.QRect(20, 120, 111, 41))
        font = QtGui.QFont()
        font.setPointSize(15)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_help.setFont(font)
        self.pushButton_help.setStyleSheet("background-color: rgb(160, 160, 160);")
        self.pushButton_help.setObjectName("pushButton_help")
        self.lineEdit_mouse_pos = QtWidgets.QLineEdit(self.widget_set)
        self.lineEdit_mouse_pos.setGeometry(QtCore.QRect(20, 20, 111, 41))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.lineEdit_mouse_pos.setFont(font)
        self.lineEdit_mouse_pos.setObjectName("lineEdit_mouse_pos")
        self.widget_use = QtWidgets.QWidget(self.centralWidget)
        self.widget_use.setGeometry(QtCore.QRect(150, 0, 401, 181))
        self.widget_use.setStyleSheet("background-color: rgb(245, 245, 245);")
        self.widget_use.setObjectName("widget_use")
        self.pushButton_stop = QtWidgets.QPushButton(self.widget_use)
        self.pushButton_stop.setGeometry(QtCore.QRect(270, 120, 111, 41))
        font = QtGui.QFont()
        font.setPointSize(15)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_stop.setFont(font)
        self.pushButton_stop.setStyleSheet("background-color: rgb(255, 85, 0);")
        self.pushButton_stop.setObjectName("pushButton_stop")
        self.pushButton_start = QtWidgets.QPushButton(self.widget_use)
        self.pushButton_start.setGeometry(QtCore.QRect(70, 120, 111, 41))
        font = QtGui.QFont()
        font.setPointSize(15)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_start.setFont(font)
        self.pushButton_start.setStyleSheet("background-color: rgb(85, 255, 127);")
        self.pushButton_start.setObjectName("pushButton_start")
        self.lineEdit_times = QtWidgets.QLineEdit(self.widget_use)
        self.lineEdit_times.setGeometry(QtCore.QRect(70, 20, 61, 41))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.lineEdit_times.setFont(font)
        self.lineEdit_times.setObjectName("lineEdit_times")
        self.label_times = QtWidgets.QLabel(self.widget_use)
        self.label_times.setGeometry(QtCore.QRect(0, 20, 61, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_times.setFont(font)
        self.label_times.setObjectName("label_times")
        self.lineEdit_command = QtWidgets.QLineEdit(self.widget_use)
        self.lineEdit_command.setGeometry(QtCore.QRect(70, 70, 311, 41))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.lineEdit_command.setFont(font)
        self.lineEdit_command.setObjectName("lineEdit_command")
        self.label_command = QtWidgets.QLabel(self.widget_use)
        self.label_command.setGeometry(QtCore.QRect(0, 70, 61, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_command.setFont(font)
        self.label_command.setObjectName("label_command")
        self.lineEdit_show = QtWidgets.QLineEdit(self.widget_use)
        self.lineEdit_show.setGeometry(QtCore.QRect(140, 20, 241, 41))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.lineEdit_show.setFont(font)
        self.lineEdit_show.setObjectName("lineEdit_show")
        Clever_mouse.setCentralWidget(self.centralWidget)

        self.reTranslateUi(Clever_mouse)
        QtCore.QMetaObject.connectSlotsByName(Clever_mouse)

    def reTranslateUi(self, Clever_mouse):
        _translate = QtCore.QCoreApplication.translate
        Clever_mouse.setWindowTitle(_translate("Clever_mouse", "CleverMouse"))
        self.pushButton_history.setText(_translate("Clever_mouse", "历史"))
        self.pushButton_help.setText(_translate("Clever_mouse", "说明"))
        self.pushButton_stop.setText(_translate("Clever_mouse", "停止"))
        self.pushButton_start.setText(_translate("Clever_mouse", "执行"))
        self.lineEdit_times.setText(_translate("Clever_mouse", "1"))
        self.lineEdit_show.setText(_translate("Clever_mouse", "提示语"))
        self.label_times.setText(_translate("Clever_mouse", "次数："))
        self.label_command.setText(_translate("Clever_mouse", "指令："))

        self.start_set()

    def start_set(self):
        self.pushButton()

        thread = Thread(target=self.show_mouse_pos)
        thread.start()

    def show_mouse_pos(self):
        try:
            while True:
                x, y = position()
                self.lineEdit_mouse_pos.setText(str(x) + ',' + str(y))
                sleep(0.5)
        except RuntimeError:
            pass

    def pushButton(self):
        self.pushButton_history.clicked.connect(self.history)
        self.pushButton_help.clicked.connect(self.help)
        self.pushButton_start.clicked.connect(self.start)
        self.pushButton_stop.clicked.connect(self.stop)

    def history(self):
        thread = Thread(target=self.history_thread)
        thread.start()

    def history_thread(self):
        if self.error_history == 0:
            self.error_history = 1

            thread = Thread(target=history_main)
            thread.start()

            threads = [thread]

            while threads:
                threads.pop().join()
            self.error_history = 0

    def help(self):
        thread = Thread(target=self.help_thread)
        thread.start()

    def help_thread(self):
        if self.error_help == 0:
            self.error_help = 1

            thread = Thread(target=help_main)
            thread.start()

            threads = [thread]

            while threads:
                threads.pop().join()
            self.error_help = 0

    def start(self):
        thread = Thread(target=self.start_thread)
        thread.start()

    def start_thread(self):
        self.error = 0
        if self.error_start == 0:

            self.error_start = 1

            thread = Thread(target=self.start_main)
            thread.start()

            threads = [thread]
            while threads:
                threads.pop().join()
            self.error_start = 0

    def start_main(self):
        String = self.lineEdit_command.text()
        times = self.lineEdit_times.text()

        history_write(str(String))

        String_lt = String.split(';')

        if String_lt[-1] == '':
            String_lt.remove(String_lt[-1])

        time = r'^\d{1,3}$'
        P = r'^P:\d{1,},\d{1,}$'
        S = r'^S:\d{1,},\d{1,}:.{1,}$'
        T = r'^T:\d{1,}$'

        if match(time, times):
            error_time = 0
            self.lineEdit_show.setText('次数格式正确。')
        else:
            error_time = 1
            self.lineEdit_show.setText('次数格式错误!!!')
        sleep(1)

        error_str = 0

        if error_time == 0 and String != '':
            for Str in String_lt:
                if not match(P, Str) and not match(S, Str) and not match(T, Str):
                    error_str = 1
                    self.lineEdit_show.setText('错误:' + Str)
                    sleep(0.5)
            if error_str == 0:
                self.lineEdit_show.setText('指令正确。')
                sleep(0.5)
        else:
            error_str = 1
            self.lineEdit_show.setText('指令不能未空!!!')

        if error_time == 0 and error_str == 0 and self.error == 0:
            if self.error == 0:
                self.lineEdit_show.setText('开始执行。')
                sleep(0.5)
                self.lineEdit_show.setText('执行中请不要操作')
            if self.error == 0:
                sleep(1)
                self.lineEdit_show.setText('3s后开始执行指令!')
            if self.error == 0:
                sleep(1)
                self.lineEdit_show.setText('3')
            if self.error == 0:
                sleep(1)
                self.lineEdit_show.setText('2')
            if self.error == 0:
                sleep(1)
                self.lineEdit_show.setText('1')
                sleep(0.5)

            for t in range(1, int(times) + 1):
                if self.error == 0:
                    self.mouse(String_lt, t)
                else:
                    pass

            if self.error == 1:
                self.lineEdit_show.setText('指令执行停止')
            else:
                self.lineEdit_show.setText('完成指令。')

    def mouse(self, String_lt, time):
        self.lineEdit_show.setText('第%s次指令开始执行' % str(time))
        sleep(0.5)
        for Str in String_lt:
            Command_lt = Str.split(':')
            Command = Command_lt[0]
            if Command == 'P':
                x, y = Command_lt[1].split(',')
                mouse_press(int(x), int(y))
            elif Command == 'S':
                x, y = Command_lt[1].split(',')
                mouse_search(int(x), int(y), Command_lt[2].replace('[num]', str(time)))
            elif Command == 'T':
                sleep(int(Command_lt[1]))
            else:
                pass
        self.lineEdit_show.setText('第%s次指令执行成功' % str(time))
        sleep(0.5)

    def stop(self):
        self.error = 1


def show_Window():
    App = QtWidgets.QApplication(argv)
    App.setWindowIcon(QtGui.QIcon('CleverMouse_1.ico'))
    MainWindow = QtWidgets.QMainWindow()
    ui = UiCleverMouse()
    ui.setupUi(MainWindow)
    MainWindow.show()
    exit(App.exec_())


if __name__ == '__main__':
    show_Window()
