# !/usr/bin/python
# -*- coding: utf-8 -*-
import os
import shutil
import subprocess
import sys, re
from PyQt5.QtWidgets import QApplication, QLabel, QWidget, QPushButton, QAction, QMainWindow
from PyQt5 import QtWidgets, QtGui, QtCore
# from PyQt5.QtCore import QMimeData, Qt, QRect
import PerformanceTestingReport as res
from pathlib import Path
import webbrowser

btn_style = """
QPushButton {
    font-weight: bold;
    border:1px solid rgb(255, 164, 81); 
    width:300px; 
    background-color:rgb(255, 164, 81);
    color:white;
    border-radius:10px; 
    padding:2px 4px;
}
QPushButton:hover:!pressed
{
    font-weight: bold;
    border:1px solid rgb(247, 186, 133); 
    width:300px; 
    background-color:rgb(244, 142, 53);
    color:white;
    border-radius:10px; 
    padding:2px 4px;
}
QPushButton:pressed
{
    border:1px solid rgb(255, 164, 81); 
    width:300px; 
    background-color:white;
    color:rgb(255, 164, 81);
    border-radius:10px; 
    padding:2px 4px;
}
"""

btn_style2 = """
QPushButton {
    border:1px solid rgb(255, 164, 81); 
    width:300px; 
    background-color:white;
    color:rgb(255, 164, 81);
    border-radius:10px; 
    padding:2px 4px;
}
QPushButton:hover:!pressed
{
    border:1px solid rgb(244, 172, 110); 
    width:300px; 
    background-color:rgb(244, 172, 110);
    color:white;
    border-radius:10px; 
    padding:2px 4px;
}
QPushButton:pressed
{
    border:1px solid rgb(244, 172, 110); 
    width:300px; 
    background-color:rgb(244, 172, 110);
    color:white;
    border-radius:10px; 
    padding:2px 4px;
}
"""


class DropLabel(QLabel):
    def __init__(self, *args, **kwargs):
        QLabel.__init__(self, *args, **kwargs)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        file_name = event.mimeData().text()
        if file_name.split('.')[-1] in ['xlsx']:
            event.acceptProposedAction()
        else:
            event.ignore()

    # def dragEnterEvent(self, event):
    #     if event.mimeData().hasText():
    #         event.acceptProposedAction()

    def dropEvent(self, event):

        pos = event.pos()
        file_path = event.mimeData().text()
        file_path = re.sub('file:///', '', file_path)
        self.setText(file_path)
        event.acceptProposedAction()


class Widget(QMainWindow):
    def __init__(self):
        super(Widget, self).__init__()
        self.setWindowTitle('效能測試圖表')
        # self.resize(335, 250)
        self.initUI()

        # Window size
        self.WIDTH = 470
        self.HEIGHT = 350
        self.resize(self.WIDTH, self.HEIGHT)

        # Widget

        self.setWindowOpacity(0.9)
        radius = 30
        self.setStyleSheet(
            """
            background:rgb(255, 255, 255);
            border-top-left-radius:{0}px;
            border-bottom-left-radius:{0}px;
            border-top-right-radius:{0}px;
            border-bottom-right-radius:{0}px;
            """.format(radius)
        )

    def initUI(self):
        self.label0 = QtWidgets.QLabel('數據由 PerfDog(v5.1) 版本導出:', self)
        self.label0.setGeometry(QtCore.QRect(40, 10, 341, 41))
        self.label0.setFont(QtGui.QFont('Microsoft JhengHei', 12))
        # self.label0.setAlignment(QtCore.Qt.AlignCenter)

        self.DLabel = DropLabel("將【PD_2021XXXX.xlsx】 檔案 拖曳至此", self)

        self.DLabel.setFont(QtGui.QFont('Microsoft JhengHei', 12))
        self.DLabel.setGeometry(40, 50, 391, 121)
        self.DLabel.setWordWrap(True)  # 自動換行
        self.DLabel.setAlignment(QtCore.Qt.AlignCenter)
        # self.DLabel.setStyleSheet('background-color: rgb(255, 255, 255);')
        self.DLabel.setStyleSheet(
            """
            color:rgb(158, 153, 153);
            font-weight: bold;
            background:rgb(239, 237, 238);
            border-top-left-radius:5px;
            border-bottom-left-radius:5px;
            border-top-right-radius:5px;
            border-bottom-right-radius:5px;
            """)
        # label_to_drag = DraggableLabel("drag this", self)

        # self.pushButton = QtWidgets.QPushButton(qta.icon('fa.arrow-circle-right', color='white'), " 轉換 ", self)
        self.pushButton = QtWidgets.QPushButton(" Start  ", self)
        self.pushButton.setGeometry(QtCore.QRect(100, 215, 280, 45))
        self.pushButton.setFont(QtGui.QFont('Microsoft JhengHei', 12))
        self.pushButton.clicked.connect(self.buttonClicked)
        self.pushButton.setStyleSheet(btn_style)

        self.label = QtWidgets.QLabel(self)
        self.label.setFont(QtGui.QFont('Microsoft JhengHei', 10))
        self.label.setGeometry(QtCore.QRect(70, 170, 341, 41))
        self.label.setWordWrap(True)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setStyleSheet("color:rgb(178, 173, 175)")

        self.pushButton_del = QtWidgets.QPushButton("📁 開啟測試報告", self)
        self.pushButton_del.setFont(QtGui.QFont('Microsoft JhengHei', 11))
        self.pushButton_del.setGeometry(QtCore.QRect(100, 273, 135, 40))
        self.pushButton_del.setStyleSheet(btn_style2)
        self.pushButton_del.clicked.connect(self.DelbuttonClicked)

        self.pushButton_2 = QtWidgets.QPushButton("🚫 檢視效能圖表", self)
        self.pushButton_2.setGeometry(QtCore.QRect(245, 273, 135, 40))
        self.pushButton_2.setFont(QtGui.QFont('Microsoft JhengHei', 11))
        self.pushButton_2.setStyleSheet(btn_style2)
        self.pushButton_2.clicked.connect(self.openButtonClicked)
        self.pushButton_2.setDisabled(True)
        self.show()

    def buttonClicked(self):
        datapath = self.DLabel.text()
        CheckDataExist = os.path.exists(datapath)

        if CheckDataExist == True:
            try:
                res.plotrun(datapath)
                self.finalPath, self.finalName = res.changeName(datapath)
                self.label.setText('檔名:' + str(self.finalName))
                self.pushButton_2.setDisabled(False)
                self.pushButton_2.setText('✅ 檢視效能圖表')
            except Exception as e:
                print(e)
        else:
            self.label.setText('路徑錯誤。')

    def DelbuttonClicked(self):
        self.label.setText('')
        path = os.getcwd()
        # subprocess.call(['explorer', path])

    def openButtonClicked(self):
        # pathReport = os.path.split(os.path.abspath(self.finalINFO))[0]
        pathReport = os.path.split(self.finalPath)[0]
        print('圖表路徑:', pathReport)

        subprocess.call(['explorer', ".\Report"])
        webbrowser.get("C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s").open(
            'file://' + self.finalPath)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Breeze')
    window = Widget()
    window.show()
    sys.exit(app.exec_())
