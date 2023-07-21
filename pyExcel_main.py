#mqtt转发表检查约束条件：
#前提：
#通信方向：D3_DIRECTION 为 0(monitor)   errorcode: 0b00000001
#规约类型D3_PROT_TYPE为MQTT  errorcode: 0b00000010

#1、传输策略D3_TRNS_POLICY为周期传输，D3_CATEGORY 不能为报警/事件/遥控/遥调，且传输周期D3_TRANS_INTERVAL不能为0 errorcode: 0b00000100

#(1)是否记录D3_STORE_ACTIVE，若为是，断线存储时间不能是即时保存或者0.   errorcode: 0b00001000
#2、传输策略D3_TRNS_POLICY为即时传输，D3_CATEGORY 只能是报警/事件。且是否记录D3_STORE_ACTIVE只能选即时保存。    errorcode: 0b00010000


import openpyxl
import sys
import os
import time
from pprint import pprint
from pyExcelProcess import MyProcess
from PySide6 import QtWidgets

class MyWidget(QtWidgets.QWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setWindowTitle("配置工具检测v1.0")
        self.dialog = QtWidgets.QFileDialog(self)
        self.info_le = QtWidgets.QLineEdit(self)    #显示文件路径
        self.browse_button2 = QtWidgets.QPushButton(self)   #分析按钮
        self.TypeComboBox = QtWidgets.QComboBox(self)   #转发表类型
        self.ContentEdit = QtWidgets.QTextEdit(self)    #文字显示框
        self.progress_bar = QtWidgets.QProgressBar(self)    #进度条

        self.setup_dialog()
        self.setup_ui()
        self.setup_signals()
                
    def setup_dialog(self):
        self.dialog.setAcceptMode(QtWidgets.QFileDialog.AcceptOpen)
        self.dialog.setFileMode(QtWidgets.QFileDialog.ExistingFiles)
        self.dialog.setLabelText(QtWidgets.QFileDialog.DialogLabel.Accept, "选择")  # 为「接受」按键指定文本
        self.dialog.setLabelText(QtWidgets.QFileDialog.DialogLabel.Reject, "取消")  # 为「拒绝」按键指定文本
        self.dialog.setNameFilter("Excel Files(*.xlsx *.xls *.xlsm")

    def setup_ui(self):
        "主窗口大小"
        self.resize(400, 300)

        "配置控件"
        FileLabel = QtWidgets.QLabel("转发表路径:", self)
        TypeLabel = QtWidgets.QLabel("转发表类型:", self)

        browse_button1 = QtWidgets.QPushButton("浏览", self)
        browse_button1.setMaximumWidth(75)
        browse_button1.clicked.connect(self.dialog.open)

        self.browse_button2.setText("分析")
        self.browse_button2.setMaximumWidth(75)
        self.browse_button2.setStyleSheet("background-color:rgb(255, 255, 255)")
        self.browse_button2.clicked.connect(self.analyse_process)
        
        self.info_le.setMinimumWidth(200)

        self.TypeComboBox.setMaximumWidth(100)
        self.TypeComboBox.addItem("MQTT")
        self.TypeComboBox.addItem("IEC104")
        self.TypeComboBox.addItem("modbusTCP")

        self.progress_bar.setMaximumHeight(20)

        self.ContentEdit.setReadOnly(True)
        #ContentEdit.setHorizontalScrollBarPolicy(QtWidgets.Qt.ScrollBarAsNeeded)

        "布局"
        grid = QtWidgets.QGridLayout()
        grid.setSpacing(10)
        grid.addWidget(FileLabel, 1, 0)
        grid.addWidget(self.info_le, 2, 0)
        grid.addWidget(browse_button1, 2, 1)
        grid.addWidget(TypeLabel, 3, 0)
        grid.addWidget(self.TypeComboBox, 4, 0)
        grid.addWidget(self.browse_button2, 4, 1)
        grid.setColumnStretch(5, 0)
        grid.addWidget(self.progress_bar, 5, 0)
        grid.addWidget(self.ContentEdit, 6, 0, 1, 2)

        self.setLayout(grid)


    def setup_signals(self):
        self.dialog.fileSelected.connect(lambda path: self.info_le.setText(path))  # type: ignore
        #self.dialog.currentChanged.connect(lambda path: self.info_label.setText(path))  # type: ignore
        self.dialog.filesSelected.connect(lambda files: pprint(f"filesSelected: {files}"))  # type: ignore
        self.dialog.directoryEntered.connect(lambda dir_: print(f"directoryEntered: {dir_}"))  # type: ignore
        self.dialog.filterSelected.connect(lambda filter_: print(f"filterSelected: {filter_}"))  # type: ignore

    def PrintText(self, text):
        print(text, flush=True)
        self.ContentEdit.append(text)

    def analyse_process(self):
        self.browse_button2.setDisabled(True)
        self.browse_button2.setText("请稍后")
        #self.browse_button2.setStyleSheet("background-color:rgb(255, 0, 0)")
        print("分析中...")
        #time.sleep(3)
        type = self.TypeComboBox.currentText()
        #print(type)
        myprocess = MyProcess()
        if type == "MQTT":
            errcount = myprocess.MQTTAnalyse(self.info_le.text())
        elif type == "IEC104":
            errcount = myprocess.IEC104Analyse()
        elif type == "modbusTCP":
            errcount = myprocess.modbusTCPAnalyse()

        if errcount != 0:
            for i in range(errcount):
                text = (f'错误数据[{i}]:序号{myprocess.err_list.ErrEleList[i].count}, 行号{myprocess.err_list.ErrEleList[i].row}, 错误码{myprocess.err_list.ErrEleList[i].errcode}')
                self.PrintText(text)
                print(text)
            text = "分析完成\n"
            self.PrintText(text)
            print(text)
        else:
            text = "分析完成，无错误\n"
            self.PrintText(text)
            print(text)

        self.browse_button2.setEnabled(True)
        self.browse_button2.setText("分析")
        #self.browse_button2.setStyleSheet("background-color:rgb(255, 255, 255)")


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MyWidget()
    window.show()
    sys.exit(app.exec())
