#!/usr/bin/env python
# encoding: utf-8
import sys
import time
from pathlib import Path
from utils.compare import Compare
from utils.message import Message
from utils.parsing_word_document import ParsingWord
from tqdm import tqdm
from utils.log import LoggerFactory
from utils.exists_mkdir import ExistsMkDir
from utils.parsing_excel import ParingExcel
from utils.glo import Globals
from PyQt5 import QtCore, QtGui, QtWidgets
_path = Path(__file__).resolve().parent
glo = Globals()
log = LoggerFactory('info')
glo._init()
glo.set_value('_path', _path)
log.create_logger()


class Ui_MainWindow(QtWidgets.QMainWindow):
    m_singal = QtCore.pyqtSignal(str)
    def __init__(self, parent=None):
        super(Ui_MainWindow, self).__init__(parent)
        self.setupUi(self)
        self.m_singal.connect(self.show_msg)

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1127, 913)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.word_radio = QtWidgets.QRadioButton(self.centralwidget)
        self.word_radio.setObjectName("word_radio")
        self.word_radio.toggled.connect(lambda: self.radio_state(self.word_radio))
        self.horizontalLayout.addWidget(self.word_radio)
        self.excel_radio = QtWidgets.QRadioButton(self.centralwidget)
        self.excel_radio.setObjectName("excel_radio")
        self.excel_radio.toggled.connect(lambda: self.radio_state(self.excel_radio))
        self.horizontalLayout.addWidget(self.excel_radio)
        self.txt_radio = QtWidgets.QRadioButton(self.centralwidget)
        self.txt_radio.setObjectName("txt_radio")
        self.txt_radio.toggled.connect(lambda: self.radio_state(self.txt_radio))
        self.horizontalLayout.addWidget(self.txt_radio)
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(lambda: self.display())
        self.horizontalLayout.addWidget(self.pushButton)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setMouseTracking(False)
        self.tabWidget.setTabletTracking(False)
        self.tabWidget.setAcceptDrops(False)
        self.tabWidget.setAutoFillBackground(False)
        self.tabWidget.setTabPosition(QtWidgets.QTabWidget.North)
        self.tabWidget.setObjectName("tabWidget")
        self.tab_log = QtWidgets.QWidget()
        self.tab_log.setObjectName("tab_log")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.tab_log)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.textEdit_log = QtWidgets.QTextEdit(self.tab_log)
        self.textEdit_log.setObjectName("textEdit_log")
        self.textEdit_log.isReadOnly()
        self.verticalLayout_3.addWidget(self.textEdit_log)
        self.tabWidget.addTab(self.tab_log, "")
        self.tab_report = QtWidgets.QWidget()
        self.tab_report.setObjectName("tab_report")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.tab_report)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.tableView = QtWidgets.QTableView(self.tab_report)
        self.tableView.setObjectName("tableView")
        self.verticalLayout_2.addWidget(self.tableView)
        self.tabWidget.addTab(self.tab_report, "")
        self.verticalLayout.addWidget(self.tabWidget)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1127, 22))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.word_radio.setText(_translate("MainWindow", "word"))
        self.excel_radio.setText(_translate("MainWindow", "excel"))
        self.txt_radio.setText(_translate("MainWindow", "txt"))
        self.pushButton.setText(_translate("MainWindow", "开始比对"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_log), _translate("MainWindow", "执行日志"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_report), _translate("MainWindow", "测试报告"))

    def show_report(self, data_list):
        self.QIM = QtGui.QStandardItemModel()
        self.QIM.clear()
        self.tableView.setModel(self.QIM)
        if data_list:
            if parsing == 'word':
                columns=['原始报告', '最新报告', '结果文件', '新增章节数', '删除章节数', '删除表数量', '段落差异数', '单元格差异数', '页眉差异数', '结果']
            elif parsing == 'excel':
                columns=['原始报告', '最新报告', '结果文件', '新增sheet数', '删除sheet数', '单元格差异', '结果']
            elif parsing == 'txt':
                columns=['原始报告', '最新报告', '结果文件', '新增行', '删除行', '行数据差异', '结果']
            for col, header in enumerate(columns):
                self.QIM.setItem(0, col, QtGui.QStandardItem(header))
            for row, z in enumerate(data_list):
                row += 1
                for x, y in enumerate(z):
                    self.QIM.setItem(row, x, QtGui.QStandardItem(str(y)))
            self.tableView.setModel(self.QIM)
            self.tableView.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        else:
            pass

    def show_msg(self, msg):
        self.textEdit_log.moveCursor(QtGui.QTextCursor.End)
        self.textEdit_log.append(msg)

    def radio_state(self, radio):
        if radio.isChecked() == True:
            global parsing
            parsing = radio.text()
            glo.set_value('parsing', parsing)

    def display(self):
        self.textEdit_log.clear()
        data = []
        if self.word_radio.isChecked() or self.excel_radio.isChecked() or self.txt_radio.isChecked():
            open(f"{_path}/log/{time.strftime('%Y-%m-%d', time.localtime())}.log", 'w').close()
            self.m_singal.emit(f"开始比对{time.asctime( time.localtime(time.time()))}")
            ExistsMkDir().exists_mk_dir() # 初始化文件
            main = TestMain()
            if parsing == 'word':
                rec_code = QtWidgets.QMessageBox.question(self, "友情提示", "请选择word解析程序：Yes=office, No=wps", QtWidgets.QMessageBox.Yes|QtWidgets.QMessageBox.No)
                if rec_code == QtWidgets.QMessageBox.No:
                    app = 'WPS'
                else:
                    app = 'OFFICE'
                log.logger.info(f"请选择办公软件(office or wps)：{app}")
                log.logger.info("开始转换doc_to_docx".center(50 // 2, "-"))
                main.test_convert_doc_to_docx(app)
                data = main.test_compare_word()
            elif parsing == 'excel':
                data = main.test_compare_excel()
            elif parsing == 'txt':
                data = main.test_compare_txt()
            with open(f"{_path}/log/{time.strftime('%Y-%m-%d', time.localtime())}.log", "r", encoding="utf-8") as lines:
                array=lines.readlines()
                for i in array:
                    i=i.strip('\n')
                    self.m_singal.emit(i)
            self.m_singal.emit(f"比对结束{time.asctime( time.localtime(time.time()))}")
            self.show_report(data)
        else:
            QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning, '警告', '未选择解析格式').exec_()
            self.m_singal.emit("未选择解析格式")

class TestMain():
    def __init__(self,):
        self.compare = Compare()
        self.message = Message()
        self.parsing_word = ParsingWord()
        self.parsing_excel = ParingExcel()

    def test_compare_word(self):
        data_lists = []
        compare_lists = self.parsing_word.get_file_path_tuple('.docx') # 获取文件路径元组
        for compare in compare_lists:
            begin_filename, end_filename = compare
            msg = self.compare.compare_deepdiff_word(begin_filename, end_filename) # 获取比对结果
            data_lists.append(msg)
        self.parsing_excel.write_excel(data_lists, parsing)
        return data_lists

    def test_compare_excel(self):
        data_lists = []
        compare_lists = self.parsing_word.get_file_path_tuple('.xls')
        for compare in compare_lists:
            begin_filename, end_filename = compare
            msg = self.compare.compare_deepdiff_excel(begin_filename, end_filename) # 获取比对结果
            data_lists.append(msg)
        self.parsing_excel.write_excel(data_lists, parsing)
        return data_lists

    def test_compare_txt(self):
        data_lists = []
        compare_lists = self.parsing_word.get_file_path_tuple('.txt')
        for compare in compare_lists:
            begin_filename, end_filename = compare
            msg = self.compare.compare_deepdiff_txt(begin_filename, end_filename) # 获取比对结果
            data_lists.append(msg)
        self.parsing_excel.write_excel(data_lists, parsing)
        return data_lists

    def test_convert_doc_to_docx(self, app):
        if app == 'OFFICE':
            sum_convert = 1
        else:
            sum_convert = 2
        for i in tqdm(range(sum_convert)): #wps情况下 doc另存为docx第一次会出现无法找到文件的问题，暂时无法解决，目前处理方式重复存一次
            self.parsing_word.convert_doc_to_docx(app)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    ui = Ui_MainWindow()
    ui.setWindowTitle("自动化比对工具")
    ui.show()
    sys.exit(app.exec_())