import os
import re
import sys

from PySide6 import QtCore, QtGui
from PySide6.QtCore import QEventLoop, QTimer
from colorama import Fore
from docx import Document
from docx.opc.exceptions import PackageNotFoundError

import main
from configparser import ConfigParser

from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog
from openpyxl.reader.excel import load_workbook

from form import Ui_MainWindow


class EmittingStr(QtCore.QObject):
    textWritten = QtCore.Signal(str)

    def write(self, text):
        self.textWritten.emit(str(text))
        loop = QEventLoop()
        QTimer.singleShot(100, loop.quit)
        loop.exec()
        QApplication.processEvents()


class MainWindow(Ui_MainWindow, QMainWindow):
    workdir = ''
    file_prj = ''
    file_pat = ''
    pat_dict = {}

    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)
        sys.stdout = EmittingStr()
        sys.stdout.textWritten.connect(self.outputWritten)
        sys.stderr = EmittingStr()
        sys.stderr.textWritten.connect(self.outputWritten)

        self.actionSelect_Dir.triggered.connect(self.setDocUrl)
        self.actioncheck.triggered.connect(self.checkpatent)
        self.actionreplace.triggered.connect(self.replaceprj)

        self.config = ConfigParser()
        try:
            self.config.read('config.ini', encoding='UTF-8')
            self.workdir = self.config['config']['lasting']
            self.onchangeworkdir()
        except:
            self.config.add_section('config')
            pass
        self.lineEdit.setText(self.workdir)

    def outputWritten(self, text):
        cursor = self.textEdit.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.textEdit.setTextCursor(cursor)
        self.textEdit.ensureCursorVisible()

    def setDocUrl(self):
        # 重新选择输入和输出目录时，进度条设置为0，文本框的内容置空
        tempdir = QFileDialog.getExistingDirectory(self, "选中项目所在目录", r"")
        if tempdir != '':
            self.workdir = tempdir
            self.onchangeworkdir()
            with open('config.ini', 'w', encoding='utf-8') as file:
                self.config['config']['lasting'] = self.workdir
                self.config.write(file)  # 数据写入配置文件

    def onchangeworkdir(self):
        self.lineEdit.setText(self.workdir)
        for file_sum in os.listdir(self.workdir):
            if file_sum.endswith('立项报告汇总表.xlsx') and not file_sum.startswith('~$'):
                self.file_prj = self.workdir + '/' + file_sum
            if file_sum.endswith('知识产权汇总表.xlsx') and not file_sum.startswith('~$'):
                self.file_pat = self.workdir + '/' + file_sum

    def replaceprj(self):
        wb = load_workbook(self.file_prj, data_only=True)
        ws = wb.active
        if str(ws['A1'].value).find(u'公司') != -1:
            com_name = str(ws['A1'].value).split("公司")[0] + '公司'
        else:
            print("Error: 找不到 公司名")
            return

        max_row_num = ws.max_row
        rangeCell = ws[f'A3:P{max_row_num}']
        for r in rangeCell:
            if r[0].value is None:
                break
            project = main.Project()
            project.p_comname = com_name
            project.p_order = str(r[0].value).strip().zfill(2)
            project.p_name = str(r[1].value).strip()
            project.p_start = r[2].value.strftime('%Y-%m-%d')
            project.p_end = r[3].value.strftime('%Y-%m-%d')
            project.p_cost = str(r[5].value).strip()
            project.p_people = str(r[6].value).strip()  # 人数
            project.p_owner = str(r[7].value).strip()  # 项目负责人
            project.p_rnd = str(r[8].value).strip()  # 研发人员
            project.p_money = str(r[9].value).strip()  # 总预算

            try:
                doc_name = self.workdir + '/RD' + project.p_order + project.p_name + '.docx'
                document = Document(doc_name)
                # debug_doc(document)
                # TODO to be fixed
                # replace_header(document)
                main.first_table(document, project)
                main.start_time(document, project)
                main.second_table(document, project)
                main.third_table(document, project)
                document.save(doc_name)
            except PackageNotFoundError:
                self.textEdit.append('Error打开文件错误：' + doc_name)

    def checkpatent(self):
        self.pat_dict.clear()
        wb = load_workbook(self.file_pat, read_only=True, data_only=True)
        ws = wb.active
        max_row_num = ws.max_row
        rangeCell = ws[f'A3:B{max_row_num}']
        for r in rangeCell:
            if r[0].value is None:
                break
            p_order = str(r[0].value).strip().zfill(2)
            p_name = str(r[1].value).strip()
            self.pat_dict[p_name] = p_order

        wb = load_workbook(self.file_prj)
        ws = wb.active
        max_row_num = ws.max_row
        rangeCell = ws[f'L3:P{max_row_num}']
        i: int = 0
        for r in rangeCell:
            i = i + 1
            if r[0].value is None:
                break
            p_name = str(r[0].value).strip()
            lst = p_name.splitlines()
            rep = [self.pat_dict[x] if x in self.pat_dict else x for x in lst]
            for element in rep:
                if re.search('\d\d', element) is None:
                    self.textEdit.append('Error没有找到专利：' + element)
            rep = map(lambda element: 'IP' + element, rep)
            new_ip = ';'.join(rep)
            if r[3].value != new_ip:
                self.textEdit.append(str(i) + ' 行: ' + str(r[3].value) + ' 替换为 ' + new_ip)
                r[3].value = new_ip
        try:
            wb.save(self.file_prj)
            self.textEdit.append('检查和更新完成.')
        except PermissionError:
            self.textEdit.append('写文件失败，关闭其他占用该文件的程序.' + self.file_prj)


def exceptOutConfig(exctype, value, tb):
    print('My Error Information:')
    print('Type:', exctype)
    print('Value:', value)
    print('Traceback:', tb)


if __name__ == "__main__":
    sys.excepthook = exceptOutConfig
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
