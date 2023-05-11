import sys
from configparser import ConfigParser

from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog
from form import Ui_MainWindow


class MainWindow(Ui_MainWindow, QMainWindow):
    workdir = ''

    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)

        self.actionSelect_Dir.triggered.connect(self.setDocUrl)
        self.actioncheck.triggered.connect(self.checkpatent)

        self.config = ConfigParser()
        try:
            self.config.read('config.ini', encoding='UTF-8')
            self.workdir = self.config['config']['lasting']
        except:
            self.config.add_section('config')
            pass
        self.lineEdit.setText(self.workdir)

    def setDocUrl(self):
        # 重新选择输入和输出目录时，进度条设置为0，文本框的内容置空
        tempdir = QFileDialog.getExistingDirectory(self, "选中项目所在目录", r"")
        if tempdir != '':
            self.workdir = tempdir
            self.lineEdit.setText(self.workdir)
            with open('config.ini', 'w', encoding='utf-8') as file:
                self.config['config']['lasting'] = self.workdir
                self.config.write(file)  # 数据写入配置文件

    def checkpatent(self):
        # 重新选择输入和输出目录时，进度条设置为0，文本框的内容置空
        str = QFileDialog.getExistingDirectory(self, "选中项目所在目录", r"")
        self.lineEdit.setText(str)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
