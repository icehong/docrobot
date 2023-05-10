import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog
from form import Ui_MainWindow


class MainWindow(Ui_MainWindow, QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)

        self.actionSelect_Dir.triggered.connect(self.setDocUrl)

    def setDocUrl(self):
        # 重新选择输入和输出目录时，进度条设置为0，文本框的内容置空
        str = QFileDialog.getExistingDirectory(self, "选中项目所在目录", r"")
        self.lineEdit.setText(str)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
