# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'form.ui'
##
## Created by: Qt User Interface Compiler version 6.4.1
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QAction, QBrush, QColor, QConicalGradient,
    QCursor, QFont, QFontDatabase, QGradient,
    QIcon, QImage, QKeySequence, QLinearGradient,
    QPainter, QPalette, QPixmap, QRadialGradient,
    QTransform)
from PySide6.QtWidgets import (QApplication, QLabel, QLineEdit, QMainWindow,
    QMenu, QMenuBar, QSizePolicy, QStatusBar,
    QTextEdit, QWidget)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(800, 600)
        self.actionSelect_Dir = QAction(MainWindow)
        self.actionSelect_Dir.setObjectName(u"actionSelect_Dir")
        self.actioncheck = QAction(MainWindow)
        self.actioncheck.setObjectName(u"actioncheck")
        self.actionreplace = QAction(MainWindow)
        self.actionreplace.setObjectName(u"actionreplace")
        self.actioncheckall = QAction(MainWindow)
        self.actioncheckall.setObjectName(u"actioncheckall")
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.textEdit = QTextEdit(self.centralwidget)
        self.textEdit.setObjectName(u"textEdit")
        self.textEdit.setGeometry(QRect(10, 40, 781, 521))
        self.label = QLabel(self.centralwidget)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(10, 10, 91, 16))
        self.lineEdit = QLineEdit(self.centralwidget)
        self.lineEdit.setObjectName(u"lineEdit")
        self.lineEdit.setGeometry(QRect(100, 10, 691, 20))
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName(u"menubar")
        self.menubar.setGeometry(QRect(0, 0, 800, 26))
        self.menu = QMenu(self.menubar)
        self.menu.setObjectName(u"menu")
        self.menu_2 = QMenu(self.menubar)
        self.menu_2.setObjectName(u"menu_2")
        self.menu_3 = QMenu(self.menubar)
        self.menu_3.setObjectName(u"menu_3")
        self.menu_4 = QMenu(self.menubar)
        self.menu_4.setObjectName(u"menu_4")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.menubar.addAction(self.menu.menuAction())
        self.menubar.addAction(self.menu_4.menuAction())
        self.menubar.addAction(self.menu_2.menuAction())
        self.menubar.addAction(self.menu_3.menuAction())
        self.menu.addAction(self.actionSelect_Dir)
        self.menu_2.addAction(self.actionreplace)
        self.menu_3.addAction(self.actioncheck)
        self.menu_4.addSeparator()
        self.menu_4.addAction(self.actioncheckall)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"MainWindow", None))
        self.actionSelect_Dir.setText(QCoreApplication.translate("MainWindow", u"Select Dir", None))
        self.actioncheck.setText(QCoreApplication.translate("MainWindow", u"\u68c0\u67e5", None))
        self.actionreplace.setText(QCoreApplication.translate("MainWindow", u"\u6279\u91cf\u66ff\u6362", None))
        self.actioncheckall.setText(QCoreApplication.translate("MainWindow", u"\u68c0\u67e5\u6240\u6709\u9879", None))
        self.label.setText(QCoreApplication.translate("MainWindow", u"\u5f53\u524d\u5904\u7406\u76ee\u5f55\uff1a", None))
        self.menu.setTitle(QCoreApplication.translate("MainWindow", u"\u8bbe\u7f6e", None))
        self.menu_2.setTitle(QCoreApplication.translate("MainWindow", u"\u9879\u76ee", None))
        self.menu_3.setTitle(QCoreApplication.translate("MainWindow", u"\u4e13\u5229", None))
        self.menu_4.setTitle(QCoreApplication.translate("MainWindow", u"\u68c0\u67e5", None))
    # retranslateUi

