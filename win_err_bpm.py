from PyQt5 import QtCore, QtGui, QtWidgets
import sys


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(344, 239)
        MainWindow.setMinimumSize(QtCore.QSize(344, 239))
        MainWindow.setMaximumSize(QtCore.QSize(344, 239))
        MainWindow.setWindowTitle("sm_bpm")
        MainWindow.setStyleSheet("background-color: rgb(87, 114, 173);")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(80, 200, 191, 31))
        self.pushButton.setMinimumSize(QtCore.QSize(191, 31))
        self.pushButton.setMaximumSize(QtCore.QSize(191, 31))
        self.pushButton.setStyleSheet("\n"
                                      "font: 75 10pt \"Arial\";\n"
                                      "background-color: rgb(210, 215, 255);\n"
                                      "font: 9pt \"Arial\";")
        self.pushButton.setAutoDefault(False)
        self.pushButton.setDefault(False)
        self.pushButton.setFlat(False)
        self.pushButton.setObjectName("pushButton")
        # self.pushButton.clicked.connect(self.close_prog)
        self.textBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser.setGeometry(QtCore.QRect(0, 0, 341, 191))
        self.textBrowser.setMinimumSize(QtCore.QSize(341, 191))
        self.textBrowser.setMaximumSize(QtCore.QSize(341, 191))
        self.textBrowser.setStyleSheet("font: 87 12pt \"Arial\";")
        self.textBrowser.setObjectName("textBrowser")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def close_prog(self):
        sys.exit()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        self.pushButton.setText(_translate("MainWindow", "Закрыть"))
        self.textBrowser.setHtml(_translate("MainWindow",
                                            "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
                                            "<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
                                            "p, li { white-space: pre-wrap; }\n"
                                            "</style></head><body style=\" font-family:\'Arial\'; font-size:12pt; font-weight:80; font-style:normal;\">\n"
                                            "<p align=\"center\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400;\"><br /></p>\n"
                                            "<p align=\"center\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'MS Shell Dlg 2\'; font-weight:400;\"><br /></p>\n"
                                            "<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\'; font-weight:400;\">Работа приложения завершена с ошибкой! </span></p>\n"
                                            "<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\'; font-weight:400;\">Попробуйте запустить еще раз.</span></p></body></html>"))

    def run_win(self):
        app = QtWidgets.QApplication(sys.argv)
        app.setStyle('Fusion')

        Window = QtWidgets.QMainWindow()
        Window.setWindowFlags(QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowStaysOnTopHint)
        ui = Ui_MainWindow()
        ui.setupUi(Window)
        Window.show()

        ui.pushButton.clicked.connect(self.close_prog)

        sys.exit(app.exec_())


if __name__ == "__main__":
    Ui_MainWindow().run_win()
