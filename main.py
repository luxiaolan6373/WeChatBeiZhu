from PyQt5.QtWidgets import QApplication,QMainWindow
import ui
from codes.code_ui_main import Init_window_main
import sys,threading,copy
if __name__ == '__main__':
    app=QApplication(sys.argv)
    window_main=QMainWindow()#主界面
    ui_main=ui.ui_main.Ui_MainWindow()#实例化
    ui_main.setupUi(window_main)
    Init_window_main(window_main,ui_main)
    with open('qss\main.qss','r')as f:
        style=f.read()
    window_main.setStyleSheet(style)
    window_main.show()
    sys.exit(app.exec_())