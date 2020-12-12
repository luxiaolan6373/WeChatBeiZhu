from PyQt5.QtWidgets import QFrame, QMainWindow, QTableWidgetItem, QPushButton
from ui.ui_kjy import Ui_Form
from ui.ui_main import Ui_MainWindow
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QPoint
from codes.moni import Mouse_And_Key
from PyQt5.QtCore import pyqtSignal,QObject
import pickle
import time,codes.qjbl

class Init_ui_kjy():
    def __init__(self, window: QMainWindow, ui: Ui_Form, main_Window: QMainWindow, ui_main: Ui_MainWindow):
        self.window = window
        self.ui = ui
        self.main_Window = main_Window
        self.ui_main = ui_main
        self.wechatHwnd = codes.qjbl.get_hwnd()
        self.isok = False
        self.path = "datas\\快捷语.txt"
        self.setPath = 'datas\\changed.pkl'
        self.set()

    def set(self):

        self.window.setWindowTitle('快捷语句')
        self.window.resize(self.window.width(),self.main_Window.height())
        self.window.move(self.main_Window.x()+self.main_Window.width(),self.main_Window.y())
        self.window.setWindowIcon(QIcon('logo.ico'))
        # 设置,位置,大小,窗口样式
        self.window.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.WindowCloseButtonHint |Qt.WindowMinimizeButtonHint)
        self.window.move(self.main_Window.x() + self.main_Window.window().width(), self.main_Window.y())
        self.window.resize(self.window.width(), self.main_Window.height())
        # 绑定信号
        self.ui.tabw_kjy.cellChanged.connect(self.cellChanged)
        self.ui.cb_enter.stateChanged[int].connect(self.cb_state_changed)
        self.ui.cb_ctrl_enter.stateChanged[int].connect(self.cb_state_changed)
        self.ui.bt_addLine.clicked.connect(self.add_line)
        self.ui.bt_subLine.clicked.connect(self.sub_line)

        # 读取设置
        try:
            with open(self.setPath, 'rb')as f:
                set_jl = pickle.load(f)
                self.ui.cb_enter.setChecked(set_jl['Enter键'])
                self.ui.cb_ctrl_enter.setChecked(set_jl['Ctrl_Enter键'])
        except:
            pass
        # 设置允许换行
        self.ui.tabw_kjy.setWordWrap(True)
        # 加载快捷语列表
        self.tab_init()

    def tab_init(self):
        with open(self.path, 'r') as f:
            text = f.read()
        ts = text.split('\n')
        # 设置显示的表格行和列数  最大
        self.ui.tabw_kjy.setRowCount(len(ts))
        self.ui.tabw_kjy.setColumnCount(2)
        self.ui.tabw_kjy.setColumnWidth(0, 200)
        self.ui.tabw_kjy.setColumnWidth(1, 65)
        self.ui.tabw_kjy.setHorizontalHeaderLabels(('快捷语', ''))
        for i, item in enumerate(ts):
            text_kjy = QTableWidgetItem(item)
            self.ui.tabw_kjy.setItem(i, 0, text_kjy)
            bt = QPushButton('发送')
            # 绑定发送按钮的信号
            bt.clicked.connect(lambda: self.send_message())
            self.ui.tabw_kjy.setCellWidget(i, 1, bt)
        self.isok = True


    def cellChanged(self, row, column):
        # 列大小根据内容调整大小
        self.ui.tabw_kjy.resizeRowsToContents()
        if self.isok == True:
            text = ''
            length = self.ui.tabw_kjy.rowCount()
            for i in range(length):
                try:
                    if i == length - 1:
                        text = text + self.ui.tabw_kjy.item(i, 0).text()
                    else:
                        text = text + self.ui.tabw_kjy.item(i, 0).text() + '\n'
                except:
                    pass
            with open(self.path, 'w') as f:
                f.write(text)

    def send_message(self):
        self.wechatHwnd=codes.qjbl.get_hwnd()
        print(self.wechatHwnd)
        x = self.window.sender().frameGeometry().x()
        y = self.window.sender().frameGeometry().y()
        # 利用按钮的指针,通过为止得到表格的item项目
        index = self.ui.tabw_kjy.indexAt(QPoint(x, y))
        row = index.row()
        message = self.ui.tabw_kjy.item(row, 0).text()
        message = self.message_formatting(message)
        # 创建模拟对象
        mn = Mouse_And_Key()
        # 置剪辑版文本
        mn.set_text(message)
        # 激活微信窗口
        mn.set_window_foregGroun(self.wechatHwnd)
        # 粘贴文本
        mn.key_even_zuhe(86, 17)  # ctrl+V
        time.sleep(0.1)
        if self.ui.cb_enter.isChecked() == True:
            mn.key_even(13)  # 回车
        elif self.ui.cb_ctrl_enter.isChecked() == True:
            mn.key_even_zuhe(13, 17)  # ctrl+回车

    def message_formatting(self, text):
        for i in range(self.ui_main.tabw_main.rowCount()):
            label = self.ui_main.tabw_main.item(i, 0).text()
            try:
                label_text = self.ui_main.tabw_main.item(i, 1).text()
            except:
                label_text = ''
            text = text.replace(f'[{label}]', label_text)
        return text

    def cb_state_changed(self, a):
        item = self.window.sender()
        if a == 2:
            if item.text() == 'Enter键':
                self.ui.cb_ctrl_enter.setChecked(False)
            else:
                self.ui.cb_enter.setChecked(False)
            '''保存设置字典'''
            # 保存字典到文件
        changed = {'Enter键': self.ui.cb_enter.checkState(), 'Ctrl_Enter键': self.ui.cb_ctrl_enter.checkState()}
        with open(self.setPath, 'wb')as f:
            pickle.dump(changed, f)

    def add_line(self):
        self.ui.tabw_kjy.setRowCount(self.ui.tabw_kjy.rowCount() + 1)
        bt = QPushButton('发送')
        # 绑定发送按钮的信号
        bt.clicked.connect(lambda: self.send_message())
        self.ui.tabw_kjy.setCellWidget(self.ui.tabw_kjy.rowCount() - 1, 1, bt)

    def sub_line(self):
        if self.ui.tabw_kjy.rowCount() > 0:
            self.ui.tabw_kjy.setRowCount(self.ui.tabw_kjy.rowCount() - 1)
    def __del__(self):
        print ('__del__')
