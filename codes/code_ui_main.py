from PyQt5.QtWidgets import QMainWindow, qApp, QTableWidgetItem, QFrame, qApp,QMessageBox
from ui.ui_main import Ui_MainWindow
from PyQt5.QtCore import Qt
from PyQt5.Qt import QCursor, QIcon
from PyQt5.QtGui import QBrush, QColor
from codes.xlsxCaoZuo import KeHuGuanLi
from codes.baiduOrz import Baiduorc
import time, win32gui, os, copy, threading, win32con
from PyQt5.QtCore import pyqtSignal, QObject
from ui.ui_kjy import Ui_Form
from codes.code_ui_kjy import Init_ui_kjy
import sip, codes.qjbl, os, time, shutil


class MySignal(QObject):
    move_sg = pyqtSignal(int, int, int, int, bool)


class Init_window_main:
    def __init__(self, window: QMainWindow, ui: Ui_MainWindow):
        self.window = window
        self.window_kjyj = 0
        self.window_move_sg = 0
        self.kg_kjyj = False
        self.window.setAttribute(Qt.WA_DeleteOnClose, True)

        self.ui = ui
        self.desktopWidth = qApp.desktop().width()
        self.screen = qApp.primaryScreen()
        self.x = self.desktopWidth - 2  # 位置在窗口最右边隐藏起来
        self.img_2 = 0
        self.y = 0
        self.row = 0
        self.hwnd = codes.qjbl.get_hwnd()  # 窗口句柄
        self.zzyx = False  # 标记正在走特效,防止干扰
        self.b = 0
        self.column = 0
        self.rowSelection = 0
        # 创建客户表格管理对象
        self.excelPath = 'datas/客户信息.xlsx'
        self.kh = KeHuGuanLi(self.excelPath)
        # 获取最大行列数量
        self.row = self.kh.getRow()
        self.column = self.kh.getColumn()
        self.set()

    def set(self):
        # 设置标题
        self.window.setWindowTitle('客户备注管理系统')
        self.window.setWindowIcon(QIcon('./logo.ico'))

        self.window.move(self.x, self.y)
        self.window.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.WindowCloseButtonHint)
        # 自定义信号
        self.move_sg = MySignal()
        # 绑定信号
        self.move_sg.move_sg.connect(self.move_window)
        self.ui.tabw_main.cellChanged.connect(self.cellChanged)
        self.ui.bt_excel.clicked.connect(lambda: os.startfile(f'{os.getcwd()}\\{self.excelPath}'))
        self.ui.bt_kjyj.clicked.connect(self.open_kjyj)
        self.window.closeEvent = self.close

        self.ui.bt_excel.setToolTip('你可以打开excel表格然后多增加几个标签,然后修改后重开软件')
        # 水平表格头显示和隐藏
        self.ui.tabw_main.horizontalHeader().setVisible(False)
        # 垂直表格头显示和隐藏
        self.ui.tabw_main.verticalHeader().setVisible(False)
        # 设置允许换行
        self.ui.tabw_main.setWordWrap(True)
        self.ui.tabw_main.resizeRowsToContents()
        self.ui.tabw_main.resizeColumnsToContents()

        # 设置显示的表格行和列数  最大
        self.ui.tabw_main.setRowCount(self.column)
        self.ui.tabw_main.setColumnCount(2)
        # 将所有标签读取出来

        for i, item in enumerate(self.kh.getLabel()):
            labelItem = QTableWidgetItem(item)
            self.ui.tabw_main.setItem(i, 0, labelItem)
        # 配置百度文字识别
        # 570 1080-100 200 https://console.bce.baidu.com/ai/#/ai/ocr/overview/index
        # client_id 为官网获取的AK， client_secret 为官网获取的SK
        AK = "Weui83nXBM1ox6ozFzPF9bng"
        SK = "4AEmIGiInMv7gzlATntAs3pjHZrlCrsK"
        self.bd = Baiduorc(AK="Weui83nXBM1ox6ozFzPF9bng", SK="4AEmIGiInMv7gzlATntAs3pjHZrlCrsK")
        self.imagePath = "datas\\dqyh.png"
        self.windowHwind = int(self.window.winId())
        self.window_kjyj_Hwind = 0
        t = threading.Thread(target=self.getWeChatNameNow)
        t.setDaemon(True)
        t.start()


    def close(self, enent):
        try:
            # 备份
            # 判断是否有这个目录,没有的话就建立
            if os.path.isdir(self.standby_path) != True:
                os.makedirs(self.standby_path)
            shutil.copyfile(self.excelPath, f'{self.standby_path}\\{int(time.time())}.xlsx')
        except:
            pass
        qApp.exit()

    def move_window(self, x, y, w, h, vis):

        if vis == False:
            if self.window.isEnabled() == True:
                self.window.setVisible(vis)
        else:
            if self.window.x() != x + w or self.window.y() != y or self.window.height() != h - 30:
                self.window.move(x + w, y)
                self.window.resize(self.window.width(), h - 30)

            self.window.setVisible(vis)

    def getWeChatNameNow(self):
        self.hwnd_kjyj = 0
        sfjh = True
        change = False

        while True:
            try:
                xx, yy, ww, hh = 0, 0, 0, 0
                # 查找窗口句柄会优先获得最前的句柄
                hwnd = win32gui.GetForegroundWindow()
                # 判断是否为支持的窗口
                if hwnd == 0 :
                    continue

                title = win32gui.GetWindowText(hwnd)
                class_ = win32gui.GetClassName(hwnd)
                if title == '微信' and class_ == 'WeChatMainWndForPC':  # 目标为微信
                    self.hwnd = hwnd
                    codes.qjbl.set_hwnd(self.hwnd)
                    windowPost = [329, 23, 350, 30]
                elif title == '企业微信' and class_ == 'WeWorkWindow':  # 目标为微信
                    self.hwnd = hwnd
                    codes.qjbl.set_hwnd(self.hwnd)
                    windowPost = [317,25,350,30]
                elif title.find('阿里旺旺') != -1 and class_ == 'StandardFrame':  # 目标为阿里
                    self.hwnd = hwnd
                    codes.qjbl.set_hwnd(self.hwnd)
                    windowPost = [161, 18, 360, 45]
                elif title != 'QQ' and class_ == 'TXGuiFoundation':  # 目标为qq
                    self.hwnd = hwnd
                    codes.qjbl.set_hwnd(self.hwnd)
                    windowPost = None
                    xx, yy, ww, hh = 0, 6, -8, -14
                elif title == '钉钉' and class_ == 'StandardFrame_DingTalk':  # 目标为钉钉
                    self.hwnd = hwnd
                    # 得到目标窗口的矩阵信息
                    rect = win32gui.GetWindowRect(self.hwnd)
                    x, y, w, h = rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]
                    codes.qjbl.set_hwnd(self.hwnd)
                    windowPost = [374, 37, 272, 25]
                elif hwnd == self.windowHwind :  # 目标为本体,不操作
                    continue
                else:
                    # 得到目标窗口的矩阵信息
                    rect = win32gui.GetWindowRect(self.hwnd)
                    x, y, w, h = rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]
                    # 触发信号,把信息传递过去
                    self.move_sg.move_sg.emit(x + xx, y + yy, w + ww, h + hh, False)

                    continue
                # 得到目标窗口的矩阵信息
                rect = win32gui.GetWindowRect(self.hwnd)
                x, y, w, h = rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]
                # 触发信号,把信息传递过去
                self.move_sg.move_sg.emit(x + xx, y + yy, w + ww, h + hh, True)
                if windowPost != None:
                    # 截图
                    img = self.screen.grabWindow(self.hwnd, windowPost[0], windowPost[1], windowPost[2],
                                                 windowPost[3]).toImage()
                    img.save(self.imagePath)
                    with open(self.imagePath, 'rb')as f:
                        img = f.read()
                    if img != self.img_2 or change == True:
                        try:
                            jstext = self.bd.get_text(image=img)  # 调用百度识字
                            print(jstext)
                            self.img_2 = copy.copy(img)
                            NameNow = jstext['words_result'][0]['words']
                            # 识图结果为空跳过
                            if NameNow == '':
                                return
                            NameNow = NameNow.replace('[淘宝网会员]', '').replace('②', '').replace('[宝网会员]', '').replace(
                                '遍宝网会员]量', '')
                            info = []
                            for i, item in enumerate(self.kh.ws.rows):
                                if i == 0:  # 跳过第一行 是标签
                                    continue
                                if item[0].value == NameNow:  # 找到这个微信数据
                                    self.rowSelection = i + 1
                                    info = item
                                    break
                            self.loadInfo(info, NameNow)

                        except Exception as err:
                            print(err, '识图直接报错了')
                        finally:
                            change = False
                else:
                    change = True
                    NameNow = title.split('等')[0]
                    print(NameNow)
                    info = []
                    for i, item in enumerate(self.kh.ws.rows):
                        if i == 0:  # 跳过第一行 是标签
                            continue

                        if item[0].value == NameNow:  # 找到这个微信数据
                            self.rowSelection = i + 1
                            print(i)
                            info = item
                            break
                    self.loadInfo(info, NameNow)
            except Exception as err:
                print(err)
            time.sleep(0.3)

    def open_kjyj(self):

        if self.kg_kjyj==True:
            #将控件单独出来,然后才能删除
            self.window_kjyj.setParent(None)

            width=self.window_kjyj.width()
            self.ui.hbox_main.removeWidget(self.window_kjyj)
            self.window.resize(570-width, self.window.height())
            self.kg_kjyj = False
        else:
            self.window_kjyj = QFrame()
            self.ui_kjyj = Ui_Form()
            self.ui_kjyj.setupUi(self.window_kjyj)
            init = Init_ui_kjy(self.window_kjyj, self.ui_kjyj, self.window, self.ui)
            self.ui.hbox_main.addWidget(self.window_kjyj)
            self.window.resize(570,self.window.height())
            self.window_kjyj.show()
            self.kg_kjyj=True
    def cellChanged(self, row, column):
        # 行列大小根据内容调整大小
        self.ui.tabw_main.resizeRowsToContents()
        self.ui.tabw_main.resizeColumnsToContents()
        self.set_colour(self.ui.tabw_main.item(row,column).text(), self.ui.tabw_main.item(row,column))
        self.save()
    def loadInfo(self, item, NameNow):
        '''
        根据表格某一行的数据,生成控件
        :param item:
        :return:
        '''
        self.ui.tabw_main.clearContents()
        if len(item) == 0:
            labelItem = QTableWidgetItem(NameNow)
            # 在最下面添加一行数据,并且把整个id记下
            self.kh.ws.append([NameNow, ])
            self.row = self.kh.getRow()
            self.rowSelection = self.row
            self.ui.tabw_main.setItem(0, 1, labelItem)
        # 将所有标签读取出来
        for i, iItem in enumerate(self.kh.getLabel()):
            labelItem = QTableWidgetItem(iItem)
            self.ui.tabw_main.setItem(i, 0, labelItem)
        # 测试读取某个微信名的客户数据
        for j, jitem in enumerate(item):
            if jitem.value == None:
                dataItem = QTableWidgetItem('')
            else:

                dataItem = QTableWidgetItem(str(jitem.value))

            self.ui.tabw_main.setItem(j, 1, dataItem)
            self.set_colour(jitem.value, dataItem)
        self.ui.tabw_main.resizeRowsToContents()
        self.ui.tabw_main.resizeColumnsToContents()
    def set_colour(self, t, dataItem):
        t=str(t)
        colours = []
        colours.append({'key': '#红色', 'col': [220, 20, 60]})
        colours.append({'key': '#粉色', 'col': [255, 20, 147]})
        colours.append({'key': '#蓝色', 'col': [0, 0, 255]})
        colours.append({'key': '#绿色', 'col': [0, 255, 0]})
        colours.append({'key': '#橙色', 'col': [255, 165, 0]})
        colours.append({'key': '#黄色', 'col': [255, 215, 0]})
        colours.append({'key': '#青色', 'col': [0, 255, 255]})
        colours.append({'key': '#茶色', 'col': [205, 133, 63]})
        colours.append({'key': '#草色', 'col': [127, 255, 0]})
        for item in colours:
            print(t.find(item['key']))
            if t.find(item['key']) != -1:
                dataItem.setForeground(QBrush(QColor(item['col'][0],item['col'][1], item['col'][2])))
    def save(self):
        try:
            for i in range(self.ui.tabw_main.rowCount()):
                ch = chr(i + 65)
                try:
                    self.kh.ws[f'{ch}{self.rowSelection}'] = self.ui.tabw_main.item(i, 1).text()

                except:
                    self.kh.ws[f'{ch}{self.rowSelection}'] = ''
            # 保存
            self.kh.save()
        except Exception as err:
            print(err)
