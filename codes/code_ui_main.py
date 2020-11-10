from PyQt5.QtWidgets import QMainWindow, qApp,QTableWidgetItem,QFrame
from ui.ui_main import Ui_MainWindow
from PyQt5.QtCore import Qt
from PyQt5.Qt import QCursor
from codes.xlsxCaoZuo import KeHuGuanLi
from codes.baiduOrz import Baiduorc
import time,win32gui,os
class Init_window_main:
    def __init__(self, window: QMainWindow, ui: Ui_MainWindow):
        self.window = window
        self.ui = ui
        self.desktopWidth=qApp.desktop().width()
        self.screen = qApp.primaryScreen()
        self.x =self.desktopWidth-2#位置在窗口最右边隐藏起来
        self.y = 0
        self.row=0
        self.hwnd=0#窗口句柄
        self.zzyx = False#标记正在走特效,防止干扰
        self.b = 0
        self.column=0
        self.rowSelection=0
        # 创建客户表格管理对象
        self.excelPath='datas/客户信息.xlsx'
        self.kh = KeHuGuanLi(self.excelPath)
        # 获取最大行列数量
        self.row = self.kh.getRow()
        self.column = self.kh.getColumn()
        self.set()
    def set(self):
        # 设置标题
        self.window.setWindowTitle('客户备注')

        self.window.move(self.x, self.y)
        self.window.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)
        # 重写事件---主界面的动画效果
        self.window.enterEvent = self.enterEvent
        self.window.leaveEvent = self.leaveEvent
        self.ui.tabw_main.cellChanged.connect(self.cellChanged)
        self.ui.bt_save.clicked.connect(self.save)
        self.ui.bt_excel.clicked.connect(lambda: os.system(f"start {os.getcwd()}/{self.excelPath}"))
        self.ui.bt_excel.setToolTip('你可以打开excel表格然后多增加几个标签,然后修改后重开软件')

        # 水平表格头显示和隐藏
        self.ui.tabw_main.horizontalHeader().setVisible(False)
        # 垂直表格头显示和隐藏
        self.ui.tabw_main.verticalHeader().setVisible(False)
        # 设置允许换行
        self.ui.tabw_main.setWordWrap(True)
        self.ui.tabw_main.resizeRowsToContents()
        self.ui.tabw_main.resizeColumnsToContents()

        #设置显示的表格行和列数  最大
        self.ui.tabw_main.setRowCount(self.column)
        self.ui.tabw_main.setColumnCount(2)
        #将所有标签读取出来

        for i,item in enumerate(self.kh.getLabel()):
            labelItem = QTableWidgetItem(item)
            self.ui.tabw_main.setItem(i, 0, labelItem)
        # 配置百度文字识别
        # 570 1080-100 200 https://console.bce.baidu.com/ai/#/ai/ocr/overview/index
        # client_id 为官网获取的AK， client_secret 为官网获取的SK
        AK = "Weui83nXBM1ox6ozFzPF9bng"
        SK = "4AEmIGiInMv7gzlATntAs3pjHZrlCrsK"
        self.bd = Baiduorc(AK="Weui83nXBM1ox6ozFzPF9bng", SK="4AEmIGiInMv7gzlATntAs3pjHZrlCrsK")
        self.imagePath = "datas\\dqyh.png"
    def getWeChatNameNow(self):
        try:
            # 查找微信的窗口句柄
            if self.hwnd==0:
                self.hwnd = win32gui.FindWindow('WeChatMainWndForPC', '微信')
            img = self.screen.grabWindow(self.hwnd , 329, 23, 350, 30).toImage()
            img.save(self.imagePath)
            jstext = self.bd.get_text(imagePath=self.imagePath)  # 调用百度识字
            print(jstext)
            NameNow=jstext['words_result'][0]['words']
            # 识图结果为空跳过
            if NameNow=='':
                return
            info=[]
            for i, item in enumerate(self.kh.ws.rows):
                if i == 0:#跳过第一行 是标签
                    continue

                if item[0].value == NameNow:#找到这个微信数据
                    self.rowSelection=i+1
                    print(i)
                    info=item
                    break
            self.loadInfo(info,NameNow)
        except:
            pass
        self.ui.bt_save.setEnabled(False)
    def cellChanged(self,row,column):
        # 行列大小根据内容调整大小
        self.ui.tabw_main.resizeRowsToContents()
        self.ui.tabw_main.resizeColumnsToContents()
        self.ui.bt_save.setEnabled(True)
    def loadInfo(self,item,NameNow):
        '''
        根据表格某一行的数据,生成控件
        :param item:
        :return:
        '''
        print(111)
        self.ui.tabw_main.clearContents()
        if len(item)==0:
            labelItem = QTableWidgetItem(NameNow)
            #在最下面添加一行数据,并且把整个id记下
            self.kh.ws.append([NameNow, ])
            self.row = self.kh.getRow()
            self.rowSelection =self.row
            self.ui.tabw_main.setItem(0, 1, labelItem)
        # 将所有标签读取出来
        for i, iItem in enumerate(self.kh.getLabel()):
            labelItem = QTableWidgetItem(iItem)
            self.ui.tabw_main.setItem(i, 0, labelItem)
        # 测试读取某个微信名的客户数据
        for j, jitem in enumerate(item):
            if jitem.value == None :
                dataItem = QTableWidgetItem('')
            else:
                dataItem = QTableWidgetItem(str(jitem.value))
            self.ui.tabw_main.setItem(j, 1, dataItem)
        self.ui.tabw_main.resizeRowsToContents()
        self.ui.tabw_main.resizeColumnsToContents()
    def enterEvent(self, event):
        '''
        鼠标进入主界面时触发
        :param event:
        :return:
        '''

        if self.window.x()>self.desktopWidth-self.window.width()+2 and self.zzyx == False:

            self.getWeChatNameNow()
            for i in range(self.window.width()):
                self.zzyx = True
                self.window.move(self.x - i + 1, self.y)
                time.sleep(0.001)
            self.zzyx = False
        print('进入')
    def leaveEvent(self, event):
        '''
        鼠标离开主界面时触发
        :param event:
        :return:
        '''
        #防止点不了标题栏
        if QCursor.pos().y()>30 and self.zzyx == False:
            x=self.window.x()
            for i in range(self.window.width()):
                self.zzyx = True
                self.window.move(x + i - 1, self.y)
                time.sleep(0.001)
            self.zzyx = False
            print('离开')
    def save(self):


        for i in range(self.ui.tabw_main.rowCount()):
            ch = chr(i + 65)
            print(f'{ch}{self.rowSelection}')
            try:

                self.kh.ws[f'{ch}{self.rowSelection}'] = self.ui.tabw_main.item(i, 1).text()
            except:
                self.kh.ws[f'{ch}{self.rowSelection}'] =''
        #保存
        self.kh.save


