from PyQt5.QtWidgets import QMainWindow, qApp,QTableWidgetItem,QFrame,qApp
from ui.ui_main import Ui_MainWindow
from PyQt5.QtCore import Qt
from PyQt5.Qt import QCursor,QIcon
from codes.xlsxCaoZuo import KeHuGuanLi
from codes.baiduOrz import Baiduorc
import time,win32gui,os,copy,threading,win32con
from PyQt5.QtCore import pyqtSignal,QObject
from ui.ui_kjy import Ui_Form
from codes.code_ui_kjy import Init_ui_kjy
import sip,codes.qjbl



class MySignal(QObject):
    move_sg=pyqtSignal(int,int,int,int)
class Init_window_main:
    def __init__(self, window: QMainWindow, ui: Ui_MainWindow):
        self.window = window
        self.window.setAttribute(Qt.WA_DeleteOnClose,True)
        self.ui = ui
        self.desktopWidth=qApp.desktop().width()
        self.screen = qApp.primaryScreen()
        self.x =self.desktopWidth-2#位置在窗口最右边隐藏起来
        self.img_2=0
        self.y = 0
        self.row=0
        self.hwnd=codes.qjbl.get_hwnd()#窗口句柄

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
        self.window.setWindowIcon(QIcon('logo.ico'))

        self.window.move(self.x, self.y)
        self.window.setWindowFlags(Qt.WindowStaysOnTopHint  | Qt.WindowCloseButtonHint)
        # 自定义信号
        self.move_sg = MySignal()
        # 绑定信号
        self.move_sg.move_sg.connect(self.move_window)
        self.ui.tabw_main.cellChanged.connect(self.cellChanged)
        self.ui.bt_save.clicked.connect(self.save)
        self.ui.bt_excel.clicked.connect(lambda: os.system(f"start {os.getcwd()}/{self.excelPath}"))
        self.ui.bt_kjyj.clicked.connect(self.open_kjyj)
        self.window.closeEvent=self.close
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
        self.windowHwind = int(self.window.winId())
        self.window_kjyj_Hwind=0
        t=threading.Thread(target=self.getWeChatNameNow)
        t.setDaemon(True)
        t.start()
    def close(self,enent):
        qApp.exit()
    def move_window(self,x,y,w,h):
        if self.window.x()!=x+w or self.window.y()!=y or self.window.height()!=h-30 :
            self.window.move(x+w,y)
            self.window.resize(self.window.width(),h-30)
            #这个地方,为什么不用move,,因为qt多线程非常的容易卡死,所以还不如用api来实现
            try:
                win32gui.SetWindowPos(self.hwnd_kjyj, win32con.HWND_TOPMOST, self.window.x()+self.window.width()-5, self.window.y(), self.init_ui_kjyj.window.width()+20, self.window.height()+38,
                                  win32con.SWP_SHOWWINDOW)
            except Exception as err:
                print(err,'move_winidow')
    def getWeChatNameNow(self):
        self.hwnd_kjyj=0
        sfjh=True
        change = False
        while True:
            try:
                xx,yy,ww,hh = 0,0,0,0
                # 查找微信的窗口句柄会优先获得最前的句柄
                hwnd=win32gui.GetForegroundWindow()
                #判断是否为支持的窗口
                if hwnd==0:
                    continue
                title=win32gui.GetWindowText(hwnd)
                class_=win32gui.GetClassName(hwnd)
                if title=='微信' and class_=='WeChatMainWndForPC':
                    self.hwnd=hwnd
                    codes.qjbl.set_hwnd(self.hwnd)
                    windowPost=[329, 23, 350, 30]
                elif title.find('阿里旺旺')!=-1 and class_=='StandardFrame':
                    self.hwnd=hwnd
                    codes.qjbl.set_hwnd(self.hwnd)
                    windowPost=[161,18,360,45]
                elif title!='QQ' and class_=='TXGuiFoundation' :
                    self.hwnd = hwnd
                    codes.qjbl.set_hwnd(self.hwnd)
                    windowPost=None
                    xx,yy,ww,hh = 0,6,-8,-14
                elif title=='钉钉' and class_=='StandardFrame_DingTalk':
                    self.hwnd = hwnd
                    codes.qjbl.set_hwnd(self.hwnd)
                    windowPost = [374,37,272,25]
                elif hwnd == self.windowHwind or hwnd == self.hwnd_kjyj:
                    win32gui.ShowWindow(self.windowHwind, win32con.SW_SHOW)
                    win32gui.ShowWindow(self.hwnd_kjyj, win32con.SW_SHOW)
                    continue
                else:
                    win32gui.ShowWindow(self.windowHwind, win32con.SW_HIDE)
                    win32gui.ShowWindow(self.hwnd_kjyj, win32con.SW_HIDE)
                    continue
                # 得到目标窗口的矩阵信息
                rect = win32gui.GetWindowRect(self.hwnd)
                x, y, w, h = rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]
                # 这个地方,为什么不用普通的qt函数,,因为qt多线程非常的容易卡死,所以还不如用api来实现
                #窗口置显示状态
                win32gui.ShowWindow(self.hwnd_kjyj, win32con.SW_SHOW)
                win32gui.ShowWindow(self.windowHwind, win32con.SW_SHOW)
                # 触发信号,把信息传递过去
                self.move_sg.move_sg.emit(x+xx,y+yy,w+ww,h+hh)
                if windowPost!=None:
                    # 截图
                    img = self.screen.grabWindow(self.hwnd, windowPost[0], windowPost[1], windowPost[2],windowPost[3]).toImage()
                    img.save(self.imagePath)
                    with open(self.imagePath, 'rb')as f:
                        img = f.read()
                    if img != self.img_2 or change==True:
                        try:
                            jstext = self.bd.get_text(image=img)  # 调用百度识字
                            print(jstext)
                            self.img_2 = copy.copy(img)
                            NameNow = jstext['words_result'][0]['words']
                            # 识图结果为空跳过
                            if NameNow == '':
                                return
                            NameNow = NameNow.replace('[淘宝网会员]', '').replace('②', '').replace('[宝网会员]', '').replace('遍宝网会员]量','')
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
                            change=False
                else:
                    change = True
                    NameNow=title.split('等')[0]
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
        try:
            self.init_ui_kjyj.window.close()
            self.init_ui_kjyj.__del__()
        except:
            pass
        window_kjyj=QFrame()
        ui_kjyj=Ui_Form()
        ui_kjyj.setupUi(window_kjyj)
        self.init_ui_kjyj = Init_ui_kjy(window_kjyj, ui_kjyj, self.window, self.ui)
        with open('qss\main.qss', 'r')as f:
            style = f.read()
        self.init_ui_kjyj.window.setStyleSheet(style)
        self.init_ui_kjyj.window.show()
        self.hwnd_kjyj=int(self.init_ui_kjyj.window.winId())
    def cellChanged(self,row,column):
        # 行列大小根据内容调整大小
        self.ui.tabw_main.resizeRowsToContents()
        self.ui.tabw_main.resizeColumnsToContents()
    def loadInfo(self,item,NameNow):
        '''
        根据表格某一行的数据,生成控件
        :param item:
        :return:
        '''
        self.ui.tabw_main.clearContents()
        if len(item)==0:
            labelItem = QTableWidgetItem(NameNow)
            #在最下面添加一行数据,并且把整个id记下
            self.kh.ws.append([NameNow,])
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
    def save(self):
        for i in range(self.ui.tabw_main.rowCount()):
            ch = chr(i + 65)
            try:
                self.kh.ws[f'{ch}{self.rowSelection}'] = self.ui.tabw_main.item(i, 1).text()

            except:
                self.kh.ws[f'{ch}{self.rowSelection}'] =''
        #保存
        self.kh.save()





