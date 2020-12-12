import win32api, win32con,time,win32gui
import win32clipboard as w
class Mouse_And_Key():
    '''
    鼠标键盘前台模拟api调用类
    '''
    def getCursorPos(self):
            '''
            获取当前鼠标相对屏幕的位置
            :return:x,y
            '''
            pos = win32api.GetCursorPos()
            return int(pos[0]), int(pos[1])

    def mouse_move(self, x, y):
        '''
        设置鼠标位置
        :return:
        '''
        win32api.SetCursorPos((x, y))

    def mouseLeftDown(self):
        '''
        鼠标左键按下
        :return:
        '''
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

    def mouseLeftUP(self):
        '''
        鼠标左键弹起
        :return:
        '''
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

    def mouseRightDown(self):
        '''
        鼠标右键按下
        :return:
        '''
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)

    def mouseRightUp(self):
        '''
        鼠标右键弹起
        :return:
        '''
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)

    def mouse_left_click(self, new_x=0, new_y=0, times=1, ys=0.05, ys2=0.02):
        '''
        鼠标移至点击
        :param new_x:x坐标
        :param new_y:y坐标
        :param times:点击次数
        :param ys:移动到点击之间的延迟单位秒 默认为0.05s
        :param ys2:点击按下弹起之间的延迟单位秒 默认为0.02s
        :return:
        '''
        self.mouse_move(new_x, new_y)
        time.sleep(0.05)
        while times:
            self.mouseLeftDown()
            time.sleep(0.02)
            self.mouseLeftUP()
            times -= 1

    def mouse_right_click(self, new_x=0, new_y=0, ys=0.05, ys2=0.02):
        '''
        鼠标移至点击
        :param new_x:x坐标
        :param new_y:y坐标
        :param ys:移动到点击之间的延迟单位秒 默认为0.05s
        :param ys2:点击按下弹起之间的延迟单位秒 默认为0.02s
        :return:
        '''
        self.mouse_move(new_x, new_y)
        time.sleep(0.05)
        self.mouseReftDown()
        time.sleep(0.02)
        self.mouseReftUP()
    def key_input(self, input_words=''):
        '''
        键盘输入文本
        :param input_words: 输入的字符串文本
        :return:
        '''
        for word in input_words:
            win32api.keybd_event(word, 0, 0, 0)
            win32api.keybd_event(word, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.1)

    def key_even(self, input_key):
        '''
        模拟按键
        :param input_key:
        :return:
        '''
        win32api.keybd_event(input_key, 0, 0, 0)  # enter
        win32api.keybd_event(input_key, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
    def key_even_zuhe(self, input_key,function_key,ys=0.05):
        '''
        模拟组合按键,都是填键码
        :param input_key:主键
        :param function_key:功能键
        :param ys:点击按下弹起之间的延迟单位秒 默认为0.02s
        :return:
        '''
        win32api.keybd_event(function_key, 0, 0, 0)
        win32api.keybd_event(input_key, 0, 0, 0)
        time.sleep(0.02)
        win32api.keybd_event(input_key, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(function_key, 0, win32con.KEYEVENTF_KEYUP, 0)
    def get_text(self):
        '''
        读取剪辑版文本
        :return:
        '''
        w.OpenClipboard()
        d = w.GetClipboardData(win32con.CF_TEXT)
        w.CloseClipboard()
        return d.decode('GBK')

    def set_text(self,aString):
        '''
        设置剪辑版文本
        :param aString:
        :return:
        '''
        w.OpenClipboard()
        w.EmptyClipboard()
        w.SetClipboardData(win32con.CF_UNICODETEXT, aString)
        w.CloseClipboard()
    def set_window_foregGroun(self,hwnd):
        '''
        激活指定窗口
        :param hwnd:
        :param flag:
        :return:
        '''
        win32gui.SetForegroundWindow(hwnd)


