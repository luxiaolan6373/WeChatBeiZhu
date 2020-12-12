#设置全局变量
hwnd=0

def set_hwnd(i):
    global hwnd
    hwnd=i
def get_hwnd():
    global hwnd
    return hwnd