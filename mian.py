import win32gui, win32con,win32com.client
import re, traceback
from time import sleep
import os
import pywinauto
import pyautogui
from pywinauto.timings import wait_until
import time
from pywinauto import findwindows

def is_handle_closed(handle, max_attempts=10, wait_time=1):
    attempts = 0
    while attempts < max_attempts:
        try:
            findwindows.find_element(handle=handle)
        except Exception as e:
            return True
        time.sleep(wait_time)
        attempts += 1
    return False
class cWindow:
    def __init__(self):
        self._hwnd = None
    def SetAsForegroundWindow(self):
        # First, make sure all (other) always-on-top windows are hidden.
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        self.hide_always_on_top_windows()
        win32gui.SetForegroundWindow(self._hwnd)
    def Maximize(self):
        win32gui.ShowWindow(self._hwnd, win32con.SW_MAXIMIZE)
    def _window_enum_callback(self, hwnd, regex):
        '''Pass to win32gui.EnumWindows() to check all open windows'''
        if self._hwnd is None and re.match(regex, str(win32gui.GetWindowText(hwnd))) is not None:
            self._hwnd = hwnd

    def find_window_regex(self, regex):
        self._hwnd = None
        win32gui.EnumWindows(self._window_enum_callback, regex)
    def hide_always_on_top_windows(self):
        win32gui.EnumWindows(self._window_enum_callback_hide, None)
    def _window_enum_callback_hide(self, hwnd, unused):
        if hwnd != self._hwnd: # ignore self
            # Is the window visible and marked as an always-on-top (topmost) window?
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) & win32con.WS_EX_TOPMOST:
                # Ignore windows of class 'Button' (the Start button overlay) and
                # 'Shell_TrayWnd' (the Task Bar).
                className = win32gui.GetClassName(hwnd)
                if not (className == 'Button' or className == 'Shell_TrayWnd'):
                    # Force-minimize the window.
                    # Fortunately, this seems to work even with windows that
                    # have no Minimize button.
                    # Note that if we tried to hide the window with SW_HIDE,
                    # it would disappear from the Task Bar as well.
                    win32gui.ShowWindow(hwnd, win32con.SW_FORCEMINIMIZE)
    def get_hwnd (self):
        return (self._hwnd)
def active_gui(title):
 
    try:
        regex = ".*{}.*".format(title)
        cW = cWindow()
        cW.find_window_regex(regex)
        cW.Maximize()
        cW.SetAsForegroundWindow()
        
    except:
        #f = open(os.path.join(os.getcwd(),"log.txt"), "w")
        #f.write(traceback.format_exc())
        print(traceback.format_exc())

def get_handle(title):
    try:
        regex = ".*{}.*".format(title)
        cW = cWindow()
        cW.find_window_regex(regex)
        return cW.get_hwnd()
    except:
        pass
def get_app():
    
    os.environ.update({"__COMPAT_LAYER":"RUnAsInvoker"})

    app=pywinauto.Application(backend='uia').start(r'D:\QQ\Bin\QQScLauncher.exe', timeout=20)
    time.sleep(2)
    app=pywinauto.Application(backend='uia')
    handle=get_handle('QQ')
    app.connect(handle=handle)
    main_dlg = app.window(handle=handle)
    wait_until(10, 1, lambda: main_dlg.child_window(title="密码", control_type="Pane").exists())
    password_pane = main_dlg.child_window(title="密码", control_type="Pane")
    password_pane_rect = password_pane.rectangle()
    password_center_x = (password_pane_rect.left + password_pane_rect.right) // 2
    password_center_y = (password_pane_rect.top + password_pane_rect.bottom) // 2
    pyautogui.click(password_center_x, password_center_y)


    pyautogui.click(password_center_x, password_center_y)
    [pyautogui.press('backspace') for x in range(20)]
    pyautogui.typewrite('*********', interval=0.1)#密码
    app.window(handle=handle)['登录Button'].click_input()
    for i in range(2):
        
        if is_handle_closed(handle):
            print("窗口已关闭")
            break
        else:
            print("窗口仍然打开")
            time.sleep(0.5)
            

        
    for i in range(20):
        
        handle=get_handle('QQ')
        if handle:
            
            break
        else:
            time.sleep(1)
            
    app.connect(handle=handle)

    #打开聊天窗口 以自己QQ号为初始窗口
    app.window(handle=handle).child_window(title="搜索：联系人、群聊、企业", control_type="Edit").click_input()
    app.window(handle=handle).child_window(title="搜索：联系人、群聊、企业", control_type="Edit").type_keys('33333333')
    target_rect=app.window(handle=handle)['机械飞升 33333333'].rectangle()
    target_center = target_rect.mid_point()
    pyautogui.doubleClick(target_center.x, target_center.y)

    time.sleep(1)
    app.window(handle=handle).child_window(title="搜索：联系人、群聊、企业", control_type="Edit").click_input()
    app.window(handle=handle).child_window(title="搜索：联系人、群聊、企业", control_type="Edit").type_keys('55555555')
    target_rect=app.window(handle=handle)['Python技术交流群 55555555'].rectangle()
    target_center = target_rect.mid_point()
    pyautogui.doubleClick(target_center.x, target_center.y)
    #聊条窗口句柄
    for i in range(10):
        
        handle_show=get_handle('机械飞升')
        if handle_show:
            break
        else:
            time.sleep(0.5)

    #链接聊条窗口
    app_show=app.connect(handle=handle_show)
    return app_show,handle_show
def connect(title):#进入指定窗口
    
    app_show=app.connect(handle=handle_show)
    app_show.window(handle=handle_show).child_window(title="搜索：联系人、群聊、企业", control_type="Edit").click_input()
    app_show.window(handle=handle_show).child_window(title="搜索：联系人、群聊、企业", control_type="Edit").type_keys('33333333')

    wait_until(10, 1, lambda: app_show.window(handle=handle_show)['机械飞升 33333333'].exists())
    target_rect=app_show.window(handle=handle_show)['机械飞升 33333333'].rectangle()

    target_center = target_rect.mid_point()
    pyautogui.doubleClick(target_center.x, target_center.y)
    time.sleep(0.5)


def send_msg(text):#发送消息
    app_show=app.connect(handle=handle_show)
    app_show.window(handle=handle_show).child_window(title="输入", control_type="Edit").click_input()
    app_show.window(handle=handle_show).child_window(title="输入", control_type="Edit").type_keys(text)
    app_show.window(handle=handle_show).child_window(title="发送(&S)", control_type="Button").click_input()
