import win32gui, win32con, win32api
import time, datetime
from apscheduler.schedulers.blocking import BlockingScheduler

#不使用守护线程  
schedudler = BlockingScheduler()
task_list = [("09:21", "do A"),
             ("02:36", "do B"),
             ("21:27", "do C")]


def getCurTime():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

#遍历所有窗口
def reset_window_pos():  
    hWndList = []  
    win32gui.EnumWindows(lambda hWnd, param: param.append(hWnd), hWndList)  
    for hwnd in hWndList:
        classname = win32gui.GetClassName(hwnd)
        title = win32gui.GetWindowText(hwnd)
        if "WeChatMainWndForPC" in classname and "微信" in title:
            print(classname, title, hwnd)

#遍历所有子窗口
def get_child_windows(parent):        
    '''     
    获得parent的所有子窗口句柄
     返回子窗口句柄列表
     '''     
    if not parent:         
        return      
    hwndChildList = []     
    win32gui.EnumChildWindows(parent, lambda hwnd, param: param.append(hwnd),  hwndChildList)          
    print(hwndChildList)

def sendMessage(message):
    handle = win32gui.FindWindow("WeChatMainWndForPC", "微信") 
    print(getCurTime(), "父窗口句柄：", hex(handle) )
    if not handle:
        print(getCurTime(), "找不到微信窗口,跳过。。。")
        return
    #置顶窗口
    print(getCurTime(), "置顶窗口")
    win32gui.ShowWindow(handle, win32con.SWP_SHOWWINDOW)
    win32gui.SetWindowPos(handle, win32con.HWND_TOPMOST, 200,100,1000,600, win32con.SWP_SHOWWINDOW)
    time.sleep(2)
    print(getCurTime(), "激活输入框")
    win32api.SetCursorPos((200 + 700, 100 + 550))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    print(getCurTime(), "输入信息。。。")
    for c in message:
        win32api.SendMessage(handle, win32con.WM_CHAR, ord(c), 0)
    #回车
    win32gui.PostMessage(handle,win32con.WM_KEYDOWN,win32con.VK_RETURN,0)
    win32gui.PostMessage(handle,win32con.WM_KEYUP,win32con.VK_RETURN,0)
    #发送完信息隐藏窗口
    time.sleep(1)
    win32gui.ShowWindow(handle, win32con.SW_HIDE)

#这里因为不知道哪里丢失开始8个字符，这里先随机填空8个字符，再正式输出
# sendMessage("微信通知")


def job():
    print(getCurTime())
    sendMessage("微信通知")
    time.sleep(2)
# schedudler.add_job(job, 'cron', day_of_week="0-4", hour=8, minute=55)
schedudler.add_job(job, 'cron', day_of_week="0-4", hour=15, minute=26)
schedudler.start()