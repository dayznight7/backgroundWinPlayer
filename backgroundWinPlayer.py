from ctypes import windll, wintypes, byref, sizeof
from PIL import Image
import win32gui, win32ui, win32api, win32con
import pyautogui as pag
import time, math, random
import numpy as np


def get_window_list():
    def callback(hwnd, hwnd_list: list):
        title = win32gui.GetWindowText(hwnd)
        if win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd) and title:
            hwnd_list.append((title, hwnd))
        return True
    output = []
    win32gui.EnumWindows(callback, output)
    return output


def print_window_list():
    print(' '*6 + "HWND WINDOW_TITLE")
    print("\n".join("{: 10d} {}".format(__, _) for _, __ in get_window_list()))


def print_mouse_location():
    while True:
        x, y = pag.position()
        pos = '{:04d} {:04d}'.format(x, y)
        print("\r"+pos, end='')
        time.sleep(1)


# use it when FindWindow doesn't work
def get_window_rect(hwnd):
    f = windll.dwmapi.DwmGetWindowAttribute
    rect = wintypes.RECT()
    DWMWA_EXTENDED_FRAME_BOUNDS = 9
    f(wintypes.HWND(hwnd),
      wintypes.DWORD(DWMWA_EXTENDED_FRAME_BOUNDS),
      byref(rect),
      sizeof(rect)
      )
    return rect.left, rect.top, rect.right, rect.bottom


# HAYSTACK
def get_inactive_img(window_title):
    if str(type(window_title)) == "<class 'str'>":
        hwnd = win32gui.FindWindow(None, window_title)
        left, top, right, bot = win32gui.GetWindowRect(hwnd)
    else:
        hwnd = window_title
        left, top, right, bot = get_window_rect(hwnd)
    width = right - left
    height = bot - top
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()
    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    saveDC.SelectObject(saveBitMap)
    result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)
    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)
    img = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1
    )
    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)
    if result == 1:
        return img
    else:
        return 0


# return x1,y1,x2,y2
def detect(address, window_title):
    img_needle = Image.open(address)
    img_haystack = get_inactive_img(window_title)
    Matr = pag.locate(img_needle, img_haystack, confidence=0.9)
    if Matr != None:
        return Matr[0], Matr[1], Matr[0]+Matr[2], Matr[1]+Matr[3]
    else:
        time.sleep(1)
        return detect(address, window_title)


def wait_for(address, window_title, timescale=10, cfd=0.9):
    img_needle = Image.open(address)
    img_haystack = get_inactive_img(window_title)
    Matr = pag.locate(img_needle, img_haystack, confidence=cfd)
    if Matr != None:
        return
    else:
        time.sleep(timescale)
        return wait_for(address, window_title, timescale, cfd)


def coordinate_rect(x1, y1, x2, y2):
    x0 = int((x1+x2)/2)
    y0 = int((y1+y2)/2)
    x = np.random.binomial(n=x0*4, p=0.25, size=1)
    y = np.random.binomial(n=y0*4, p=0.25, size=1)
    if x1 <= x[0] <= x2 and y1 <= y[0] <= y2:
        return x[0], y[0]
    else:
        return coordinate_rect(x1, y1, x2, y2)


def coordinate_circle(x1, y1, x2, y2):
    x0 = int((x1 + x2) / 2)
    y0 = int((y1 + y2) / 2)
    R = min(x2-x0, y2-y0)
    r = np.random.normal(0, 1, 1)
    if 0 <= abs(r[0]) <= 2:
        rad = random.uniform(0, 2 * math.pi)
        x = int(x0 + 0.5 * R * r[0] * math.cos(rad))
        y = int(y0 + 0.5 * R * r[0] * math.sin(rad))
        return x, y
    else:
        return coordinate_circle(x1, y1, x2, y2)


# Keystroke level model - Human Computer Interaction - mouse click(100ms)
def send_updown_mouse(x, y, window_title):
    hwnd = win32gui.FindWindow(None, window_title)
    lParam = win32api.MAKELONG(x, y)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
    time.sleep(random.uniform(0.09, 0.12))
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lParam)
    time.sleep(random.uniform(0.03, 0.04))


def send_char(window_title, char):
    hwnd = win32gui.FindWindow(None, window_title)
    win32api.SendMessage(
        hwnd, win32con.WM_KEYDOWN, ord(char),
        0 + (0 << 8) + (ord(char) << 16) + (0 << 24))
    win32api.SendMessage(
        hwnd, win32con.WM_CHAR, ord(char),
        0 + (0 << 8) + (ord(char) << 16) + (0 << 24))
    win32api.SendMessage(
        hwnd, win32con.WM_KEYUP, ord(char),
        0 + (0 << 8) + (ord(char) << 16) + (0xC0 << 24))


# https://docs.microsoft.com/ko-kr/windows/win32/inputdev/virtual-key-codes#requirements
def get_vk(str):
    if str == "backspace":
        return win32con.VK_BACK
    elif str == "tab":
        return win32con.VK_TAB
    elif str == "shift":
        return win32con.VK_SHIFT
    elif str == "enter":
        return win32con.VK_RETURN
    elif str == "ctrl":
        return win32con.VK_CONTROL
    elif str == "alt":
        return win32con.VK_MENU
    elif str == "caps lock":
        return win32con.VK_CAPITAL
    elif str == "esc":
        return win32con.VK_ESCAPE
    elif str == "space":
        return win32con.VK_SPACE
    elif str == "up":
        return win32con.VK_UP
    elif str == "down":
        return win32con.VK_DOWN
    elif str == "left":
        return win32con.VK_LEFT
    elif str == "right":
        return win32con.VK_RIGHT
    elif str == "insert":
        return win32con.VK_INSERT
    elif str == "delete":
        return win32con.VK_DELETE
    elif str == "f5":
        return win32con.VK_F5
    else:
        return


def send_ctrlkey(window_title, ctrlkey):
    tmp = get_vk(ctrlkey)

    hwnd = win32gui.FindWindow(None, window_title)
    win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, tmp)
    win32api.SendMessage(hwnd, win32con.WM_KEYUP, tmp)


def send_string(window_title, str):
    for _ in str:
        send_char(window_title, _)

