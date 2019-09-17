import win32clipboard as wc
import time
import win32con
import win32api

# main()

# 设置剪切板内容
def setText(s):
    wc.OpenClipboard()
    wc.EmptyClipboard()
    wc.SetClipboardData(win32con.CF_UNICODETEXT,s)
    wc.CloseClipboard()

# 读取剪切板
def getText():
    wc.OpenClipboard()
    ret = wc.GetClipboardData(win32con.CF_TEXT)
    wc.CloseClipboard()
    return ret

VK_CODE ={
    'enter':0x0D,
    'ctrl':0x11,
    'a':0x41,
    'c':0x43,
    'v':0x56,
    'x':0x58
    }


def keyDown(keyName):
    win32api.keybd_event(VK_CODE[keyName],0, 0, 0)


def keyUp(keyName):
    win32api.keybd_event(VK_CODE[keyName], 0, win32con.KEYEVENTF_KEYUP, 0)

# 复制
def copyByText(content):
    # 将content内容设置到剪切板中
    setText(content)
    # 从剪切板中获取刚设置到剪切板中的内容
    time.sleep(0.5)
    getContent = getText()
    return getContent

def copyByKey():
    keyDown('ctrl')
    keyDown('c')
    keyUp('c')
    keyUp('ctrl')
    time.sleep(0.5)

# 粘贴
def paste():
    # 按下ctrl+v组合键
    keyDown('ctrl')
    keyDown('v')
    # 释放ctrl+v组合键
    keyUp('v')
    keyUp('ctrl')
    time.sleep(0.5)

# 剪切
def cut():
    keyDown('ctrl')
    keyDown('x')
    keyUp('x')
    keyUp('ctrl')

def selectAll():
    keyDown('ctrl')
    keyDown('a')
    keyUp('a')
    keyUp('ctrl')
    time.sleep(0.5)
