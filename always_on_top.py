import ctypes
from ctypes import wintypes
import keyboard

user32 = ctypes.WinDLL('user32', use_last_error=True)

HWND_TOPMOST = -1
HWND_NOTOPMOST = -2
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001

SetWindowPos = user32.SetWindowPos
SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint]
SetWindowPos.restype = wintypes.BOOL

GetForegroundWindow = user32.GetForegroundWindow
GetForegroundWindow.argtypes = []
GetForegroundWindow.restype = wintypes.HWND

def set_always_on_top(hwnd, top=True):
    if top:
        SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)
    else:
        SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)

def on_shift_f2():
    hwnd = GetForegroundWindow()
    set_always_on_top(hwnd, True)
    print("현재 창을 항상 맨 위로 고정했습니다.")

def on_shift_f1():
    hwnd = GetForegroundWindow()
    set_always_on_top(hwnd, False)
    print("현재 창의 항상 맨 위 고정을 해제했습니다.")

keyboard.add_hotkey('shift+f2', on_shift_f2)
keyboard.add_hotkey('shift+f1', on_shift_f1)

print("SHIFT+F2로 창 고정, SHIFT+F1로 해제 가능합니다.")
keyboard.wait()
