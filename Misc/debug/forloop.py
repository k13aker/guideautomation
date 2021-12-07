import subprocess
import win32con, win32api, win32gui, ctypes, ctypes.wintypes, threading, pyautogui as pag
from queue import Queue
from Misc.MyFunctions import readFromQueue

queue = Queue(1)


def thread(thread1, q):
    global data
    data = []

    class COPYDATASTRUCT(ctypes.Structure):
        _fields_ = [
            ('dwData', ctypes.wintypes.LPARAM),
            ('cbData', ctypes.wintypes.DWORD),
            ('lpData', ctypes.c_void_p)
        ]

    PCOPYDATASTRUCT = ctypes.POINTER(COPYDATASTRUCT)

    class Listener:
        def __init__(self):
            message_map = {
                win32con.WM_COPYDATA: self.OnCopyData
            }
            wc = win32gui.WNDCLASS()
            wc.lpfnWndProc = message_map
            wc.lpszClassName = 'MyPythonWindowClass'
            hinst = wc.hInstance = win32api.GetModuleHandle(None)
            classAtom = win32gui.RegisterClass(wc)
            self.hwnd = win32gui.CreateWindow(
                classAtom,
                "win32gui test",
                0,
                0,
                0,
                win32con.CW_USEDEFAULT,
                win32con.CW_USEDEFAULT,
                0,
                0,
                hinst,
                None
            )
            print(self.hwnd)

        def OnCopyData(self, hwnd, msg, wparam, lparam):
            pCDS = ctypes.cast(lparam, PCOPYDATASTRUCT)
            data = ctypes.wstring_at(pCDS.contents.lpData), wparam
            q.queue.clear()
            q.put(data, False)

    l = Listener()
    win32gui.PumpMessages()


api = threading.Thread(target=thread, args=("Thread-1", queue))
api.start()

subprocess.Popen("C:\\Program Files (x86)\\Dolphin\\GuideConnect\\Guide.EXE")

data = readFromQueue(queue)

i = 0
while i < 10:
    data = readFromQueue(queue)
    pag.sleep(2)
    print(data)
