import win32con, win32api, win32gui, ctypes, ctypes.wintypes, threading
from Suites.Sanity import Sanity
from queue import Queue
from Misc.MyFunctions import *

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
launch()
data = readFromQueue(queue)


while 'Emails' not in data:
    print("Guide Launching, Please wait..")
    pag.sleep(15)
    data = readFromQueue(queue)
    if 'Emails' in data:  # if no tutorial videos included (MSI)
        pag.sleep(15)
        break

    elif "GMENU_GETTING_STARTED" in data:  # if tutorial videos are included (Installers)
        gstart(queue)
        break

# Email
Sanity.emails(queue)  # Signs in, Sends Email, Receives Email

# Letters and Documents
Sanity.docs(queue)  # Tests creating/opening documents
Sanity.letter(queue)  # Tests filling sender address and creating letter

# Websites

# Scanner and Camera
Sanity.scanner(queue)  # Tests scanner
Sanity.camera(queue)   # Tests camera

# Books and Newspapers
Sanity.books(queue)     # Downloads book from epubBook and reads it
Sanity.newspapers(queue)  # Downloads newspaper and views it

# Address book and Calendar
Sanity.address_book(queue)  # Tests creating a new contact and editing a contact
Sanity.calendar(queue)  # Tests making a calendar event and viewing

# Entertainment
Sanity.radio(queue)   # Plays 3 radio stations from the list
Sanity.podcasts(queue)  # Plays and downloads podcasts
Sanity.music(queue)   # imports music and plays a track
Sanity.games(queue)   # Launches all 3 games

# Notes
Sanity.text_note(queue)    # Creates and views text note
Sanity.audio_note(queue)   # Creates and views audio note

# Tools
Sanity.calc(queue)    # Tests Calculator with: +, /, *, -
Sanity.dictionary(queue)  # Tests Dictionary: Tests 1 valid and 1 invalid entry
Sanity.gstart_vids(queue)  # Tests Getting started: the 4 integrated videos
Sanity.shortcut(queue)   # Tests the shortcut key documentation
Sanity.learn_keyboard(queue)  # Tests learn the keyboard: all keyboard inputs
Sanity.file_exp(queue)    # Moves a file from documents to downloads and checks
Sanity.about(queue)   # Tests About page: Checks premium plan status

# Settings
Sanity.settings(queue)  # Changes to theme 5 and changes font

# Exit
Sanity.exit_module(queue)
