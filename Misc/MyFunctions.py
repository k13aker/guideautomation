import pyautogui as pag, win32api
import os, shutil, subprocess, datetime
from colorama import init, Fore, Back, Style

data = []
user = os.environ['USERPROFILE']
found = []

init()
FORES = [ Fore.BLACK, Fore.RED, Fore.GREEN, Fore.YELLOW, Fore.BLUE, Fore.MAGENTA, Fore.CYAN, Fore.WHITE ]
BACKS = [ Back.BLACK, Back.RED, Back.GREEN, Back.YELLOW, Back.BLUE, Back.MAGENTA, Back.CYAN, Back.WHITE ]
BRIGHTNESS = [ Style.DIM, Style.NORMAL, Style.BRIGHT ]


def readFromQueue(q):
    global data
    if q.empty():
        return data
    else:
        data = q.get()
        return data


def returnhome():
    pag.press("esc", presses=10)


def delay():
    pag.sleep(5)


def launch():
    appdata = os.environ['USERPROFILE'] + "\\AppData\\Roaming\\Dolphin\\GuideConnect"

    # Removes Appdata folder
    try:
        shutil.rmtree(appdata)
    except OSError as e:
        print("Error: %s : %s" % (appdata, e.strerror))

    # Creates 'DIR' Guide Connect
    dir1 = os.environ['USERPROFILE'] + "\\AppData\\Roaming\\Dolphin\\GuideConnect"
    os.mkdir(dir1)

    # Creates DIR 'Scripts'
    dir2 = os.environ['USERPROFILE'] + "\\AppData\\Roaming\\Dolphin\\GuideConnect\\Scripts"
    os.mkdir(dir2)

    # Moves Testing.py to Appdata
    x = os.environ['USERPROFILE'] + "\\Desktop\\Guide Automation\\Misc\\Testing.py"
    y = os.environ['USERPROFILE'] + "\\AppData\\Roaming\\Dolphin\\GuideConnect\\Scripts"
    shutil.copy(x, y)

    # Removes Guide Connect Documents folder
    docs = os.environ['USERPROFILE'] + "\\Documents\\GuideConnect"
    try:
        shutil.rmtree(docs)
    except OSError as e:
        print("Error: %s : %s" % (docs, e.strerror))

    # Launches Guide Connect
    subprocess.Popen("C:\\Program Files (x86)\\Dolphin\\GuideConnect\\Guide.EXE")


def contacts():
    # Overrides Contacts.db with Test data
    x = os.environ['USERPROFILE'] + "\\Desktop\\Guide Automation\\Misc\\contacts\\GuideContacts.db"
    y = os.environ['USERPROFILE'] + "\\AppData\\Roaming\\Dolphin\\GuideConnect\\Contacts\\GuideContacts.db"
    shutil.copy(x, y)


def gstart(queue):
    pag.sleep(15)
    pag.press("enter", interval=1)
    data = readFromQueue(queue)
    if 'Keyboard and mouse video' in data:
        pag.sleep(3)
        pag.press("s", interval=1)
        pag.sleep(3)
        pag.press("enter", interval=1)
        pag.press("right", interval=1)
        pag.press("enter", interval=1)
        delay()

    else:
        pass


def testdata():
    pag.sleep(15)
    x = os.environ['USERPROFILE'] + "\\Desktop\\Guide Automation\\Misc\\testdata"
    y = os.environ['USERPROFILE'] + "\\Documents\\GuideConnect\\Documents\\testdata"
    try:
        shutil.copytree(x, y)
    except OSError as testdata_error:
        print("Error:", testdata_error.strerror)

    a = os.environ['USERPROFILE'] + "\\Desktop\\Guide Automation\\Misc\\Music"
    b = os.environ['USERPROFILE'] + "\\Documents\\GuideConnect\\Music\\Music"
    try:
        shutil.copytree(a, b)
    except OSError as testmusic_error:
        print("Error:", testmusic_error.strerror)


def drive_locate():
    try:
        c = win32api.GetVolumeInformation("C://")
        if c[0] == 'Regression':
            found.append("C:\\")
    except:
        pass

    try:
        d = win32api.GetVolumeInformation("D://")
        if d[0] == 'Regression':
            found.append("D:\\")

    except:
        pass

    try:
        e = win32api.GetVolumeInformation("E://")
        if e[0] == 'Regression':
            found.append("E:\\")
    except:
        pass

    try:
        f = win32api.GetVolumeInformation("F://")
        if f[0] == 'Regression':
            found.append("F:\\")
    except:
        pass

    try:
        g = win32api.GetVolumeInformation("G://")
        if g[0] == 'Regression':
            found.append("G:\\")
    except:
        pass

    try:
        h = win32api.GetVolumeInformation("H://")
        if h[0] == 'Regression':
            found.append("H:\\")
    except:
        pass

    try:
        i = win32api.GetVolumeInformation("I://")
        if i[0] == 'Regression':
            found.append("I:\\")
    except:
        pass

    try:
        j = win32api.GetVolumeInformation("J://")
        if j[0] == 'Regression':
            found.append("J:\\")
    except:
        pass

    try:
        k = win32api.GetVolumeInformation("K://")
        if k[0] == 'Regression':
            found.append("K:\\")
    except:
        pass


def init():
    pag.sleep(20)


def print_with_color(s, color=Fore.WHITE, brightness=Style.NORMAL, **kwargs):
    print(f"{brightness}{color}{s}{Style.RESET_ALL}", **kwargs)

