import json
import win32gui
from pywinauto import Desktop
import pyautogui
import time

def load_json(path):
    if not path.exists():
        return {}

    with path.open("r", encoding="utf-8") as f:
        return json.load(f)
    
def save_json(path, data):
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

def get_window_name():
    hwnd = win32gui.GetForegroundWindow()
    desktop = Desktop(backend="uia")

    
    active_window = desktop.window(handle=hwnd)
    return active_window.window_text()

def rec_find_text(ctrl, text):
    txt = ctrl.window_text().strip()
    if txt:
        if txt == text:
            return True
    for child in ctrl.children():
        ret = rec_find_text(child, text)
        if ret == True:
            return ret
    return False
    

def find_text(text):
    hwnd = win32gui.GetForegroundWindow()
    desktop = Desktop(backend="uia")

    
    ctrl = desktop.window(handle=hwnd)
    return rec_find_text(ctrl, text)

def rec_find_subtext(ctrl, text):
    txt = ctrl.window_text().strip()
    if txt:
        if text in txt:
            return txt
    for child in ctrl.children():
        ret = rec_find_subtext(child, text)
        if ret != 0:
            return ret
    return 0
    

def find_subtext(text):
    hwnd = win32gui.GetForegroundWindow()
    desktop = Desktop(backend="uia")

    
    ctrl = desktop.window(handle=hwnd)
    return rec_find_subtext(ctrl, text)

def find_ref_dates(ctrl, count, dates = list()):
    txt = ctrl.window_text().strip()
    if txt and len(txt) >= 10 and ctrl.element_info.control_type == "Button":
        if txt[2] == "/" and txt[5] == "/":
            count += 1
            if count == 1:
                dates.append(txt)
            elif count == 3:
                dates.append(txt)
                return dates, count
    for child in ctrl.children():
        dates, count = find_ref_dates(child, count, dates)
        if len(dates) == 2:
            return dates, count
    return dates, count
    

def three_month():
    hwnd = win32gui.GetForegroundWindow()
    desktop = Desktop(backend="uia")

    
    ctrl = desktop.window(handle=hwnd)
    dates,__ = find_ref_dates(ctrl, 0)
    return dates[0].split("/")[1] != dates[1].split("/")[1]

def move_window():
    hwnd = win32gui.GetForegroundWindow()
    desktop = Desktop(backend="win32")

    
    win = desktop.window(handle=hwnd)
    win.move_window(x=100, y=100)

def find_status(ctrl, found):
    txt = ctrl.window_text().strip()
    if txt:
        if txt == "Arrived":
            found = 4
        elif txt == "Confirmed":
            found = 5
        elif txt == "Publish Ext":
            return found, True

    for child in ctrl.children():
        found, status = find_status(child, found)
        if status:
            return found, True
    return found, False

def booking_status():
    hwnd = win32gui.GetForegroundWindow()
    desktop = Desktop(backend="uia")

    
    active_window = desktop.window(handle=hwnd)
    return find_status(active_window, 0)[0]

def find_keywords(ctrl, found):
    txt = ctrl.window_text().strip()
    if txt:
        if txt == "Patient":
            return 5, found
        if txt == "Medicare":
            if found == 2:
                return 4, found
            found += 1
        if txt == "DVA":
            return 3, found
        if txt == "Workcover":
            return 2, found
        if txt == "3rd Party":
            return 1, found
        if txt == "Health Fund":
            if found == 2:
                return 0, found
            found += 1
    for child in ctrl.children():
        ret, found = find_keywords(child, found)
        if ret != -1:
            return ret, found
    return -1, found
    

def determineInvoice():
    hwnd = win32gui.GetForegroundWindow()
    desktop = Desktop(backend="uia")

    
    ctrl = desktop.window(handle=hwnd)
    return find_keywords(ctrl, 0)[0]


def find_dates(ctrl):
    txt = ctrl.window_text().strip()
    if txt:
        if len(txt) >= 13 and txt[2] == "/" and txt[5] == "/" and txt[11:13] == "to":
            return txt

    for child in ctrl.children():
        ret = find_dates(child)
        if ret != 0:
            return ret
    
    return 0

def referral_date_range():
    hwnd = win32gui.GetForegroundWindow()
    desktop = Desktop(backend="uia")

    
    active_window = desktop.window(handle=hwnd)
    return find_dates(active_window)

def switch_to_window(text):
    time.sleep(0.5)
    count = 0
    while get_window_name() != text:
        count += 1
        pyautogui.keyDown("alt")
        pyautogui.press("tab", count) 
        pyautogui.keyUp("alt")
        time.sleep(0.5) 

def orderTimes(time_1, time_2):
    if len(time_1) < 4:
        time_1 = [time_1[0], time_1[1:]]
    else:
        time_1 = [time_1[0:2], time_1[2:]]

    if len(time_2) < 4:
        time_2 = [time_2[0], time_2[1:]]
    else:
        time_2 = [time_2[0:2], time_2[2:]]

    if int(time_1[0]) > int(time_2[0]):
        earlier = time_2
        later = time_1
    elif int(time_1[0]) < int(time_2[0]):
        earlier = time_1
        later = time_2
    elif int(time_1[1]) > int(time_2[1]):
        earlier = time_2
        later = time_1
    else:
        earlier = time_1
        later = time_2
    
    earlierPM = False
    laterPM = False
    if int(earlier[0]) >= 12:
        if int(earlier[0]) != 12:
            earlier[0] = str(int(earlier[0]) - 12)
        earlierPM = True
    if int(later[0]) >= 12:
        if int(later[0]) != 12:
            later[0] = str(int(later[0]) - 12)
        laterPM = True
    
    return earlier, later, earlierPM, laterPM
