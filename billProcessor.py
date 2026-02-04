from pywinauto import Desktop
import pyautogui
import time
import keyboard
import os
import win32gui
import string
import win32gui
from pathlib import Path
import json
from configureButtons import configure, ButtonPress
import re
from datetime import datetime
import winsound
import builtins

skip = False

# Save originals
_original_print = builtins.print
_original_input = builtins.input

def beep_print(*args, **kwargs):
    winsound.Beep(800, 60)
    _original_print(*args, **kwargs)

def beep_input(prompt=""):
    winsound.Beep(800, 60)
    return _original_input(prompt)

builtins.print = beep_print
builtins.input = beep_input

def processPatient(patient, buttons):
    global skip
    skip = False
    
    def load_json(path):
        if not path.exists():
            return {}

        with path.open("r", encoding="utf-8") as f:
            return json.load(f)
    
    def save_json(path, data):
        with path.open("w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

    def skip_patient():
        global skip
        skip = True
        print("Skipping Patient - Press SHIFT+TAB to continue with action")

    def reverse_skip():
        global skip
        skip = False
        print("Reversing skip, patient will continue to be processed - Press SHIFT+TAB")

    keyboard.add_hotkey("ctrl+alt+s", skip_patient)
    keyboard.add_hotkey("ctrl+alt+a", reverse_skip)

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

        # Get the foreground window by handle
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

        # Get the foreground window by handle
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

        # Get the foreground window by handle
        ctrl = desktop.window(handle=hwnd)
        dates,__ = find_ref_dates(ctrl, 0)
        return dates[0].split("/")[1] != dates[1].split("/")[1]

    def get_window_name():
        hwnd = win32gui.GetForegroundWindow()
        desktop = Desktop(backend="uia")

        # Get the foreground window by handle
        active_window = desktop.window(handle=hwnd)
        return active_window.window_text()

    def move_window():
        hwnd = win32gui.GetForegroundWindow()
        desktop = Desktop(backend="win32")

        # Get the foreground window by handle
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

        # Get the foreground window by handle
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

        # Get the foreground window by handle
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

        # Get the foreground window by handle
        active_window = desktop.window(handle=hwnd)
        return find_dates(active_window)

    

    #for patient in patients
    buttons["PatientWindow"].clickButton()
    buttons["FindPatient"].clickButton()

    fullName = patient["PatientName"]
    fullName = fullName.replace(",", "").replace("/", "").replace("\n", "")

    names = fullName.lower().split(" ")
    i = 0
    count = 0
    TITLES = {
        "mr", "mrs", "ms", "miss",
        "dr", "prof", "sir", "madam",
        "mx"
    }
    while count < 2:
        if names[i] in TITLES:
            i += 1
        else:
            pyautogui.typewrite(names[i])
            pyautogui.press("tab")
            i += 1
            count += 1

    pyautogui.press("left")

    dob = patient["PatientDOB"]
    pyautogui.typewrite(dob)
    pyautogui.press("enter")
    time.sleep(3)
    if not find_text("Row 2") and find_text("Row 1"):
        while get_window_name() != "Find Patient":
            pyautogui.hotkey(["alt", "tab"])
            time.sleep(0.2)  
        pyautogui.press("enter", 2)
    else:
        print("please put in the correct patient and remove window and then press SHIFT+TAB")

        keyboard.wait("shift+tab", suppress=True)
        if skip:
            pyautogui.hotkey("alt", "f4")
            return 1
    
    time.sleep(0.2)
    while get_window_name() != "HealthTrack":
        time.sleep(0.2)

    buttons["Demo"].clickButton()

    buttons["OPV"].clickButton()

    time.sleep(0.2)
    while find_text("Processing..."):
        time.sleep(0.2)
    
    if find_text("Patient is eligible to claim for Medicare with details provided."):
        pyautogui.press("enter")
    elif find_subtext("not an exact match"):
        pyautogui.press("tab", 2)
        pyautogui.press("enter")
        buttons["Merge1"].clickButton()
        buttons["Merge2"].clickButton()
        buttons["Merge3"].clickButton()
        buttons["MergeSelected"].clickButton()
        pyautogui.press("tab", 2)
        pyautogui.press("enter")
        time.sleep(0.2)
        while find_text("Processing..."):
            time.sleep(0.2)
        
        pyautogui.press("enter")
    else:
        pyautogui.press("enter")
        buttons["EditPatient"].clickButton()
        buttons["MCNum"].clickButton()
        buttons["MCNum"].clickButton()
        if patient["Hosp"] == "SGPH":
            pyautogui.typewrite(patient["MCNum"].replace(" ", "").replace("-", "").replace("/", ""))
            pyautogui.press("tab")
            pyautogui.typewrite(patient["MCRef"])
        else:
            MCNum = patient["MCNum"].replace(" ", "").replace("-", "").replace("/", "")
            pyautogui.typewrite(MCNum[:-1])
            pyautogui.press("tab")
            pyautogui.typewrite(MCNum[-1])

        buttons["SavePatient"].clickButton()
        buttons["OPV"].clickButton()

        time.sleep(0.2)
        while find_text("Processing..."):
            time.sleep(0.2)
        if find_text("Patient is eligible to claim for Medicare with details provided."):
            pyautogui.press("enter")
        elif find_subtext("not an exact match"):
            pyautogui.press("tab", 2)
            pyautogui.press("enter")
            buttons["Merge1"].clickButton()
            buttons["Merge2"].clickButton()
            buttons["Merge3"].clickButton()
            buttons["MergeSelected"].clickButton()
            pyautogui.press("tab", 2)
            pyautogui.press("enter")
            time.sleep(0.2)
            while find_text("Processing..."):
                time.sleep(0.2)
            
            pyautogui.press("enter")
        else:
            print("please correct yourself and return to demo screen and then press SHIFT+TAB")
            keyboard.wait("shift+tab", suppress=True)
            if skip:
                return 1
    if "HF" in patient and patient["HF"] == "DVA":
        buttons["OVV"].clickButton()
        time.sleep(0.2)
        while find_text("Processing..."):
            time.sleep(0.2)
        
        if find_text("Patient is known to DVA with details provided."):
            pyautogui.press("enter")
        elif find_subtext("not an exact match"):
            pyautogui.press("tab", 2)
            pyautogui.press("enter")
            buttons["Merge1"].clickButton()
            buttons["Merge2"].clickButton()
            buttons["Merge3"].clickButton()
            buttons["MergeSelected"].clickButton()
            pyautogui.press("tab", 2)
            pyautogui.press("enter")
            time.sleep(0.2)
            while find_text("Processing..."):
                time.sleep(0.2)
            
            pyautogui.press("enter")
        else:
            pyautogui.press("enter")
            buttons["EditPatient"].clickButton()
            buttons["DVANum"].clickButton()
            buttons["DVANum"].clickButton()
            DVANum = patient["DVANum"].replace(" ", "").replace("-", "").replace("/", "")
            pyautogui.typewrite(DVANum)
            
            buttons["SavePatient"].clickButton()
            buttons["OVV"].clickButton()

            time.sleep(0.2)
            while find_text("Processing..."):
                time.sleep(0.2)
            if find_text("Patient is known to DVA with details provided."):
                pyautogui.press("enter")
            elif find_subtext("not an exact match"):
                pyautogui.press("tab", 2)
                pyautogui.press("enter")
                buttons["Merge1"].clickButton()
                buttons["Merge2"].clickButton()
                buttons["Merge3"].clickButton()
                buttons["MergeSelected"].clickButton()
                pyautogui.press("tab", 2)
                pyautogui.press("enter")
                time.sleep(0.2)
                while find_text("Processing..."):
                    time.sleep(0.2)
                
                pyautogui.press("enter")
            else:
                print("please correct yourself and return to demo screen and then press SHIFT+TAB")
                keyboard.wait("shift+tab", suppress=True)
                if skip:
                    return 1
    else:
        buttons["PVF"].clickButton()
        time.sleep(0.2)
        while get_window_name() != "Please select the doctor and provider number to check":
            time.sleep(0.2)
        pyautogui.press("tab", 16)
        buttons["HFLoc"].clickButton()

        HosLocations = {"EHC": 0, "POWP": 3, "SGPH": 4, "SHC": 7}
        pyautogui.press("down", HosLocations[patient["Hosp"]])
        buttons["HFAccept"].clickButton()

        time.sleep(0.5)
        while find_text("Checking details") and not find_text("No healthfund code has been set."):
            time.sleep(0.5)

        if find_text("0 - Patient is known to the Health Fund specified in the request."):
            pyautogui.press("enter")
        else:
            pyautogui.press("enter")
            buttons["EditPatient"].clickButton()
            buttons["HFNum"].clickButton()
            buttons["HFNum"].clickButton()
            HFNum = patient["HFNum"].replace(" ", "").replace("-", "").replace("/", "")
            pyautogui.typewrite(HFNum)
            buttons["HF"].clickButton()
            
            hfPosition = {"AMA": 4, "AHM": 7, "AUH": 8, "BUP": 9, "CBH": 13, "CHF": 16, "DHF": 19, "ESH": 22, "FHI": 25, "GMH": 28, "HBF": 34, "HCF": 35, "HIF": 38, "SPS": 39, "HEA": 41, "LHS": 45, "MBF": 48, "MBP": 50, "MPL": 50, "NHB": 54, "NIB": 56, "OMF": 60, "LHM": 61, "PHF": 62, "POL": 63, "RTH": 65, "RBH": 67, "SLM": 72, "TFH": 73, "WFD": 78}
            if "HF" in patient:
                hf = patient["HF"]
                if len(hf) == 5 and hf[0:2] == "HF":
                    hf = hf[2:]
                if hf in hfPosition:
                    pyautogui.press("up", 80)
                    pyautogui.press("down", hfPosition[hf])
                    pyautogui.press("enter")
                else:
                    print(f"{patient["HF"]} currently has no data")

            buttons["SavePatient"].clickButton()
            buttons["PVF"].clickButton()
            time.sleep(0.2)
            while get_window_name() != "Please select the doctor and provider number to check":
                time.sleep(0.2)
            pyautogui.press("tab", 16)
            buttons["HFLoc"].clickButton()

            HosLocations = {"EHC": 0, "POWP": 3, "SGPH": 4, "SHC": 7}
            pyautogui.press("down", HosLocations[patient["Hosp"]])
            buttons["HFAccept"].clickButton()
            time.sleep(0.2)
            while find_text("Checking details") and not find_text("No healthfund code has been set."):
                time.sleep(0.2)

            if find_text("0 - Patient is known to the Health Fund specified in the request."):
                pyautogui.press("enter")
            else:
                print("please correct yourself and return to demo screen and then press SHIFT+TAB")
                keyboard.wait("shift+tab", suppress=True)
                if skip:
                    return 1

    buttons["Bookings"].clickButton()

    # print("Press ENTER Once Bookings Loaded")
    # keyboard.wait("ENTER", suppress=True)
    time.sleep(3)

    procdate = patient["ProcDate"]
    buttons["StartDate"].clickButton()
    pyautogui.typewrite(procdate)
    buttons["EndDate"].clickButton()
    pyautogui.typewrite(procdate)
    buttons["Refresh"].clickButton()
    time.sleep(0.5)
    if not find_text("Row 1"):
        buttons["Terminal"].clickButton()
        inp = input("Is procedure date correct? y/n")
        while inp != "y" and inp != "n":
            inp = input("Is procedure date correct? y/n")
        if inp == "y":
            buttons["NewBooking"].clickButton()
            time.sleep(0.2)
            while get_window_name() != "Booking Form":
                time.sleep(0.2)
            move_window()
            pyautogui.typewrite("remote monitor billing")
            pyautogui.press("tab")
            pyautogui.typewrite(procdate)
            pyautogui.press("tab")
            pyautogui.typewrite("8")
            pyautogui.press("tab", 6)
            pyautogui.typewrite("s")
            pyautogui.press("tab")
            pyautogui.press("down", 3)
            pyautogui.press("tab")
            pyautogui.typewrite("s")
            pyautogui.press("tab", 12)
            pyautogui.press("enter")
            
            time.sleep(0.5)
            if get_window_name() == "No Referral Selected":
                pyautogui.press("enter")
            time.sleep(0.2)
            while get_window_name() == "Booking Form":
                time.sleep(0.2)
            if get_window_name() == "Booking out of operating hours":
                pyautogui.press("enter")
            
        else:
            procdate = input("Please insert true date")
            buttons["StartDate"].clickButton()
            pyautogui.typewrite(procdate)
            buttons["EndDate"].clickButton()
            pyautogui.typewrite(procdate)

    buttons["Booking1"].clickButton()
    buttons["Booking1"].clickButton()
    time.sleep(0.2)
    while get_window_name() != "Booking Form":
        time.sleep(0.2)
    move_window()

    if find_text("Edit"):
        print("Patient already processed")
        skip_patient()
        keyboard.wait("shift+tab", suppress=True)
        if skip:
            pyautogui.hotkey("alt", "f4")
            return 0

    if find_text("Row 1"):
        buttons["Referral"].clickButton()
        buttons["SaveAndCloseBooking"].clickButton()
        time.sleep(0.5)
        if get_window_name() == "No Referral Selected":
            pyautogui.press("tab")
            pyautogui.press("enter")
            buttons["Referral"].clickButton()
            buttons["SaveAndCloseBooking"].clickButton()
            time.sleep(0.5)
            if get_window_name() == "No Referral Selected":
                print("Error: No Referral, press SHIFT+TAB once fixed")
                keyboard.wait("shift+tab", suppress=True)
                if skip:
                    pyautogui.hotkey("alt", "f4")
                    return 1
        
        while get_window_name() == "Booking Form":
            time.sleep(0.2)
        if get_window_name() == "Booking out of operating hours":
            pyautogui.press("enter")
            

        buttons["Booking1"].clickButton()
        buttons["Booking1"].clickButton()
        while get_window_name() != "Booking Form":
            time.sleep(0.2)
        move_window()
    else:
        buttons["NewIn"].clickButton()
        pyautogui.typewrite(procdate)
        pyautogui.press("tab")
        pyautogui.typewrite(procdate)
        pyautogui.press("tab", 2)
        pyautogui.typewrite(procdate)
        pyautogui.press("tab", 15)
        pyautogui.typewrite("PROC-SS")
        pyautogui.press("tab", 12)
        pyautogui.typewrite("wss")
        pyautogui.press("tab", 34)

        ref = patient["Referral"]
        referral_json = Path("ref.json")
        referral_db = load_json(referral_json)

        if ref in referral_db.keys():
            if referral_db[ref] == "-1":
                ref = ref.translate(str.maketrans("", "", string.punctuation))
                ref = ref.lower().split(" ")
                if len(ref) >= 2:
                    pyautogui.typewrite(ref[1])
                    pyautogui.press("tab")
                    pyautogui.typewrite(ref[0])
                    pyautogui.press("enter")
                else:
                    pyautogui.typewrite(ref[0])
                    pyautogui.press("tab")
                    pyautogui.press("enter")

                print("Choose correct Referral ")
                print("Press SHIFT+TAB when finished")
                keyboard.wait("shift+tab", suppress=True)
                if skip:
                    pyautogui.hotkey("alt", "f4")
                    return 1
                # buttons["Terminal"].clickButton()
                inp = input("Do you want to save details? y/n")
                while inp != "y" and inp != "n":
                    inp = input("Do you want to save details? y/n")
                
                time.sleep(0.5)
                count = 0
                while get_window_name() != "Add/Modify Referral":
                    count += 1
                    pyautogui.keyDown("alt")
                    pyautogui.press("tab", count) 
                    pyautogui.keyUp("alt")
                    time.sleep(0.5) 
                if inp == "y":
                    line = find_subtext("Provider No:")
                    if line == 0:
                        print("Couldn't save doctor")
                    else: 
                        sections = line.split()
                        referral_db[patient["Referral"]] = sections[sections.index("No:")+1]
            else:
                pyautogui.typewrite(referral_db[patient["Referral"]])
                pyautogui.press("tab")
                pyautogui.press("enter")
                pyautogui.press("tab", 43)
                pyautogui.press("enter")
        else:
            if ref[0].isdigit():
                pyautogui.typewrite(ref)
                pyautogui.press("tab")
                pyautogui.press("enter")
                time.sleep(0.5)
                if find_text("Row 1"):
                    pyautogui.press("tab", 43)
                    pyautogui.press("enter")
                else:
                    print("Couldn't find referral, try again and return to referral form and press SHIFT+TAB")
                    keyboard.wait("shift+tab", suppress=True)
                    if skip:
                        pyautogui.hotkey("alt", "f4")
                        return 1
            else:
                ref = ref.translate(str.maketrans("", "", string.punctuation))
                ref = ref.lower().split(" ")

                if len(ref) >= 2:
                    pyautogui.typewrite(ref[1])
                    pyautogui.press("tab")
                    pyautogui.typewrite(ref[0])
                    pyautogui.press("enter")
                else:
                    pyautogui.typewrite(ref[0])
                    pyautogui.press("tab")
                    pyautogui.press("enter")

                if not find_text("Row 2") and find_text("Row 1"):
                    pyautogui.press("tab", 43)
                    pyautogui.press("enter")
                else:
                    print("Choose correct Referral and press SHIFT+TAB on Referral window")
                    keyboard.wait("shift+tab", suppress=True)
                    if skip:
                        pyautogui.hotkey("alt", "f4")
                        return 1
        time.sleep(1)
        if three_month():
            pyautogui.press("tab", 24)
            pyautogui.press("down")
            pyautogui.press("tab", 19)
            pyautogui.press("enter")
        else:
            pyautogui.press("tab", 5)
            pyautogui.press("enter")
        
        save_json(referral_json, referral_db)
        time.sleep(0.5)
        buttons["Referral"].clickButton()
    time.sleep(0.2)
    lspnCount = [0,0]
    def prepareInvoice(buttons, patient, lspnCount):
        num = determineInvoice()
        buttons["InvoiceTo"].clickButton()
        
        if num == -1:
            print("Error, please make Invoice To Health Fund then press SHIFT+TAB")
            keyboard.wait("shift+tab", suppress=True)
            if skip:
                pyautogui.hotkey("alt", "f4")
                pyautogui.hotkey("alt", "f4")
                return 1
        else:
            pyautogui.press("down", num)
            if "HF" in patient and patient["HF"] == "DVA":
                pyautogui.press("up", 3)
            pyautogui.press("enter")

        buttons["ServiceLocation"].clickButton()
        
        if patient["Hosp"] == "EHC":
            pyautogui.typewrite("eastern heart clinic")
            lspnCount[0] = 4
            lspnCount[1] = 2
        elif patient["Hosp"] == "POWP":
            pyautogui.typewrite("prince of wales private hospital")
            lspnCount[0] = 11
        elif patient["Hosp"] == "SGPH":
            pyautogui.typewrite("st george private hospital")
            lspnCount[0] = 5
            lspnCount[1] = 13
        elif patient["Hosp"] == "SHC":
            pyautogui.typewrite("sutherland heart clinic")
            lspnCount[0] = 16
            lspnCount[1] = 10
        else:
            print("Error, please put in correct hospital then press SHIFT+TAB")
            keyboard.wait("shift+tab", suppress=True)
            if skip:
                pyautogui.hotkey("alt", "f4")
                pyautogui.hotkey("alt", "f4")
                return 1

        pyautogui.press("enter")
        pyautogui.press("tab")
        pyautogui.typewrite("singa")

        buttons["PrintNow"].clickButton()

    
    cathAndToe = False
    earlier = 0
    earlierPM = False
    later = 0
    laterPM = False
    if "55118" in patient["ProcNumbers"] and "61109" in patient["ProcNumbers"]:
        cathAndToe = True
        if "Times" in patient and len(patient["Times"]) == 2:
            time_1 = "".join(ch for ch in patient["Times"][0] if ch.isdigit())
            time_2 = "".join(ch for ch in patient["Times"][1] if ch.isdigit())
        else:
            inp = input("Confirm that 55118 and 61109 are required y/n")
            while inp != "y" and inp != "n":
                inp = input("Confirm that 55118 and 61109 are required y/n")
            
            if inp == "y":
                time1 = input("Put in first time with no need for punctuation or ':")
                time2 = input("Put in second time with no need for punctuation or ':")
            else:
                cathAndToe = False
        
        if cathAndToe:
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
            
            if int(earlier[0]) >= 12:
                if int(earlier[0]) != 12:
                    earlier[0] = str(int(earlier[0]) - 12)
                earlierPM = True
            if int(later[0]) >= 12:
                if int(later[0]) != 12:
                    later[0] = str(int(later[0]) - 12)
                laterPM = True
            
            buttons["PatientWindow"].clickButton()
            buttons["Accounts"].clickButton()
            buttons["CreateInvoice"].clickButton()
            time.sleep(1)
            while get_window_name() == "HealthTrack":
                time.sleep(0.2)
            
            prepareInvoice(buttons, patient, lspnCount)
            buttons["ServiceDate"].clickButton()
            time.sleep(0.2)
            pyautogui.typewrite(procdate)
            
            count = 0
            while True:
                buttons["FindReferral"].clickButton()
                if count > 0:
                    pyautogui.press("down", count)
                pyautogui.press("enter")
                referral_dates = referral_date_range()
                if referral_dates[14] != "-":
                    higherDate = [int(referral_dates[14:16]), int(referral_dates[17:19]), int(referral_dates[20:24])]
                    proc = re.split(r"[\/\.\-]+", procdate)
                    if len(proc[2]) == 2:
                        year = int(proc[2]) + 2000
                    else:
                        year = int(proc[2])
                    if year > higherDate[2]:
                        count += 1
                        continue
                    elif year == higherDate[2] and int(proc[1]) > higherDate[1]:
                        count += 1
                        continue
                    elif year == higherDate[2] and int(proc[1]) == higherDate[1] and int(proc[0]) > higherDate[0]:
                        count += 1
                        continue
                lowerDate = [int(referral_dates[0:2]), int(referral_dates[3:5]), int(referral_dates[6:10])]
                proc = re.split(r"[\/\.\-]+", procdate)
                if len(proc[2]) == 2:
                    year = int(proc[2]) + 2000
                else:
                    year = int(proc[2])
                if year < lowerDate[2]:
                    count += 1
                    continue
                elif year == lowerDate[2] and int(proc[1]) < lowerDate[1]:
                    count += 1
                    continue
                elif year == lowerDate[2] and int(proc[1]) == lowerDate[1] and int(proc[0]) < lowerDate[0]:
                    count += 1
                    continue
                
                break

            buttons["ItemNumber"].clickButton()
            pyautogui.typewrite("55118")
            buttons["ThreeDots"].clickButton()
            pyautogui.press("tab")
            pyautogui.typewrite(earlier[0])
            pyautogui.press("right")
            pyautogui.typewrite(earlier[1])
            currentTime = datetime.now()
            if earlierPM:
                pyautogui.press("right")
                pyautogui.press("up")
            buttons["CheckBox"].clickButton()
            pyautogui.press("tab", 19)
            pyautogui.press("enter")
            buttons["ItemNumber"].clickButton()
            pyautogui.press("enter", 2)
            buttons["TopLPSN"].clickButton()
            pyautogui.press("down", lspnCount[1])
            pyautogui.press("tab", 3)
            pyautogui.press("enter", 2)



            print("Confirm billing is correct by pressing SHIFT+TAB")
            keyboard.wait("shift+tab", suppress=True)
            if skip:
                pyautogui.hotkey("alt", "f4")
                pyautogui.hotkey("alt", "f4")
                return 1

            buttons["Issue"].clickButton()
            time.sleep(1)
            pyautogui.hotkey("alt", "f4")
            patient["ProcNumbers"].remove("55118")
            time.sleep(0.2)
            count = 0
            while get_window_name() != "Booking Form":
                count += 1
                pyautogui.keyDown("alt")
                pyautogui.press("tab", count) 
                pyautogui.keyUp("alt")
                time.sleep(0.2) 

    time.sleep(3)
    num = booking_status()
    if num == 0:
        print("Error, please make closed then press SHIFT+TAB")
        keyboard.wait("shift+tab", suppress=True)
        if skip:
            pyautogui.hotkey("alt", "f4")
            return 1
    else:
        for i in range(num):
            buttons["Status"].clickButton()

    buttons["SaveAndCloseBooking"].clickButton()
    time.sleep(0.5)
    if get_window_name() == "No Referral Selected":
        pyautogui.press("tab")
        pyautogui.press("enter")
        buttons["Referral"].clickButton()
        buttons["SaveAndCloseBooking"].clickButton()
    time.sleep(0.2)
    while get_window_name() == "Booking Form":
        time.sleep(0.2)
    if get_window_name() == "Booking out of operating hours":
            pyautogui.press("enter")
    
    prepareInvoice(buttons, patient, lspnCount)

    buttons["ItemNumber"].clickButton()

    
        
    if isinstance(patient["ProcNumbers"], list):
        for itemNumber in patient["ProcNumbers"]:
            itemNumber = itemNumber.translate(str.maketrans("", "", string.punctuation))
            pyautogui.typewrite(itemNumber)
            if itemNumber == "110" or itemNumber == "116":
                pyautogui.typewrite("H")
            if cathAndToe and itemNumber == "61109":
                buttons["ThreeDots"].clickButton()
                pyautogui.press("tab")
                pyautogui.typewrite(later[0])
                pyautogui.press("right")
                pyautogui.typewrite(later[1])
                currentTime = datetime.now()
                if laterPM:
                    pyautogui.press("right")
                    pyautogui.press("up")
                buttons["CheckBox"].clickButton()
                pyautogui.press("tab", 19)
                pyautogui.press("enter")
                buttons["ItemNumber"].clickButton()
            pyautogui.press("enter", 2)
            if itemNumber == "61109":
                buttons["TopLPSN"].clickButton()
                pyautogui.press("down", lspnCount[0])
                pyautogui.press("tab", 3)
                pyautogui.press("enter")

            elif itemNumber == "55118":
                buttons["TopLPSN"].clickButton()
                pyautogui.press("down", lspnCount[1])
                pyautogui.press("tab", 3)
                pyautogui.press("enter")
    else:
        itemNumber = patient["ProcNumbers"]
        itemNumber = itemNumber.translate(str.maketrans("", "", string.punctuation))
        pyautogui.typewrite(itemNumber)
        if itemNumber == "110" or itemNumber == "116":
            pyautogui.typewrite("H")
        if cathAndToe and itemNumber == "61109":
            buttons["ThreeDots"].clickButton()
            pyautogui.press("tab")
            pyautogui.typewrite(later[0])
            pyautogui.press("right")
            pyautogui.typewrite(later[1])
            currentTime = datetime.now()
            if laterPM and currentTime.hour < 12 or (not laterPM) and currentTime.hour >= 12:
                pyautogui.press("right")
                pyautogui.press("up")
            buttons["CheckBox"].clickButton()
            pyautogui.press("tab", 19)
            pyautogui.press("enter")
            buttons["ItemNumber"].clickButton()
        pyautogui.press("enter", 2)
        if itemNumber == "61109":
            buttons["TopLPSN"].clickButton()
            pyautogui.press("down", lspnCount[0])
            pyautogui.press("tab", 3)
            pyautogui.press("enter")

        elif itemNumber == "55118":
            buttons["TopLPSN"].clickButton()
            pyautogui.press("down", lspnCount[1])
            pyautogui.press("tab", 3)
            pyautogui.press("enter")

    buttons["ApplyRules"].clickButton()
    buttons["AcceptRules"].clickButton()

    print("Confirm billing is correct by pressing SHIFT+TAB")
    keyboard.wait("shift+tab", suppress=True)
    if skip:
        pyautogui.hotkey("alt", "f4")
        pyautogui.hotkey("alt", "f4")
        return 1

    buttons["Issue"].clickButton()
    time.sleep(0.5)
    if get_window_name() == "Invalid Item Ordering":
        pyautogui.press("enter")
    
    return 0
