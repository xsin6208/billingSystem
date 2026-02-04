import pyautogui
import json
import keyboard
from pathlib import Path
import time
import win32gui
from pywinauto import Desktop


class ButtonPress():
    def __init__(self, location):
        self.location = location
    
    def clickButton(self):
        pyautogui.moveTo(self.location[0], self.location[1], duration=0.1)
        pyautogui.click()

    def dragToButton(self):
        pyautogui.dragTo(self.location[0], self.location[1], duration=0.1)

def jsonl_iter(path):
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                yield json.loads(line)

def move_window():
        hwnd = win32gui.GetForegroundWindow()
        desktop = Desktop(backend="win32")

        # Get the foreground window by handle
        win = desktop.window(handle=hwnd)
        win.move_window(x=100, y=100)

def configure():

    buttons = dict()

    
    configure = False
    configFile = "config.jsonl"
    path = Path(configFile)
    
    if configure:
        if path.exists() and path.stat().st_size > 0:
            path.write_text("")
    
    if configure or not path.exists():
        print("System needs to configure")
        print("Press SHIFT+TAB when ready")
        keyboard.wait("shift+tab", suppress=True)
        print("Hover over the Terminal Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["Terminal"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the Patient Window, then press ENTER...")
        keyboard.wait("enter", suppress=True)

        buttons["PatientWindow"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        buttons["PatientWindow"].clickButton()


        print("Hover over the FIND PATIENT Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["FindPatient"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the DEMOGRAPHIC Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["Demo"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        buttons["Demo"].clickButton()

        print("Hover over the EDIT PATIENT Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["EditPatient"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the NEW PATIENT Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["SavePatient"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the OPV Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["OPV"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the MEDICARE NUMBER, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["MCNum"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        print("Hover over the MERGE CHECKBOX 1, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["Merge1"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        print("Hover over the MERGE CHECKBOX 2, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["Merge2"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        print("Hover over the MERGE CHECKBOX 3, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["Merge3"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        print("Hover over the MERGE SELECTED Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["MergeSelected"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the PVF Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["PVF"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        buttons["PVF"].clickButton()

        print("Hover over the EHC LOCATION, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["HFLoc"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the ACCEPT Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["HFAccept"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        pyautogui.hotkey(["alt", "f4"])

        print("Hover over the HEALTH FUND NUMBER, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["HFNum"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the HEALTH FUND, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["HF"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the OVV Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["OVV"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        print("Hover over the DVA Number, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["DVANumber"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        print("Hover over the Documents Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["Documents"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the BOOKINGS Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["Bookings"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        buttons["Bookings"].clickButton()

        print("Hover over the START DATE Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["StartDate"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the END DATE Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["EndDate"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the REFRESH Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["Refresh"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the FIRST BOOKING, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["Booking1"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        print("Hover over the NEW BOOKING, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["NewBooking"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        buttons["Booking1"].clickButton()
        buttons["Booking1"].clickButton()


        time.sleep(4)
        move_window()

        print("Hover over the FIRST REFERRAL, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["Referral"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the SAVE AND CLOSE Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["SaveAndCloseBooking"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the NEW IN, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["NewIn"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the CONFIRMED/ARRIVED Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["Status"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the InvoiceTo, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["InvoiceTo"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the SERVICE LOCATION, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["ServiceLocation"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the PRINT NOW, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["PrintNow"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        print("Hover over the SERVICE DATE, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["ServiceDate"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the ITEM NUMBER, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["ItemNumber"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the THREE DOTS, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["ThreeDots"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the CHECK BOX, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["CheckBox"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")

        print("Hover over the TOP LSPN, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["TopLPSN"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        print("Hover over the APPLY RULES Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["ApplyRules"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        print("Hover over the ACCEPT RULES Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["AcceptRules"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        print("Hover over the FIND REFERRAL Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["FindReferral"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        print("Hover over the ISSUE Button, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["Issue"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        
        print("Hover over the ACCOUNTS Tab, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["Accounts"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")
        buttons["Accounts"].clickButton()

        print("Hover over CREATE INVOICE, then press ENTER...")
        keyboard.wait("enter", suppress=True)
        buttons["CreateInvoice"] = ButtonPress((pyautogui.position().x, pyautogui.position().y))
        with open(configFile, "a", encoding="utf-8") as f:
            json.dump(pyautogui.position(), f)
            f.write("\n")



    else:
        it = jsonl_iter(configFile)

        buttons["Terminal"] = ButtonPress(next(it))

        buttons["PatientWindow"] = ButtonPress(next(it))

        buttons["FindPatient"] = ButtonPress(next(it))

        buttons["Demo"] = ButtonPress(next(it))

        buttons["EditPatient"] = ButtonPress(next(it))

        buttons["SavePatient"] = ButtonPress(next(it))

        buttons["OPV"] = ButtonPress(next(it))

        buttons["MCNum"] = ButtonPress(next(it))

        buttons["Merge1"] = ButtonPress(next(it))

        buttons["Merge2"] = ButtonPress(next(it))

        buttons["Merge3"] = ButtonPress(next(it))

        buttons["MergeSelected"] = ButtonPress(next(it))

        buttons["PVF"] = ButtonPress(next(it))

        buttons["HFLoc"] = ButtonPress(next(it))

        buttons["HFAccept"] = ButtonPress(next(it))

        buttons["HFNum"] = ButtonPress(next(it))

        buttons["HF"] = ButtonPress(next(it))

        buttons["OVV"] = ButtonPress(next(it))

        buttons["DVANumber"] = ButtonPress(next(it))

        buttons["Documents"] = ButtonPress(next(it))

        buttons["Bookings"] = ButtonPress(next(it))

        buttons["StartDate"] = ButtonPress(next(it))

        buttons["EndDate"] = ButtonPress(next(it))

        buttons["Refresh"] = ButtonPress(next(it))

        buttons["Booking1"] = ButtonPress(next(it))

        buttons["NewBooking"] = ButtonPress(next(it))

        buttons["Referral"] = ButtonPress(next(it))

        buttons["SaveAndCloseBooking"] = ButtonPress(next(it))

        buttons["NewIn"] = ButtonPress(next(it))

        buttons["Status"] = ButtonPress(next(it))

        buttons["InvoiceTo"] = ButtonPress(next(it))

        buttons["ServiceLocation"] = ButtonPress(next(it))

        buttons["PrintNow"] = ButtonPress(next(it))

        buttons["ServiceDate"] = ButtonPress(next(it))

        buttons["ItemNumber"] = ButtonPress(next(it))
    
        buttons["ThreeDots"] = ButtonPress(next(it))

        buttons["CheckBox"] = ButtonPress(next(it))

        buttons["TopLPSN"] = ButtonPress(next(it))

        buttons["ApplyRules"] = ButtonPress(next(it))

        buttons["AcceptRules"] = ButtonPress(next(it))

        buttons["FindReferral"] = ButtonPress(next(it))

        buttons["Issue"] = ButtonPress(next(it))

        buttons["Accounts"] = ButtonPress(next(it))

        buttons["CreateInvoice"] = ButtonPress(next(it))

        
    
    print("System is configured")
    print("Close any extra windows to diary and patient window and then press SHIFT+TAB")
    keyboard.wait("shift+tab", suppress=True)
    return buttons