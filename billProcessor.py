import pyautogui
import time
import keyboard
import string
from pathlib import Path
from configureButtons import configure, ButtonPress
import re
import winsound
import builtins
from helperFunctions import load_json, save_json, get_window_name, find_subtext, find_text, three_month, move_window, booking_status, determineInvoice, referral_date_range, switch_to_window, orderTimes

skip = False

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

def find_patient(patient, buttons):
    global skip
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

def merge_differences(buttons):
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

def edit_mc(patient, buttons):
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

def mc_check(patient, buttons):
    global skip
    time.sleep(0.2)
    while find_text("Processing..."):
        time.sleep(0.2)
    
    if find_text("Patient is eligible to claim for Medicare with details provided."):
        pyautogui.press("enter")
    elif find_subtext("not an exact match"):
        merge_differences(buttons)
    else:
        pyautogui.press("enter")
        edit_mc(patient, buttons)
        buttons["OPV"].clickButton()

        time.sleep(0.2)
        while find_text("Processing..."):
            time.sleep(0.2)
        if find_text("Patient is eligible to claim for Medicare with details provided."):
            pyautogui.press("enter")
        elif find_subtext("not an exact match"):
            merge_differences(buttons)
        else:
            print("please correct yourself and return to demo screen and then press SHIFT+TAB")
            keyboard.wait("shift+tab", suppress=True)
            if skip:
                return 1

def edit_dva(patient, buttons):
    buttons["EditPatient"].clickButton()
    buttons["DVANum"].clickButton()
    buttons["DVANum"].clickButton()
    DVANum = patient["DVANum"].replace(" ", "").replace("-", "").replace("/", "")
    pyautogui.typewrite(DVANum)
    
    buttons["SavePatient"].clickButton()

def dva_check(patient, buttons):
    global skip
    buttons["OVV"].clickButton()
    time.sleep(0.2)
    while find_text("Processing..."):
        time.sleep(0.2)
    
    if find_text("Patient is known to DVA with details provided."):
        pyautogui.press("enter")
    elif find_subtext("not an exact match"):
        merge_differences(buttons)
    else:
        pyautogui.press("enter")
        
        edit_dva(patient, buttons)
        buttons["OVV"].clickButton()

        time.sleep(0.2)
        while find_text("Processing..."):
            time.sleep(0.2)
        if find_text("Patient is known to DVA with details provided."):
            pyautogui.press("enter")
        elif find_subtext("not an exact match"):
            merge_differences(buttons)
        else:
            print("please correct yourself and return to demo screen and then press SHIFT+TAB")
            keyboard.wait("shift+tab", suppress=True)
            if skip:
                return 1

def hf_check_provider(patient, buttons):
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

def edit_hf(patient, buttons):
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

def hf_check(patient, buttons):
    global skip
    hf_check_provider(patient, buttons)
    if find_text("0 - Patient is known to the Health Fund specified in the request."):
        pyautogui.press("enter")
    else:
        pyautogui.press("enter")
        edit_hf(patient, buttons)
        hf_check_provider(patient, buttons)

        if find_text("0 - Patient is known to the Health Fund specified in the request."):
            pyautogui.press("enter")
        else:
            print("please correct yourself and return to demo screen and then press SHIFT+TAB")
            keyboard.wait("shift+tab", suppress=True)
            if skip:
                return 1

def check_numbers(patient, buttons):
    time.sleep(0.2)
    while get_window_name() != "HealthTrack":
        time.sleep(0.2)

    buttons["Demo"].clickButton()

    buttons["OPV"].clickButton()

    mc_check(patient, buttons)
    
    if "HF" in patient and patient["HF"] == "DVA":
        dva_check(patient, buttons)
    else:
        hf_check(patient, buttons)

def create_booking(patient, buttons):
    buttons["NewBooking"].clickButton()
    time.sleep(0.2)
    while get_window_name() != "Booking Form":
        time.sleep(0.2)
    move_window()
    pyautogui.typewrite("remote monitor billing")
    pyautogui.press("tab")
    pyautogui.typewrite(patient["ProcDate"])
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


def find_booking(patient, buttons):
    buttons["Bookings"].clickButton()
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
            create_booking(patient, buttons)
        else:
            procdate = input("Please insert true date")
            patient["ProcDate"] = procdate
            buttons["StartDate"].clickButton()
            pyautogui.typewrite(procdate)
            buttons["EndDate"].clickButton()
            pyautogui.typewrite(procdate)
            if not find_text("Row 1"):
                create_booking(patient, buttons)

    buttons["Booking1"].clickButton()
    buttons["Booking1"].clickButton()
    time.sleep(0.2)
    while get_window_name() != "Booking Form":
        time.sleep(0.2)
    move_window()

def referral_memory_search(patient, buttons, referral_db):
    global skip
    ref = patient["Referral"]
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

        inp = input("Do you want to save details? y/n")
        switch_to_window("Add/Modify Referral")
        while inp != "y" and inp != "n":
            inp = input("Do you want to save details? y/n")
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

def referral_alt_search(patient, buttons):
    global skip
    ref = patient["Referral"]
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

def determine_dr(patient, buttons):
    ref = patient["Referral"]
    referral_json = Path("ref.json")
    referral_db = load_json(referral_json)

    if ref in referral_db.keys():
        referral_memory_search(patient, buttons, referral_db)
    else:
        referral_alt_search(patient, buttons)
        
    save_json(referral_json, referral_db)

def find_referral(patient, buttons):
    if find_text("Row 1"):
        buttons["Referral"].clickButton()
    else:
        buttons["NewIn"].clickButton()
        pyautogui.typewrite(patient["ProcDate"])
        pyautogui.press("tab")
        pyautogui.typewrite(patient["ProcDate"])
        pyautogui.press("tab", 2)
        pyautogui.typewrite(patient["ProcDate"])
        pyautogui.press("tab", 15)
        pyautogui.typewrite("PROC-SS")
        pyautogui.press("tab", 12)
        pyautogui.typewrite("wss")
        pyautogui.press("tab", 34)

        determine_dr(patient, buttons)
        time.sleep(1)

        if three_month():
            pyautogui.press("tab", 24)
            pyautogui.press("down")
            pyautogui.press("tab", 19)
            pyautogui.press("enter")
        else:
            pyautogui.press("tab", 5)
            pyautogui.press("enter")
        
        
        time.sleep(0.5)
        buttons["Referral"].clickButton()
    time.sleep(0.2)

def getHospitalandLSPN(patient, buttons, lspnCount):
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

def prepareInvoice(patient, buttons, lspnCount):
    global skip
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
    getHospitalandLSPN(patient, buttons, lspnCount)
    pyautogui.press("tab")
    pyautogui.typewrite("singa")

    buttons["PrintNow"].clickButton()

def closeInvoice(buttons):
    global skip
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




def chooseCorrectInvoiceReferral(patient, buttons):
    count = 0
    while True:
        buttons["FindReferral"].clickButton()
        if count > 0:
            pyautogui.press("down", count)
        pyautogui.press("enter")
        referral_dates = referral_date_range()
        if referral_dates[14] != "-":
            higherDate = [int(referral_dates[14:16]), int(referral_dates[17:19]), int(referral_dates[20:24])]
            proc = re.split(r"[\/\.\-]+", patient["ProcDate"])
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
        proc = re.split(r"[\/\.\-]+", patient["ProcDate"])
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


def inputTimedItemNumber(buttons, lspnCount, time, pm, number):
    buttons["ThreeDots"].clickButton()
    pyautogui.press("tab")
    pyautogui.typewrite(time[0])
    pyautogui.press("right")
    pyautogui.typewrite(time[1])
    if pm:
        pyautogui.press("right")
        pyautogui.press("up")
    buttons["CheckBox"].clickButton()
    pyautogui.press("tab", 19)
    pyautogui.press("enter")
    buttons["ItemNumber"].clickButton()
    pyautogui.press("enter", 2)
    if number == "61109":
        buttons["TopLPSN"].clickButton()
        pyautogui.press("down", lspnCount[0])
    elif number == "55118":
        buttons["TopLPSN"].clickButton()
        pyautogui.press("down", lspnCount[1])
    pyautogui.press("tab", 3)
    pyautogui.press("enter")

def performCathAndToe(patient, buttons, lspnCount):
    if "55118" in patient["ProcNumbers"] and "61109" in patient["ProcNumbers"]:
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
                return False, [0,0], False

        earlier, later, earlierPM, laterPM = orderTimes(time_1, time_2)
        
        buttons["PatientWindow"].clickButton()
        buttons["Accounts"].clickButton()
        buttons["CreateInvoice"].clickButton()
        time.sleep(1)
        while get_window_name() == "HealthTrack":
            time.sleep(0.2)
        
        prepareInvoice(patient, buttons, lspnCount)

        buttons["ServiceDate"].clickButton()
        time.sleep(0.2)
        pyautogui.typewrite(patient["ProcDate"])
        
        chooseCorrectInvoiceReferral(patient, buttons)
        buttons["ItemNumber"].clickButton()
        pyautogui.typewrite("55118")
        inputTimedItemNumber(buttons, lspnCount, earlier, earlierPM, "55118")
        closeInvoice(buttons)
        
        pyautogui.hotkey("alt", "f4")
        patient["ProcNumbers"].remove("55118")
        
        switch_to_window("Booking Form")
        return True, later, laterPM

    return False, [0,0], False

def closeBooking(patient, buttons):
    global skip
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

def insertNumber(itemNumber, buttons, lspnCount, cathAndToe, later, laterPM):
    itemNumber = itemNumber.translate(str.maketrans("", "", string.punctuation))
    pyautogui.typewrite(itemNumber)
    if itemNumber == "110" or itemNumber == "116":
        pyautogui.typewrite("H")
    if cathAndToe and itemNumber == "61109":
            inputTimedItemNumber(buttons, lspnCount, later, laterPM, "61109")
    else:    
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
    

def processItems(patient, buttons, lspnCount, cathAndToe, later, laterPM):
    buttons["ItemNumber"].clickButton()
    
    if isinstance(patient["ProcNumbers"], list):
        for itemNumber in patient["ProcNumbers"]:
            insertNumber(itemNumber, buttons, lspnCount, cathAndToe, later, laterPM)
    else:
        itemNumber = patient["ProcNumbers"]
        insertNumber(itemNumber, buttons, lspnCount, cathAndToe, later, laterPM)
    closeInvoice(buttons) 

def processPatient(patient, buttons):
    global skip
    skip = False
    

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

    find_patient(patient, buttons)
    check_numbers(patient, buttons)
    find_booking(patient, buttons)

    if find_text("Edit"):
        print("Patient already processed")
        skip_patient()
        keyboard.wait("shift+tab", suppress=True)
        if skip:
            pyautogui.hotkey("alt", "f4")
            return 0

    find_referral(patient, buttons)
    lspnCount = [0,0]
    cathAndToe, later, laterPM = performCathAndToe(patient, buttons, lspnCount)
    time.sleep(3)
    closeBooking(patient, buttons)
    
    prepareInvoice(patient, buttons, lspnCount)

    processItems(patient, buttons, lspnCount, cathAndToe, later, laterPM)
    return 0
