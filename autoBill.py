from pathlib import Path
from fileToText import extract_data
from billProcessor import processPatient
from configureButtons import configure, ButtonPress
import keyboard
import json
import string
import re
import pyautogui
import win32gui
from pywinauto import Desktop
import time
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

currentPatient = None
collectionMode = True

def printAllDetails():
    global currentPatient
    print(currentPatient)

keyboard.add_hotkey("alt+ctrl+p", printAllDetails, suppress=True)

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

        # Get the foreground window by handle
        active_window = desktop.window(handle=hwnd)
        return active_window.window_text()

def correctDetails(patient, buttons):
    global collectionMode
    time.sleep(0.5)
    if get_window_name() != "Windows PowerShell":
        buttons["Terminal"].clickButton()
    if "PatientName" not in patient:
        if "PatientDOB" in patient:
            print(patient["PatientDOB"])
            patient["PatientName"] = input(f"Please put in patient LASTNAME, Firstname with DOB {patient["PatientDOB"]}")
        else:
            print(patient, "Couldn't find any patient associated to this information")
            return 1
    
    
    print(f"Patient: {patient["PatientName"]}")
    if "Times" in patient:
        print("Patient has high chance of having both 55118 and 61109")
    if "ProcDate" in patient and isinstance(patient["DerivedDate"], list):
        print("Patient may have two procedures")
    time.sleep(0.5)
    if get_window_name() != "Windows PowerShell":
        buttons["Terminal"].clickButton()
    inp = input("Do you want to process patient? y/n")
    while inp != "n" and inp != "y":
        inp = input("Do you want to process patient? y/n")
    if inp == "n":
        return 1
    
    if "PatientDOB" not in patient:
        patient["PatientDOB"] = input(f"Please put in DOB dd/mm/yyyy for {patient["PatientName"]}")

    if "Hosp" not in patient:
        if "HospSecondary" not in patient:
            patient["Hosp"] = "EHC"
        elif patient["HospSecondary"] == "Sutherland Heart Clinic":
            patient["Hosp"] = "SHC"
        elif patient["HospSecondary"] == "SGP":
            patient["Hosp"] = "SGPH"
        elif patient["HospSecondary"] == "POW":
            patient["Hosp"] = "POWP"
        else:
            if get_window_name() != "Windows PowerShell":
                buttons["Terminal"].clickButton()
            hosp = input("Hospital couldn't be found, type 'skip' to skip or provide hospital acronym in caps")
            if hosp == "skip":
                return 1
            patient["Hosp"] = hosp
    
    if "DerivedDate" in patient:
        if "DateSecondary" in patient:
            dateSec = re.split(r"[\/\.\-\s]+", patient["DateSecondary"])
            derivedDate = re.split(r"[\/\.\-\s]+", patient["DerivedDate"])
            if int(dateSec[0]) == int(derivedDate[0]) and len(derivedDate) > 1 \
            and int(dateSec[1]) == int(derivedDate[1]):
                patient["ProcDate"] = patient["DateSecondary"]
            else:
                inp = input(f"Is {patient["DerivedDate"]} correct? y/n")
                if inp == "n":
                    inp = input(f"is {patient["DateSecondary"]} correct? If not input correct date")
                    if inp == "y":
                        patient["ProcDate"] = patient["DateSecondary"] 
                    else:
                        patient["ProcDate"] = inp
                else:
                    derivedDateFormatted = derivedDate[0] + "/" + derivedDate[1] + "/" + derivedDate[2]
                    patient["ProcDate"] = derivedDateFormatted
    else:
        if "DateSecondary" in patient:
            inp = input(f"is {patient["DateSecondary"]} correct? If not input correct date")
            if inp == "y":
                patient["ProcDate"] = patient["DateSecondary"] 
            else:
                patient["ProcDate"] = inp
    
    referral_json = Path("ref.json")
    referral_db = load_json(referral_json)
    if "DerivedReferral" not in patient or patient["DerivedReferral"] == "HT":
        print("Please paste in referral (preferrably number)")
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
        pyautogui.press("enter", 2)
        time.sleep(2)
        buttons["Documents"].clickButton()
        inp = input()
        patient["Referral"] = inp
    else:
        if patient["DerivedReferral"] in referral_db:
            patient["Referral"] = patient["DerivedReferral"]
        elif collectionMode:
            inp = input(f"Is {patient["DerivedReferral"]} correct? y/n")
            if inp == "y":
                referral_db[patient["DerivedReferral"]] = "-1"
                patient["Referral"] = patient["DerivedReferral"]
            else:
                inp = input("Please paste in correct referral")
                patient["Referral"] = inp
        else:
            # need to make a similarity checker
            patient["Referral"] = patient["DerivedReferral"]

    save_json(referral_json, referral_db)

    items_json = Path("itemNums.json")
    items_db = load_json(items_json)
    if collectionMode:
        if "ProcNumbers" not in patient:
            patient["ProcNumbers"] = list()
            inp = input("Enter in numbers until all in, then enter y")
            while inp != "y":
                patient["ProcNumbers"].append(inp)
                if inp not in items_db:
                    items_db[inp] = inp
        elif isinstance(patient["ProcNumbers"], list):
            for i, number in enumerate(patient["ProcNumbers"]):
                number = number.translate(str.maketrans("", "", string.punctuation))
                number = re.sub(r"\D", "", number)
                if number in items_db:
                    if items_db[number] != number:
                        inp = input(f"Is {items_db[number]} supposed to be instead of {number}? Enter y or enter the true number")
                        if inp == "y":
                            patient["ProcNumbers"][i] = items_db[number]
                        else:
                            patient["ProcNumbers"][i] = inp
                    else:
                        patient["ProcNumbers"][i] = number
                else:
                    inp = input(f"Enter 'y' if {number} is correct, if not enter true number")
                    if inp == "y":
                        items_db[number] = number
                        patient["ProcNumbers"][i] = number
                    else:
                        patient["ProcNumbers"][i] = inp
                        items_db[number] = inp
        else:
            number = patient["ProcNumbers"].translate(str.maketrans("", "", string.punctuation))
            number = re.sub(r"\D", "", number)
            if number in items_db:
                if items_db[number] != number:
                    inp = input(f"Is {items_db[number]} supposed to be instead of {number}? Enter y or enter the true number")
                    if inp == "y":
                        patient["ProcNumbers"] = items_db[number]
                    else:
                        patient["ProcNumbers"] = inp
                else:
                    patient["ProcNumbers"] = number
            else:
                inp = input(f"Enter 'y' if {number} is correct, if not enter true number")
                if inp == "y":
                    items_db[number] = number
                    patient["ProcNumbers"] = number
                else:
                    patient["ProcNumbers"] = inp
                    items_db[number] = inp
    else:
        if "ProcNumbers" not in patient:
            patient["ProcNumbers"] = list()
            inp = input("Enter in numbers until all in, then enter y")
            while inp != "y":
                patient["ProcNumbers"].append(inp)
        elif isinstance(patient["ProcNumbers"], list):
            for i, number in enumerate(patient["ProcNumbers"]):
                number = number.translate(str.maketrans("", "", string.punctuation))
                number = re.sub(r"\D", "", number)
                if number in items_db:
                    patient["ProcNumbers"][i] = items_db[number]
                else:
                    inp = input(f"Error: {number}, enter true number")
                    patient["ProcNumbers"][i] = inp
        else:
            number = patient["ProcNumbers"].translate(str.maketrans("", "", string.punctuation))
            number = re.sub(r"\D", "", number)
            if number in items_db:
                patient["ProcNumbers"] = items_db[number]
            else:
                inp = input(f"Error: {number}, enter true number")
                patient["ProcNumbers"] = inp
        
        if "Times" in patient:
            if "61109" not in patient["ProcNumbers"] and "55118" in patient["ProcNumbers"]:
                inp = input("Is 61109 missing? y/n")
                while inp != "y" and inp != "n":
                    inp = input()
                if inp == "y":
                    patient["ProcNumbers"].append("61109")
            elif "61109" in patient["ProcNumbers"] and "55118" not in patient["ProcNumbers"]:
                inp = input("Is 55118 missing? y/n")
                while inp != "y" and inp != "n":
                    inp = input()
                if inp == "y":
                    patient["ProcNumbers"].append("55118")
    save_json(items_json, items_db)
    
    if "Times" in patient:
        if isinstance(patient["Times"], list) and len(patient["Times"]) == 2:
            time_1 = "".join(ch for ch in patient["Times"][0] if ch.isdigit())
            time_2 = "".join(ch for ch in patient["Times"][1] if ch.isdigit())
            if len(time_1) < 3 or len(time_1) > 4 or len(time_1) < 4 and int(time_1[0]) < 8 or len(time_1) > 3 and int(time_1[1]) > 7:
                inp = input(f"Is {time_1} correct? y/n")
                while inp != "y" and inp != "n":
                    inp = input()
                if inp == "y":
                    patient["Times"][0] = time_1
                else:
                    patient["Times"][0] = input("Put in correct Time")
            else:
                patient["Times"][0] = time_1
            
            if len(time_2) < 3 or len(time_2) > 4 or len(time_2) < 4 and int(time_2[0]) < 8 or len(time_2) > 3 and int(time_2[1]) > 7:
                inp = input(f"Is {time_2} correct? y/n")
                while inp != "y" and inp != "n":
                    inp = input()
                if inp == "y":
                    patient["Times"][1] = time_2
                else:
                    patient["Times"][1] = input("Put in correct Time")
            else:
                patient["Times"][1] = time_2

    print(patient)
    inp = input("Patient ready, do you want to continue? y/n")
    if inp == "n":
        return 1
    return 0

folder = Path(r"C:\Users\xsing\OneDrive\Desktop\Billing")

files = [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]

print("Which file would you like to process?")
for i, p in enumerate(files):
    print(i, p.name)

num = int(input("Insert the number associated with the desired file: "))

while num < 0 or num >= len(files):
    num = int(input("Insert a number associated with a file: "))



pdf_path = files[num]
json_path = pdf_path.with_suffix(".json")
if not json_path.exists():

    print(f"Processing: {pdf_path}")


    patients = extract_data(pdf_path)

    with json_path.open("w", encoding="utf-8") as f:
        json.dump(patients, f, indent=2)
else:
    print(f"Using saved data from {json_path}")
    with json_path.open("r", encoding="utf-8") as f:
        patients = json.load(f)
    
buttons = configure()  



skipCount = 0
for patient in patients:
    currentPatient = patient
    if correctDetails(patient, buttons) == 1:
        skipCount += 1
        continue
    if processPatient(patient, buttons) == 1:
        skipCount += 1
        continue

