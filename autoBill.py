from pathlib import Path
from fileToText import extract_data
from billProcessor import processPatient
from configureButtons import configure, ButtonPress
from helperFunctions import load_json, save_json, get_window_name
import keyboard
import json
import string
import re
import pyautogui
import time
import winsound
import builtins

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

currentPatient = None
collectionMode = True

def printAllDetails():
    global currentPatient
    print(currentPatient)

keyboard.add_hotkey("alt+ctrl+p", printAllDetails, suppress=True)

def ensurePatientFound(patient, buttons):
    if "PatientName" not in patient:
        if "PatientDOB" in patient:
            print(patient["PatientDOB"])
            patient["PatientName"] = input(f"Please put in patient LASTNAME, Firstname with DOB {patient["PatientDOB"]}")
        else:
            print(patient, "Couldn't find any patient associated to this information")
            return 1
    
def processPatientQuery(patient, buttons):
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
        return False
    return True

def checkDOBandHospital(patient, buttons):
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
                return False
            patient["Hosp"] = hosp
    return True

def ensureCorrectDate(patient, buttons):
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

def correctAndSaveNumbers(patient, buttons, items_db, number, i=-1):
    number = number.translate(str.maketrans("", "", string.punctuation))
    number = re.sub(r"\D", "", number)
    if number in items_db:
        if items_db[number] != number:
            inp = input(f"Is {items_db[number]} supposed to be instead of {number}? Enter y or enter the true number")
            if inp == "y":
                if i == -1:
                    patient["ProcNumbers"] = items_db[number]
                else:
                    patient["ProcNumbers"][i] = items_db[number]
            else:
                if i == -1:
                    patient["ProcNumbers"] = inp
                else:
                    patient["ProcNumbers"][i] = inp
        else:
            if i == -1:
                patient["ProcNumbers"] = number
            else:
                patient["ProcNumbers"][i] = number
    else:
        inp = input(f"Enter 'y' if {number} is correct, if not enter true number")
        if inp == "y":
            items_db[number] = number
            if i == -1:
                patient["ProcNumbers"] = number
            else:
                patient["ProcNumbers"][i] = number
    
        else:
            if i == -1:
                patient["ProcNumbers"] = inp
            else:
                patient["ProcNumbers"][i] = inp
            items_db[number] = inp

def ensureCorrectItemNumbers(patient, buttons):
    items_json = Path("itemNums.json")
    items_db = load_json(items_json)
    if "ProcNumbers" not in patient:
        patient["ProcNumbers"] = list()
        inp = input("Enter in numbers until all in, then enter y")
        while inp != "y":
            patient["ProcNumbers"].append(inp)
            if inp not in items_db:
                items_db[inp] = inp
    elif isinstance(patient["ProcNumbers"], list):
        for i, number in enumerate(patient["ProcNumbers"]):
            correctAndSaveNumbers(patient, buttons, items_db, number, i)
    else:
        number = patient["ProcNumbers"]
        correctAndSaveNumbers(patient, buttons, items_db, number)
    save_json(items_json, items_db)

def ensureCorrectReferral(patient, buttons):
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
        else:
            inp = input(f"Is {patient["DerivedReferral"]} correct? y/n")
            if inp == "y":
                referral_db[patient["DerivedReferral"]] = "-1"
                patient["Referral"] = patient["DerivedReferral"]
            else:
                inp = input("Please paste in correct referral")
                patient["Referral"] = inp
    save_json(referral_json, referral_db)

def ensureCorrectTimes(patient, buttons):
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
    

def correctDetails(patient, buttons):
    time.sleep(0.5)
    if get_window_name() != "Windows PowerShell":
        buttons["Terminal"].clickButton()
    
    ensurePatientFound(patient, buttons)
    if not processPatientQuery(patient, buttons):
        return 1
    if not checkDOBandHospital(patient, buttons):
        return 1
    ensureCorrectDate(patient, buttons)
    ensureCorrectReferral(patient, buttons)
    ensureCorrectItemNumbers(patient, buttons)
    if "Times" in patient:
        ensureCorrectTimes(patient, buttons)
    
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

