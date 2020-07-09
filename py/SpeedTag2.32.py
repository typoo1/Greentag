import os
import os.path
import sys
import re
import datetime
import selenium
import threading
import logging
import time
import subprocess
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import win32com
import win32com.client
import win32timezone
from win32com.client import Dispatch
import pyad
from pyad import adquery
from selenium.webdriver.common.keys import Keys
import itertools as it

"""
Program by Tye Alexander Gallagher
This program automates the morning greentagging proess for registers by reading in the user's emails and parsing the reports therein.
The program then uses that information to create a string to be pasted into the greentag portal and has the option of filling out incident reports to match those reports.
"""

Cities = ["TMP", "ORL", "SDO", "SAT", "LAG", "WIL"]
global city


"""
Default customer names
"""
defaultCust = []
##TMPCus = ["Samuel Goldstein", "Tyrone Robinson", "blank", "Terrence Lattimore", "Sidra Davis", "blank", "blank", "blank", "blank"]
##ORLCus = ["Samuel Goldstein", "Tyrone Robinson", "blank", "Terrence Lattimore", "Sidra Davis", "blank", "blank", "blank", "blank"]
##SDOCus = ["Samuel Goldstein", "Tyrone Robinson", "blank", "Terrence Lattimore", "Sidra Davis", "blank", "blank", "blank", "blank"]
##SATCus = ["John Blevins", "John Blevins", "John Blevins", "John Blevins", "John Blevins", "John Blevins", "John Blevins", "John Blevins", "John Blevins"]
##LAGCus = ["Samuel Goldstein", "Tyrone Robinson", "blank", "Terrence Lattimore", "Sidra Davis", "blank", "blank", "blank", "blank"]
##WILCus = ["Samuel Goldstein", "Tyrone Robinson", "blank", "Terrence Lattimore", "Sidra Davis", "blank", "blank", "blank", "blank"]

Cus = [] #Final customer list, to be loaded during Form Filler

q = pyad.adquery.ADQuery()

"""
Global variables to hold  form filler items
"""
global P1CusC #Park 1 Culinary Customer
global P2CusC #Park 2 Culinary Customer
global P1CusM #Park 1 Merch Customer
global P1CusC #Park 2 Merch Customer
global P3CusC #Park 3 Culinary Customer
global P3CusM #Park 3 Merch Customer
global AGroup #Assignment group
global P1ADOU #Holds the OU for park 1

"""
Grabs the path the program is currently in and adds a "\", used to find other necissary files
"""
path = os.path.dirname(os.path.abspath('Main.py'))
path = path + "\\"

"""
Arrays for holding various register objects
"""
offlineReg = []
probReg = []
offlineRemQueue = []
xstoreRegaR = []
xstoreRegbR = []
xstoreRegcR = []
xstoreRegaM = []
xstoreRegbM = []
xstoreRegcM = []
culinaryRega = []
culinaryRegb = []
culinaryRegc = []
MPRRega = []
MPRRegb = []
MPRRegc = []


"""
Sets the path for the chromedriver and configures the options so that chrome does not close once the page is loaded.
"""
driverP = (path + r"chromedriver.exe")
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.add_argument('log-level=3')
chrome_options.add_argument('disable-infobars')
chrome_options.add_argument('--disable-extensions')
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--no-sandbox')

offlineC = 0 #Count for offline registers

registers = [] #Array that holds all registers

f = open("testfile.txt", "w+") #test file for debugging purposes


"""
Attaches to client outlook to grab emails
"""
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts

global tFolder #Target search folder

"""
END OF VARIABLE SETUP
"""

"""
Sets the customer values according sure the config file.
"""
def setCus():
    global P1CusC
    global P2CusC
    global P3CusC
    global P1CusM
    global P2CusM
    global P3CusM
    global P1cusMPR
    global P2cusMPR
    global P3cusMPR
    P1CusC = Cus[0]
    P2CusC = Cus[1]
    P3CusC = Cus[2]
    P1CusM = Cus[3]
    P2CusM = Cus[4]
    P3CusM = Cus[5]
    P1cusMPR = Cus[6]
    P2cusMPR = Cus[7]
    P3cusMPR = Cus[8]

"""
Takes in a string as an argument and uses it to determine the city the user is configuring their client for.
This then assigns all Park code values, full names, assignments groups, and creates the Regex strings for find register names in the emails.
"""
def setPark(City):
    global Park1
    global Park2
    global Park3
    global Park1c
    global Park1m
    global Park2c
    global Park2m
    global Park3c
    global Park3m
    global Park1MPR
    global Park2MPR
    global Park3MPR
    global Park1QQ
    global Park2QQ
    global Park3QQ
    global Park1Full
    global Park2Full
    global Park3Full
    global AGroup
    global city
    global P1ADOU
    city = City
    
    
    print("Setting city to " + city)
    if(city == "TMP"): #Tampa
            Park1 = "BGT"
            Park2 = "AIT"
            Park3 = "DNE"
            Park1Full = "Busch Gardens Tampa"
            Park2Full = "Adventure Island Tampa"
            Park3Full = ""
            AGroup = "BGT-IT"
            P1ADOU = "_BGT"
                
    elif(city == "ORL"): #Orlando
            Park1 = "SWF"
            Park2 = "APO"
            Park3 = "DCO"
            Park1Full = "SeaWorld Florida"
            Park2Full = "Aquatica Park Florida"
            Park3Full = "Discovery Cove"
            AGroup = "SWF-IT"
            P1ADOU = "_SWF"
            
    elif(city == "SDO"): #San Diego
            Park1 = "SWC"
            Park2 = "APC"
            Park3 = "DNE"
            Park1Full = "SeaWorld California"
            Park2Full = "Aquatica Park California"
            Park3Full = ""
            AGroup = "SWC-IT"
            P1ADOU = "_SWC"
            
    elif(city == "SAT"): #San Antonio
            Park1 = "SWT"
            Park2 = "APT"
            Park3 = "DNE"
            Park1Full = "SeaWorld Texas"
            Park2Full = "Aquatica Park Texas"
            Park3Full = ""
            AGroup = "SWT-IT"
            P1ADOU = "_SWT"
            
    elif(city == "LAG"): #Langhorn
            Park1 = "SPL"
            Park2 = "DNE"
            Park3 = "DNE"
            Park1Full = "Sesame Place Langhorn"
            Park2Full = ""
            Park3Full = ""
            AGroup = "SPL-IT"
            P1ADOU = "_SPL"

    elif(city == "WIL"): #Williamsburg
            Park1 = "BGW"
            Park2 = "WCW"
            Park3 = "DNE"
            Park1Full = "Busch Gardens Williamsburg"
            Park2Full = "Water Country USA"
            Park3Full = ""
            AGroup = "BGW-IT"
            P1ADOU = "_BGW"

    Park1c = "^" + Park1 + ".*CP[0-9]*"
    Park1m = "^" + Park1 + ".*POS[0-9]*"
    Park2c = "^" + Park2 + ".*CP[0-9]*"
    Park2m = "^" + Park2 + ".*POS[0-9]*"
    Park3c = "^" + Park3 + ".*CP[0-9]*"
    Park3m = "^" + Park3 + ".*POS[0-9]*"
    Park1MPR = "^" + Park1 + ".*MPR[0-9]*"
    Park2MPR = "^" + Park2 + ".*MPR[0-9]*"
    Park3MPR = "^" + Park3 + ".*MPR[0-9]*"
    Park1QQ = "^" + Park1 + ".*QQ.*[0-9]*"
    Park2QQ = "^" + Park2 + ".*QQ.*[0-9]*"
    Park3QQ = "^" + Park3 + ".*QQ.*[0-9]*"


"""

"""
def printCulGreA(reg):
    if reg:
        s = Park2 + "RCP"
        for regX in reg:
            s = s + regX.name[-3:] + ", "
        s = s[:-2]
        s = s + " offline."
        return s
    
"""

"""
def printCulGreB(reg):
    if reg:
        s = Park1 + "RCP"
        for regX in reg:
            s = s + regX.name[-3:] + ", "
        s = s[:-2]
        s = s + " offline."
        return s

"""

"""
def printCulGreC(reg):
    if reg:
        s = Park3 + "RCP"
        for regX in reg:
            s = s + regX.name[-3:] + ", "
        s = s[:-2]
        s = s + " offline."
        return s

"""

"""
def xStoreMB(reg):
    if reg:
        s = ""
        for reg in reg:
            s = s + reg.name[-3:] + ", "
        s = s[:-2]
        return s
"""
Takes a list of registers in and creates a list of all registers
"""
def xStoreRB(reg):
    s = ""
    for reg in reg:
        s = s + reg.name[-3:] + ", "
    s = s[:-2]
    return s


"""

"""
def printXStoreA(regR, regM):
    if reg:
        s = ""
        if (regR and regM):
            s = Park2 + "RMPOS" + xStoreRB(regR) + " offline; " + Park2 + "MMPOS" + xStoreMB(regM)
        elif (regR):
            s = Park2 + "RMPOS" + xStoreRB(regR) + " offline."
        elif (regM):
            s = Park2 + "MMPOS" + xStoreMB(regM) + " offline."
        return s

"""

"""
def printXStoreB(regR, regM):
    s= ""
    if(regR and regM):
        s = Park1 + "RMPOS" + xStoreRB(regR) + " offline; " + Park1 + "MMPOS" + xStoreMB(regM) + " offline."
    elif(regR):
        s = Park1 + "RMPOS" + xStoreRB(regR) + " offline."
    elif(regM):
        s = Park1 + "MMPOS" + xStoreMB(regM) + " offline."
    return s

def printXStoreC(regR, regM):
    if reg:
        s = ""
        if (regR and regM):
            s = Park3 + "RMPOS" + xStoreRB(regR) + " offline; " + Park3 + "MMPOS" + xStoreMB(regM)
        elif (regR):
            s = Park3 + "RMPOS" + xStoreRB(regR) + " offline."
        elif (regM):
            s = Park3 + "MMPOS" + xStoreMB(regM) + " offline."
        return s

def printMPRGre(reg):
    s = ""
    for reg in reg:
        s = s + reg.name + ", "
    s = s[:-2]
    s = s + " offline."
    return s


"""
Determines wich print statements to run and prints the results for the user. This result is used to create the greentag statement to copy-paste into the greentag portal.
"""
def printReg():
    if(len(offlineReg) == 0):
       print("No register offline! Congrats")
       return()
    if(culinaryRegb):
        print(printCulGreB(culinaryRegb))
    if(culinaryRega):
        print(printCulGreA(culinaryRega))
    if(culinaryRegc):
        print(printCulGreC(culinaryRegc))
    if(xstoreRegbM or xstoreRegbR):
        print(printXStoreA(xstoreRegbR, xstoreRegbM))
    if(xstoreRegaM or xstoreRegaR):
        print(printXStoreB(xstoreRegaR, xstoreRegaM))
    if(xstoreRegcM or xstoreRegcR):
        print(printXStoreC(xstoreRegcR, xstoreRegcM))
    if(MPRRegb):
        print(printMPRGre(MPRRegb))
    if(MPRRega):
        print(printMPRGre(MPRRega))
    if(MPRRegc):
        print(printMPRGre(MPRRegc))
"""
Manages the entire process of opening the browser, navigating to the new incident page, and creating the new incidents
"""
def fillForms():
    threads = []
    if(len(offlineReg) - len(probReg) <= 10):
        for reg in it.chain(offlineReg, probReg):
            x = threading.Thread(target=Forms, args=([reg]))
            x.start()
            threads.append(x)
            time.sleep(1)
        for thread in threads:
            thread.join()
            print(str(thread.name) + " done")
    else:
        for reg in probReg:
            offlineReg.append(reg)
        input("the program can be somewhat unstable when opening many chrome windows at once. \nPress enter to open the first 10 chrome windows")
        i = 0
        while i < len(offlineReg):
            x = threading.Thread(target=Forms, args=([offlineReg[i]]))
            x.start()
            threads.append(x)
            time.sleep(1)

            if((i+1) % 10 == 0):
                for thread in threads:
                    thread.join()
                    print(str(thread.name) + " done")
                    threads.remove(thread)
                input("Please submit and close all open chrome windows, then press enter to continue")
                
            i += 1
        
        

def Forms(reg):
    driver = webdriver.Chrome(driverP, options=chrome_options) #Pass chrome driver the options previous configured
    #capabilities = DesiredCapabilities.Chrome.copy()
    #print(driver.capabilities['version'])
    driver.set_page_load_timeout(300) #Set maximum time that webpage is left to load
    driver.implicitly_wait(60)
    driver.get("https://sea.service-now.com/home.do") #It is necissary to navigate to the Service now homepage to prevent a series of redirects which breaks the remainders of the script
    driver.get("https://sea.service-now.com/incident.do") #Navigate to the New incident form.
    driver.find_element_by_tag_name('body').send_keys(Keys.ESCAPE)
    time.sleep(1)
    CID = driver.find_element_by_name("sys_display.incident.caller_id")  # caller
    REGNAME = driver.find_element_by_name("sys_display.incident.cmdb_ci")  # register name
    CONTYPE = driver.find_element_by_name("incident.contact_type")  # Contact type
    CONTYPE.send_keys("Email") #Static, does not change
    SHORTDES = driver.find_element_by_name("incident.short_description")  # Short description
    GROUP = driver.find_element_by_name("sys_display.incident.assignment_group") #Assignment group
    if(reg.park == Park3Full): #Check if park3
        if(re.search(Park3c, reg.name)):
            CID.send_keys(P3CusC)
        elif(re.search(Park3m, reg.name)):
            CID.send_keys(P3CusM)
        elif(re.search(Park3MPR, reg.name) or re.search(Park3QQ, reg.name)):
            CID.send_keys(P3cusMPR)
        else:
            CID.send_keys("!!!!ERROR!!!!")
    if(reg.park == Park2Full): #Check if park2
        if(re.search(Park2c, reg.name)):
            CID.send_keys(P2CusC)
        elif(re.search(Park2m, reg.name)):
            CID.send_keys(P2CusM)
        elif(re.search(Park2MPR, reg.name) or re.search(Park2QQ, reg.name)):
            CID.send_keys(P2cusMPR)
        else:
            CID.send_keys("!!!!ERROR!!!!")
    elif (reg.park == Park1Full): #check if park1
        if (re.search(Park1c, reg.name)):
            CID.send_keys(P1CusC)
        elif (re.search(Park1m, reg.name)):
            CID.send_keys(P1CusM)
        elif(re.search(Park1MPR, reg.name) or re.search(Park1QQ, reg.name)):
            CID.send_keys(P1cusMPR)
        else:
            CID.send_keys("!!!!ERROR!!!!")
    #CID.send_keys("\ue003")  # down arrow ##Depricated, caused issues with omitting the final character initally and no longer needed
    CID.send_keys("\ue004")  # tab
    REGNAME.send_keys(reg.name)
    REGNAME.send_keys("\ue004")  # tab
    if(reg.status.lower() == "offline"):
        SHORTDES.send_keys(reg.name + " " + reg.loc + " offline on morning report.")
    if(reg.status == "HDD problem"):
        SHORTDES.send_keys(reg.name + " " + reg.loc + " is low on HDD space.")
    if(reg.status == "Repl problem"):
        SHORTDES.send_keys(reg.name + " " + reg.loc + " is experiencing Replication error.")
    if(reg.status == "Close Failure"):
        SHORTDES.send_keys(reg.name + " " + reg.loc + " experienced a store close failure.")
    SHORTDES.send_keys("\ue004")  # tab
    REGNAME.send_keys(Keys.BACKSPACE)
    time.sleep(1)
    REGNAME.send_keys(reg.name[-1:])
    time.sleep(1)
##    REGNAME.send_keys(Keys.DOWN)
##    REGNAME.send_keys("\ue004")  # tab
    GROUP.send_keys((Keys.CONTROL + "a"))
    time.sleep(1)
    GROUP.send_keys(Park1 + "-IT") #should be filling automatically based on CI
    GROUP.send_keys("\ue004")
    #driver.implicitly_wait(1) #Wait statement, no longer needed.
"""
Defines the structure of the Register object used to store register information
"""
class Register:
    
    """
    Default Constructor
    """
    def __init__(self):
        self.name = ""
        self.park = ""
        self.status = ""
        self.loc = ""
        self.HDD = 100.0
        
##    """
##    Constructor with two arguments for the name and reported status of the register
##    """
##    def __init__(self, name, status):
##        self.name = name
##        x = ""
##        if(re.search("^" + Park2, name)):
##            self.park = Park2Full
##        else:
##            if(re.search("^" + Park1, name)):
##                self.park = Park1Full
##            else:
##                self.park = "ERROR"
##        if(status == "off"):
##                status = "offline"
##        if(status == "online" or status == "offline" or status == "off"):
##            self.status = status
##            self.loc= ""
##            if(status == "offline"):
##                self.setLoc()
##        else:
##                self.status = "ERROR"
##                print("REGISTER STATUS NOT FOUND, PLEASE REVIEW " + self.name)
##        self.HDD = 100.0

    """
    Constructor with added HDD argument. This is specific to the Culinary reports for detecting HDD faults.
    """
    def __init__(self, name, status, HDDVal):
        self.name = name
        x = ""
        if(re.search("^" + Park3, name)):
            self.park = Park3Full
        elif(re.search("^" + Park2, name)):
            self.park = Park2Full      
        elif(re.search("^" + Park1, name)):
            self.park = Park1Full
        else:
            self.park = "ERROR"
        if(status == "off"):
                status = "offline"
        if(status == "online" or status == "offline" or status == "HDD problem" or status == "Repl problem" or  status == "Close Failure"):
            self.status = status
            self.loc= ""
            if(status != "online"):
                self.setLoc()
        else:
                self.status = "ERROR"
                print("REGISTER STATUS NOT FOUND, PLEASE REVIEW " + self.name)
        self.HDD = HDDVal
                
    """
    Method that defines how to print information stored in the Register object
    """
    def printReg(self, x):
        if(self.status == "offline"):
            print("Register " + self.name + " at " + self.park + ": " + self.loc + " is " + str(self.status))
            print("Register " + self.name + " at " + self.park + ": " + self.loc + " is " + str(self.status), file = x)
        if(self.status == "HDD problem"):
            print("Register " + self.name + " at " + self.park + ": " + self.loc + " is low on HDD space.")
            print("Register " + self.name + " at " + self.park + ": " + self.loc + " is low on HDD space.", file = x)
        if(self.status == "Repl problem"):
            print("Register " + self.name + " at " + self.park + ": " + self.loc + " is experiencing a replication problem.")
            print("Register " + self.name + " at " + self.park + ": " + self.loc + " is experiencing a replication problem.", file = x)
        if(self.status == "Close Failure"):
            print("Register " + self.name + " at " + self.park + ": " + self.loc + " experienced a store close failure.")
            print("Register " + self.name + " at " + self.park + ": " + self.loc + " experienced a store close failure.", file = x)
            

    def setLoc(self):
        if(re.search("MMPOS", self.name) or re.search("RMPOS", self.name)):
            q.execute_query(
                attributes = ["CN", "description"],
                where_clause = "CN = '" + self.name + "'",
                base_dn = "OU=Xstore,OU=_Special Needs - Advertise & Install - No Reboot,OU=Computers,OU=" + P1ADOU + ",DC=nam,DC=int,DC=local")
        elif(re.search("CP", self.name)):
            q.execute_query(
                attributes = ["CN", "description"],
                where_clause = "CN = '" + self.name + "'",
                base_dn = "OU=Culinary_POS,OU=_Special Needs - Advertise & Install - No Reboot,OU=Computers,OU=" + P1ADOU + ",DC=nam,DC=int,DC=local")
        elif(re.search("MPR", self.name) or re.search("QQ", self.name)):
            q.execute_query(
                attributes = ["CN", "description"],
                where_clause = "CN = '" + self.name + "'",
                base_dn = "OU=MPR,OU=_Special Needs - Advertise & Install - No Reboot,OU=Computers,OU=" + P1ADOU + ",DC=nam,DC=int,DC=local")
        for row in q.get_results():
            self.loc = str(row["description"])[2:-3]

"""
Responsible for parsing emails
Takes a folder item as input to scan through
This method grabs all the messages from a folder without reviewing their contents, sorts them by their SentOn value, and iterates through them until a SentOn value doesn't match the current system date before breaking.
As the method iterates, it checks the sender to ensure that it is from a known source of greentag emails to avoid processing info from irrelevent emails.
"""
def emailleri_al(folder):
    messages = folder.Items
    a=len(messages)
    print("Parsing the following emails...\n")
    messages = sorted(messages, key=lambda messages: messages.SentOn)
    if a>0:
        for message2 in reversed(messages):
            try:
                sender = message2.SenderEmailAddress.lower()
                sdate = str(message2.SentOn)
                if sender != "":
                    if sender == "seap2018@seaworld.com" or "xstorereport@seaworld.com" or "swt.ithelpdesk@SeaWorld.com": #TODO find correct email addresses"
                        if sdate[:-15] == str(tarDate):
                            print(message2.Subject)
                            print("***********************", file =f)
                            print(message2.Subject, file=f)
                            print("***********************", file=f)
                            print(str(message2.SentOn)[:-15], file=f)
                            #manipulate the string
                            output = message2.Body
                            output = ' '.join(output.split())
                            print(output, file=f)
                            #Create register objects
                            strings = []
                            strings = output.split()
                            #print(strings)
                            #print(len(strings))
                            i = 0
                            
                            while i < len(strings):
                                if(re.search(Park2c, strings[i]) or re.search(Park2m, strings[i]) or re.search(Park2MPR, strings[i]) or re.search(Park2QQ, strings[i])
                                        or re.search(Park1c, strings[i]) or re.search(Park1MPR, strings[i]) or re.search(Park1m, strings[i]) or re.search(Park1QQ, strings[i])
                                        or re.search(Park3c, strings[i]) or re.search(Park3MPR, strings[i]) or re.search(Park3m, strings[i]) or re.search(Park3QQ, strings[i])):
                                    z = i
                                    statSet = False
                                    if(re.search(Park2QQ, strings[i]) or re.search(Park1QQ, strings[i]) or re.search(Park1MPR, strings[i]) or re.search(Park2MPR, strings[i] or re.search(Park3QQ, strings[i]) or re.search(Park3MPR, strings[i]))): #If MPR
                                        while z < len(strings) and z < i + 20:
                                            if(strings[z].lower() == "online" or strings[z].lower() == "off"):
                                                registers.append(Register(strings[i], strings[z].lower(), 101))
                                                break
                                            z+=1
                                            
                                    elif(re.search(Park2c, strings[i]) or re.search(Park1c, strings[i]) or re.search(Park3c, strings[i])): #If Culinary
                                        while z < len(strings) and z < i + 20:
                                            if(strings[z].lower() == "online" or strings[z].lower() == "offline" or strings[z].lower() == "off" or z + 1 == len(strings)):
                                                status = strings[z].lower()
                                                if(status == "offline"):
                                                    registers.append(Register(strings[i], status, 102))
                                                    break
                                            if(strings[z] == "%"): #Detect if register is from FreedomPay report dynamically
                                               if(float(strings[z-1]) < 20):
                                                    registers.append(Register(strings[i], "HDD problem", float(strings[z-1])))
                                                    break
                                               if(strings[z-5] == "6" or strings[z-6] == "6") or (strings[z-5].lower() == "missing" or strings[z-6].lower() == "missing") :
                                                    registers.append(Register(strings[i], "Repl problem", 103))
                                                    break
                                               registers.append(Register(strings[i], status, 25))
                                               break
                                            z+=1
                                    elif(re.search(Park2m, strings[i]) or re.search(Park1m, strings[i]) or re.search(Park3m, strings[i])): #If Xstore
                                        statSet = False
                                        while z < len(strings) and z < i + 20:
                                            if(strings[z].lower() == "online" or strings[z].lower() == "offline" or strings[z].lower() == "off" or z + 1 == len(strings)):
                                                if(strings[z].lower() == "offline"):
                                                    registers.append(Register(strings[i], strings[z].lower(), 104))
                                                    break
                                                if(statSet): #should only trigger on xstore registers
                                                    #if(strings[z-1].lower() == "failed"):
                                                    #    registers.append(Register(strings[i], "Close Failure", 105)) #According to feedback, store close failure read is not needed.
                                                    registers.append(Register(strings[i], status, 101))
                                                    break #This is used to prevent lines overlapping on Xstore reports specifically
                                                else:
                                                    status = strings[z].lower()
                                                    statSet = True
                                                #break
                                            z += 1
                                        # print("**************" + registers[i].name)
                                        # regNum = regNum + 1
                                i = i + 1
                        else:
                            print()
                            break
                    print(sender, file=f)
                try:
                    message2.Save
                    message2.Close(0)
                except:
                    pass
            except:
                print(sys.exc_info())
                pass
"""
Uses the win32com package to check the version of Chrome currently installed and throws an error if it's incompatible with the current chromedriver.exe installed
"""
def getVer():
    parser = Dispatch("Scripting.FileSystemObject")
    version = parser.GetFileVersion(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    #if not(version[:2] == "74" or version[:2] == "79" or version[:2] == "80"):
     #   raise ValueError("You are using an invalid version of Chrome, please update to version 74, 79, or 80")
      #  time.sleep(600)
    return(version)
                
"""
Checks to see if a config file exiests yet, and if not goes through several prompts with the user to create the appropriate config for them.
The parser for the config file assumes that the title of each config item is only a single word long.
"""
def getConfig():
    global tFolder
    global Cus
    if not (os.path.exists(path + "config.txt")):
        print("Config file not found, performing first time setup")
        config = open("config.txt", "w+")
        x = input("Please enter the name of the folder your register reports are put in ").lower()
        config.write("OutlookFolder: " + x + "\n")
        tFolder = x
        print("the target folder is: " + tFolder)
        x = input("Please enter your city code: TMP, ORL, SDO, SAT, LAG, WIL ").upper()
        config.write("CityCode: " + x + "\n")
        City = x
        setPark(x)
        i = 0
        print("Default customer values can be found at \\\\becabtpfil001\\IT\\SpeedTag\\Default.txt")
        print("Loading default customers.")
##        if(input("Would you like to use the default customers for your park? y/n").lower() == "y"):
        Cus = getDefaults(City)
##        else:
##            print("beginning custom customer setup...")
##            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park1Full + " Culinary: "))
##            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park2Full + " Culinary: "))
##            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park3Full + " Culinary (Leave blank if you only have 2 parks): "))
##            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park1Full + " Merchandise: "))
##            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park2Full + " Merchandise: "))
##            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park3Full + " Merchandise (Leave blank if you only have 2 parks): "))
##            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park1Full + " MPR (Leave blank if you don't recieve this report): "))
##            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park2Full + " MPR (Leave blank if you don't recieve this report): "))
##            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park3Full + " MPR (Leave blank if you don't recieve this report): "))
##
##            
##        for p in Cus:
##            s = ""
##            if i%3 == 0:
##                s = s + "P1"
##            elif i%3 == 1:
##                s = s + "P2"
##            elif i%3 == 2:
##                s = s + "P3"
##            if i <= 2:
##                s = s + "Cul"
##            elif i <= 5:
##                s = s + "Merch"
##            elif i <= 8:
##                s = s + "MPR"
##            
##            s = s + "Cus: " + p + "\n"
##            config.write(s)
##            i += 1
        setCus()
        print("Customers loaded")
        print("config created")
        config.close()
    else:
        try:
            config = open("config.txt", "r")
            y = []
            x = config.readlines()
            l = len(x)
            i = 0
            for item in x:
                y.append(parseItem(item))#load y with config variables
                i+=1
            tFolder = y[0] #assign the first item in the config as the target folder
            City = y[1]
            setPark(City) #asign the second item as the city
            Cus = getDefaults(City)
            setCus()
            print("Default customer values can be found at \\\\becabtpfil001\\IT\\SpeedTag\\Default.txt")
            print("config loaded")
        except:
            raise ValueError("Something went wrong with your config file, either edit it or delete it to have the script replace it.")
        config.close()
"""
Method used to parse the individual lines of the config file. It takes every part of a given line, removes the first item (the item's title) and returns the remaining items seperated by spaces as a single string
"""
def parseItem(item):
    y = item.split()
    l = len(y)
    i = 1
    result = ""
    while i < l:
        result = result + y[i] + " "
        i += 1
    result = result[:-1]
    return result

"""
Method used to parse the in the default customer values.
"""
def parseDefault(item):
    y = item.split()
    l = len(y)
    i = 1
    result = ""
    while i < l:
        result = result + y[i] + " "
        i += 1
    result = result[:-1]
    #print(result)
    return result

def getDefaults(City):
    default = open("\\\\becabtpfil001\\IT\\SpeedTag\\Default.txt", "r")
    y = []
    x = default.readlines()
    l = len(x)
    i = 0
    for item in x:
        y.append(parseDefault(item))#load y with config variables
        i+=1
    i = 3
    if City == "TMP":
        i = 3
    elif City == "ORL":
        i = 4
    elif City == "SDO":
        i = 5
    elif City == "SAT":
        i = 6
    elif City == "LAG":
        i = 7
    elif City == "WIL":
        i = 8
    defCust = y[i].split(", ") #Break imported string into list and assign to defCust
    return defCust
    print("Default loaded")
    default.close()
"""
Takes in an outlook account as an argument, then scans through all the account's email folders and their subfolders to find the one matching the target folder.
The method then calls the emailleri_al method on the target folder to scan for the desired emails.
"""
def getEmails(Account):
    for account in accounts:
        global inbox
        global tarDate
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        print("****Account Name*********************", file=f)
        print(account.DisplayName, file=f)
        print("From " + account.DisplayName +"\n")
        print("*************************************", file=f)
        folders = inbox.Folders
        tarDate = datetime.date.today()
        print("Processing mail from the " + tFolder + " folder from " + str(tarDate) + "\n")

        for folder in folders:
            print("*****Folder Name*****************", file=f)
            print(folder, file=f)
            print("*********************************", file=f)
            if(str(folder).lower() == tFolder.lower()):
                emailleri_al(folder)
            a = len(folder.folders)
            if a>0 :
                global z
                z = outlook.Folders(account.DeliveryStore.DisplayName).Folders(folder.name)
                x = z.Folders
                for y in x:
                    print("*****Folder Name********************", file=f)
                    print("..."+y.name, file=f)
                    print("************************************", file=f)
                    if(str(y).lower() == tFolder.lower()):
                        emailleri_al(y)

def Ping(reg, txt, queue):
    print(reg.name, file=txt)
    output = subprocess.Popen(["ping.exe", reg.name, "-n", "1"],stdout = subprocess.PIPE).communicate()[0]
    if "unreachable." in str(output):
        status = "offline"
    elif "could not find host" in str(output):
        status = "offline"
    elif "timed out" in str(output):
        status = "offline"
    else:
        status = "online"
    print(reg.name + " is pinging " + str(status))

    if status == "online":
        queue.append(reg)
##    if status == "0":
##        status = "online"
##        online = (next((x for x in offlineReg if x.name == name), None))
##        queue.append(online)
##    else:
##        status = "offline"
##    print(name + " pinging " + status)


def PrintOffline():
    global offlineReg
    offTxt = open("Offline.txt", "w+")
    print("Printing ping results...")
    threads = []
    for Reg in offlineReg:
        x = threading.Thread(target = Ping, args=(Reg, offTxt, offlineRemQueue))
        x.start()
        threads.append(x)
        time.sleep(.2)
    for thread in threads:
        thread.join()

##def RemOffline(tar):
##    for reg in tar:
##        print("removing " + reg.name + " as offline")
##        offlineReg.remove(reg)
##        offlineRemQueue.remove(reg)

def RemOffline(tar):
    i = 0
    while i < len(tar):
        print("removing " + reg.name + " as offline")
        offlineReg.remove(tar[i])
        offlineRemQueue.remove(tar[i])


"""
END OF DEFS
"""



print("Speedtag 2.3V")
print("A Program by Tye A. Gallagher")
print()
getVer()
getConfig()
getEmails(accounts)

regList = open("Reglist.txt", "w+")
for reg in registers:
    if(reg.status.lower() == "offline" or reg.status.upper() == "ERROR"):
        reg.printReg(regList)
        offlineC += 1
        offlineReg.append(reg)
    if(reg.status == "HDD problem" or reg.status == "Repl problem" or reg.status == "Close Failure"):
        reg.printReg(regList)
        offlineC += 1
        probReg.append(reg)
regList.close()
print()

    


for reg in offlineReg:
    if(reg.loc[:7].lower() == "offline"):
        offlineRemQueue.append(reg)
    elif(re.search("^" + Park1 + "RMPOS", reg.name)):
        xstoreRegbR.append(reg)
    elif(re.search("^" + Park1 + "MMPOS", reg.name)):
        xstoreRegbM.append(reg)
    elif(re.search("^" + Park2 + "RMPOS", reg.name)):
        xstoreRegaR.append(reg)
    elif(re.search("^" + Park2 + "MMPOS", reg.name)):
        xstoreRegaM.append(reg)
    elif(re.search("^" + Park3 + "RMPOS", reg.name)):
        xstoreRegcR.append(reg)
    elif(re.search("^" + Park3 + "MMPOS", reg.name)):
        xstoreRegcM.append(reg)
    elif(re.search("^" + Park1 + "RCP", reg.name) or (re.search("^" + Park1 + "CP", reg.name) or (re.search("^" + Park1 + "MCP", reg.name)))):
        culinaryRegb.append(reg)
    elif(re.search("^" + Park2 + "RCP", reg.name) or (re.search("^" + Park2 + "CP", reg.name) or (re.search("^" + Park2 + "MCP", reg.name)))):
        culinaryRega.append(reg)
    elif(re.search("^" + Park3 + "RCP", reg.name) or (re.search("^" + Park3 + "CP", reg.name) or (re.search("^" + Park3 + "MCP", reg.name)))):
        culinaryRegc.append(reg)
    elif(re.search("^" + Park1 + ".*MPR", reg.name) or re.search("^" + Park1 + ".*QQ", reg.name)):
       MPRRegb.append(reg)
    elif(re.search("^" + Park2 + ".*MPR", reg.name) or re.search("^" + Park2 + ".*QQ", reg.name)):
       MPRRega.append(reg)


RemOffline(offlineRemQueue)
#RemOffline(offlineRemQueue)

    
print()
print("There are " + str(len(registers)) + " registers reported on and " + str(len(offlineReg)) + " registers that need attention")

print()
print("Greentag statements:")
printReg()
#PrintOffline() #TODO make the print statements come out nicer

if(input("\nWould you like to ping the offline registers? Y/N \n").lower() == "y"):
      PrintOffline()
      if(input("\nRemove registers that pinged online? Y/N \n").lower() == "y"):
          RemOffline(offlineRemQueue)

f.close()
            
if(input("\nProceed with form filler? Y/N \n").lower() == "y"):
    try:
        driver = webdriver.Chrome(driverP, options=chrome_options) #Pass chrome driver the options previous configured
        driver.close()
    except:
        print("There was an error with your chrome driver, please verify your chrome driver version matches your chrome version. Chromedriver is available here: https://chromedriver.chromium.org/")
        chromeDriverVer = (str(sys.exc_info()[1])[88:90])
        print()
        print("Your current chrome driver is comaptible with version: " + chromeDriverVer + " of Google Chrome.")
        print("You are currently running version " + str(getVer())[:2] + " of Chrome.")
        time.sleep(60)
        print(sys.exc_info())
    fillForms()

"""
END OF PROGRAM
"""
