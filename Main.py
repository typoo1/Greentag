import os
import sys
import re
import datetime
import selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import win32com
import win32com.client
import win32timezone


AITc = "^AIT...[0-9][0-9][0-9]"
AITm = "^AIT.....[0-9][0-9][0-9]"
BGTc = "^BGT...[0-9][0-9][0-9]"
BGTm = "^BGT.....[0-9][0-9][0-9]"

offlineReg = []
xstoreRegaR = []
xstoreRegbR = []
xstoreRegbM = []
xstoreRegaM = []
culinaryRega = []
culinaryRegb = []

driverP = (r"C:\data\Main\chromedriver.exe") #TODO find path
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)


offlineC = 0

registers = []

f = open("testfile.txt", "w+")

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts


def printCulGreB(reg):
    if reg:
        s = "BGTRCP"
        for regX in reg:
            s = s + regX.name[-3:] + ", "
        s = s[:-2]
        s = s + " offline."
        return s

def printCulGreA(reg):
    if reg:
        s = "AITRCP"
        for regX in reg:
            s = s + regX.name[-3:] + ", "
        s = s[:-2]
        s = s + " offline."
        return s

def xStoreMB(reg):
    if reg:
        s = ""
        for reg in reg:
            s = s + reg.name[-3:] + ", "
        s = s[:-2]
        return s

def xStoreRB(reg):
    s = ""
    for reg in reg:
        s = s + reg.name[-3:] + ", "
    s = s[:-2]
    return s


def printXStoreA(regR, regM):
    if reg:
        s = ""
        if (regR and regM):
            s = "AITRMPOS" + xStoreRB(regR) + " offline; AITMMPOS" + xStoreMB(regM)
        elif (regR):
            s = "AITRMPOS" + xStoreRB(regR) + " offline."
        elif (regM):
            s = "AITMMPOS" + xStoreMB(regM) + " offline."
        return s

def printXStoreB(regR, regM):
    s= ""
    if(regR and regM):
        s = "BGTRMPOS" + xStoreRB(regR) + " offline; BGTMMPOS" + xStoreMB(regM) + " offline."
    elif(regR):
        s = "BGTRMPOS" + xStoreRB(regR) + " offline."
    elif(regM):
        s = "BGTMMPOS" + xStoreMB(regM) + " offline."
    return s

def printReg():
    if(culinaryRegb):
        print(printCulGreB(culinaryRegb))
    if(culinaryRega):
        print(printCulGreA(culinaryRega))
    if(xstoreRegbM or xstoreRegbR):
        print(printXStoreB(xstoreRegbR, xstoreRegbM))
    if (xstoreRegaM or xstoreRegaR):
        print(printXStoreB(xstoreRegaR, xstoreRegaM))

def fillForms():
    for reg in offlineReg:
        driver = webdriver.Chrome(driverP, options=chrome_options)
        driver.set_page_load_timeout(15)
        driver.get("https://sea.service-now.com/nav_to.do?uri=%2Fhome.do")
        driver.get("https://sea.service-now.com/incident.do?sys_id=-1&sysparm_query=active=true&sysparm_stack=incident_list.do?sysparm_query=active=true")
        #driver.get("file:///C:/Users/TyeG/Desktop/test.html") #TODO change to other variable in prod
        CID = driver.find_element_by_name("sys_display.incident.caller_id")  # caller
        REGNAME = driver.find_element_by_name("sys_display.incident.cmdb_ci")  # register name
        CONTYPE = driver.find_element_by_name("incident.contact_type")  # Contact type
        CONTYPE.send_keys("Email") #Static, does not change
        SHORTDES = driver.find_element_by_name("incident.short_description")  # Short description
        GROUP = driver.find_element_by_name("sys_display.incident.assignment_group")
        if(reg.park == "Adventure Island"):
            if(re.search(AITc, reg.name)):
                CID.send_keys("Tyrone Robinson")
            elif(re.search(AITm, reg.name)):
                CID.send_keys("Sidra Davis")
            else:
                CID.send_keys("!!!!ERROR!!!!")
        elif (reg.park == "Busch Gardens"):
            if (re.search(BGTc, reg.name)):
                CID.send_keys("Samuel Goldstein")
            elif (re.search(BGTm, reg.name)):
                CID.send_keys("Terrence Lattimore")
            else:
                CID.send_keys("!!!!ERROR!!!!")
        #CID.send_keys("\ue003")  # down arrow
        CID.send_keys("\ue004")  # tab
        REGNAME.send_keys(reg.name)
        #REGNAME.send_keys("\ue003")  # down arrow
        REGNAME.send_keys("\ue004")  # tab
        SHORTDES.send_keys(reg.name + " [check name on AD and insert here] offline on morning report.")
        #SHORTDES.send_keys("\ue003")  # down arrow
        SHORTDES.send_keys("\ue004")  # tab
        #GROUP.send_keys(r"BGT-IT")
        #GROUP.send_keys("\ue004")
        #driver.implicitly_wait(1)


class Register:

    def __init__(self):
        self.name = ""
        self.park = ""
        self.status = ""

    def __init__(self, name, status):
        self.name = name
        if(re.search("^AIT", name)):
            self.park = "Adventure Island"
        else:
            if(re.search("^BGT", name)):
                self.park = "Busch Gardens"
            else:
                self.park = "ERROR"
        if(status == "online" or status == "offline"):
            self.status = status
        else:
                self.status = "ERROR"
                print("REGISTER STATUS NOT FOUND, PLEASE REVIEW " + self.name)

    def printReg(self):
        print("Register " + self.name + " at " + self.park + " is " + str(self.status))

def emailleri_al(folder):
    messages = folder.Items
    a=len(messages)
    messages = sorted(messages, key=lambda messages: messages.SentOn)
    if a>0:
        for message2 in reversed(messages):
            try:
                sender = message2.SenderEmailAddress.lower()
                sdate = str(message2.SentOn)
                if sender != "":
                    if sender == "tye.gallagher@buschgardens.com" or "tyeg@outlook.com" or "seap2018@seaworld.com" or "xstorereport@seaworld.com": #TODO find correct email addresses"
                        if sdate[:-15] == str(tarDate):
                            print(message2.Subject)
                            print("***********************", file =f)
                            print(message2.Subject, file=f)
                            print("***********************", file=f)
                            print(str(message2.SentOn)[:-15])
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
                                
                                if(re.search(AITc, strings[i]) or re.search(AITm, strings[i]) or re.search(BGTc, strings[i])
                                        or re.search(BGTm, strings[i])):
                                    z = i
                                    while z < len(strings) and z < i + 10:
                                        if(strings[z].lower() == "online" or strings[z].lower() == "offline"):
                                            registers.append(Register(strings[i], strings[z].lower()))
                                            break
                                        z += 1
                                    # print("**************" + registers[i].name)
                                    # regNum = regNum + 1
                                i = i + 1
                        else:
                            break
                    print(sender, file=f)
            except:
                print(sys.exc_info())
                pass

                try:
                    message2.Save
                    message2.Close(0)
                except:
                    pass

for account in accounts:
    global inbox
    global tarDate
    inbox = outlook.Folders(account.DeliveryStore.DisplayName)
    print("****Account Name*********************", file=f)
    print(account.DisplayName, file=f)
    print(account.DisplayName)
    print("*************************************", file=f)
    tFolder = input("What folder should I look in? ")
    folders = inbox.Folders
    tarDate = datetime.date.today()
    print("processing mail from " + str(tarDate))

    for folder in folders:
        print("*****Folder Name*****************", file=f)
        print(folder, file=f)
        print("*********************************", file=f)
        if(str(folder).lower() == tFolder.lower()):
            print(str(folder) + " " + tFolder)
            emailleri_al(folder)

        # a = len(folder.folders)
        # if a>0 :
        #     global z
        #     z = outlook.Folders(account.DeliveryStore.DisplayName).Folders(folder.name)
        #     x = z.Folders
        #     for y in x:
        #         emailleri_al(y)
        #         print("*****Folder Name********************", file=f)
        #         print("..."+y.name, file=f)
        #         print("************************************", file=f)


for reg in registers:
    if(reg.status.lower() == "offline" or reg.status.upper() == "ERROR"):
        reg.printReg()
        offlineC += 1
        offlineReg.append(reg)
print("There are " + str(len(registers)) + " registers reported on and " + str(offlineC) + " offline registers")

for reg in offlineReg:
    if(re.search("^BGTRMPOS", reg.name)):
        xstoreRegbR.append(reg)
    if(re.search("^BGTMMPOS", reg.name)):
        xstoreRegbM.append(reg)
    if(re.search("^AITRMPOS", reg.name)):
        xstoreRega.append(reg)
    if(re.search("^AITMMPOS", reg.name)):
        xstoreRegaM.append(reg)
    if (re.search("^BGTRCP", reg.name) or (re.search("^BGTMCP", reg.name))):
        culinaryRegb.append(reg)
    if (re.search("^AITRCP", reg.name) or (re.search("^AITMCP", reg.name))):
        culinaryRega.append(reg)

printReg()
if(input("Proceed with form filler? Y/N").lower() == "y"):
    fillForms()




