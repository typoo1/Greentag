Python 3.7.5
Setup:
Ensure that Chromedriver.exe is in the same folder as SpeedTag1_3.exe
Outlook:
Setup rules to send all register reports to one folder
The easiest method is to use the rules wizard and move all emails from the relevent email addresses to the correct folder.


User Instructions:
Launch SpeedTag1_3.exe as any other .exe, no admin permissions should be required.
A command prompt will open. Depending on if you have a config file set up yet or not, said file will either be loaded, or you'll be prompted to create the file through the script.
If an error occurs around the config file, it is recommended to delete the file and try to set it up again.
If the error persists, email Tye.Gallagher@buschgardens.com

Maintenece and Operation:
Speedtag1_3.exe connects to your outlook automatically as long as you are signed in through the Windows win32com package.
Upon opening the first time it will request a folder to search in for Greentag emails. (Capitalization is not important for any user inputs)
The program will only look at emails sent on the current sytem date.

During the first use, or if no config file is found, the script will ask a series of questions that will be used by the script.
This includes the city the park is located in, the folder the script will look in for the reports, and the names of customers the user would like incidents to be called in from.
It is possible to modify the config file manually. The first value (IE CityCode) should not be changed. If any errors occur following modification, it is recommended taht you delete the config and start with a fresh setup.

The script will go through the target folder and pull any emails from the current date.
These emails are then parsed into an array of strings which the script iterates through looking for registers and their status
It uses the following Regex identifiers to find and sort registers into their correct park and department:
    Park1c = "^" + Park1 + "...[0-9][0-9][0-9]"
    Park1m = "^" + Park1 + ".....[0-9][0-9][0-9]"
    Park2c = "^" + Park2 + "...[0-9][0-9][0-9]"
    Park2m = "^" + Park2 + ".....[0-9][0-9][0-9]"
where Park1 and Park2 refer to the 2 parks in association with the city given at setup (For Langhorn Park2 is an empty string)
The "^" symbol denotes that the string begins with the regex identifier that follows
"."s denote any character
[0-9] indicates any digit

The script then iterates through the next 10 nodes to look for either "Online" or "Offline", if neither is found "ERROR!" will be listed as the register's status.
It is important to note that extremely long location names may break this section of code
In this case, find the line reading "while z < len(strings) and z < i + 10:" and change 10. Lower numbers will yield higher performance, but may miss indicators blocked by names.
The results of the scan are used to create Register objects which hold the name, location, and status of each register.

The register objects are then sorted into arrays listing offline registers for each park and department.
Once the full email is parsed into these arrays, the script then prints all registers held in the offline registers array
The script then generates a greentag statement for offline registers to be pasted into the greentag online portal.

The final prompt will ask the user if they would like to begin filling forms.
if the user says yes, the program will then begin launching chrome windows which will navigate to the Service Now homepage and then to the new incident page
The script then iterates through each type of offline register and uses that information to fill in the incident tickets.
The script is hard coded with names to correspond with each park and department combination to fill in the callerID, and uses the Register datapoints to fill in the remaining information
At current the user is required to cross reference with AD to find the location actual and paste that into the short description on Service Now, Future iterations may link with AD to provide the location inside the report.
Users should check the fields filled in for any errors before submiting.

This script relies on Selenium to launch chrome from Chromedriver.exe. Chromedriver.exe is a lightweight version of Google chrome used for automation.
Chromedriver pulls from the currently installed version of Google Chrome and must be updated from the Selenium website if chrome is updated.
If the script fails when running the form filler, check the version of chrome by opening a chrome window, opening the triple dot menu > help > about chrome.
Document any updates to Chromedriver.exe in this Readme

Current Chromedriver.exe version: 74


"# greentag" 
