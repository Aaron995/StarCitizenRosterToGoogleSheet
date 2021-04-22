# Script to add, remove and update members in the sheet.
import gspread
import pyrsi
import time
from datetime import date

# Setting the account up and getting the sheet.
gc = gspread.service_account()
sheet = gc.open("INFINITE EMPIRE ROSTER")
sheet = sheet.get_worksheet(0)

# Settuping up varibles we need.
writes = 0
index = 0
infoList = []
startingRow = 2


# Updating status message
t = time.localtime()
current_time = time.strftime("%H:%M ", t)
date = str(date.today())
timezone = time.tzname[0]
sheet.update_acell("F1", "Last update start: " + date + " " + current_time + " " + timezone)
sheet.update_acell("H1", "Still updating!")


# Getting org members and members on the sheet and their info.
# Also making a username list
from rsi.org import OrgAPI
org = OrgAPI('INFE')
members = org.members  
membersInSheet = sheet.col_values(1)
membersInSheet.pop(0)
memberRanks = sheet.col_values(3)
memberRanks.pop(0)
memberAffilation = sheet.col_values(2)
memberAffilation.pop(0)
usernameList = []
for member in members:
   usernameList.append(member["handle"])



# Loop through members on the sheet, remove them from the lists and sheet
# If they are not in the downloaded org list.
for i, member in enumerate(membersInSheet):
    if member not in usernameList:       
        x = i + 2
        membersInSheet.pop(i)
        memberRanks.pop(i)
        memberAffilation.pop(i)
        sheet.delete_rows(x)
        writes = writes + 1
        if writes > 90:            
            time.sleep(105)
            writes = 0


# Checking member if ranks and affiliation are equal
for member in members:
    if member["handle"] in membersInSheet:        
        number = membersInSheet.index(member["handle"])
        row = number + 2
    else:
        continue
    if(str(memberAffilation[number]).lower() != str(member["affiliate"]).lower()):
        sheet.update_cell(row,2, member["affiliate"])
        writes = writes + 1
        time.sleep(100)
        if writes > 90:
            time.sleep(105)
            writes = 0
    if memberRanks[number] != member["rank"]:
        sheet.update_cell(row,3, member["rank"])
        writes = writes + 1
        if writes > 90:
            time.sleep(105)
            writes = 0

# Adding new members
for member in members:    
    # Check if the handle is already on the sheet, skip is already on it
    if member["handle"] in membersInSheet:
        continue

    keys = list(member.keys())    
    # Loop through the member dict and skip the values we don't want displayed.
    # Add the values we want to write in the speed to the list.
    while index < len(member):
        if keys[index] == "avatar" or keys[index] == "id" \
        or keys[index] == "last_online" or keys[index] == "roles" \
        or keys[index] == "name":
            index = index + 1
            continue
        infoList.append(member[keys[index]])
        index = index + 1

    # Creating the row and add the info in here. Keeping track of API calls
    # Pausing the script when close to the API limits.
    sheet.insert_row(infoList, startingRow)
    writes = writes + 5
    if writes > 90:
        time.sleep(105)
        writes = 0
    infoList.clear()
    index = 0

sheet.sort((1, 'asc'))
t = time.localtime()
current_time = time.strftime("%H:%M ", t)
sheet.update_acell("H1", "Update finished at: " + date + " " + current_time + " " + timezone)

