# Script to just set up members on the sheet.
import gspread
import pyrsi
import time

# Getting the account ready and getting the sheet
gc = gspread.service_account()
sheet = gc.open("INFINITE EMPIRE ROSTER")
sheet = sheet.get_worksheet(0)

# Initiating the varibles I need
startingRow = 2
writes = 0
index = 0
infoList = []

# Get the current org members and the members who are already on the sheet
from rsi.org import OrgAPI
org = OrgAPI('INFE')
members = org.members  
membersInSheet = sheet.col_values(1)


for member in members:    
    # Check if the handle is already on the sheet, skip is already on it
    if member["handle"] in membersInSheet:
        print("Player already in sheet, skipping")
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
# Sorting A->Z at the end based on Username
sheet.sort((1, 'asc'))




    
    

