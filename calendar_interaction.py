# Import packages
from win32com.client import Dispatch
from tabulate import tabulate
import pandas as pd
import datetime
from datetime import timedelta
import pdb

# Define the format
OUTLOOK_FORMAT = '%m/%d/%Y %H:%M'
outlook = Dispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")

# Get the outlook calendar items
appointments = ns.GetDefaultFolder(9).Items

# Define the timeline variables
subtract_days = 20
plus_days = 3

# Define the timeline (using Python 3.3 - might be slightly different for 2.7)
begin = datetime.date.today() + datetime.timedelta(days = -subtract_days)
end = datetime.date.today() + datetime.timedelta(days = plus_days);
restriction = "[Start] >= '" + begin.strftime("%m/%d/%Y") + "' AND [End] <= '" +end.strftime("%m/%d/%Y") + "'"
restrictedItems = appointments.Restrict(restriction)

# Set the parameters for starting time and if you want to include recurrent meeting
restrictedItems.Sort("[Start]")
restrictedItems.IncludeRecurrences = "True"

# Iterate through restricted AppointmentItems and print them
calcTableHeader = ['Title', 'Start', "End", 'Duration(Minutes)'];
calcTableBody = [];

# Define the keywords in meeting titles you want to be included
keywords = ["Project xxx EN | Competitors", "DMs call"]

# Get the variables from the calendar
for restrictedItem in restrictedItems:
    row = []
    row.append(restrictedItem.Subject)
    # row.append(restrictedItem.Organizer)
    row.append(restrictedItem.Start.Format(OUTLOOK_FORMAT))
    row.append(restrictedItem.End.Format(OUTLOOK_FORMAT))
    row.append(restrictedItem.Duration)
    if (datetime.datetime.strptime(row[2][:10], "%m/%d/%Y").date() >= datetime.datetime.today().date()+ timedelta(-subtract_days)) and (datetime.datetime.strptime(row[2][:10], "%m/%d/%Y").date() < datetime.datetime.today().date() + timedelta(plus_days)) and any(keyword.lower() in row[0].lower() for keyword in keywords):
        calcTableBody.append(row)

# This code is for getting the meeting table
# print (tabulate(calcTableBody, headers=calcTableHeader))

# Convert it into a pandas table
data = pd.DataFrame (calcTableBody, columns = calcTableHeader)

# Output the table into an excel file
def data_manipulating (x):
    x[["Project", "Category", "Name", "Position and company"]] = x["Title"].str.split("|", expand=True)
    x[["Start date", "Start time"]] = x["Start"].str.split(" ", expand=True)
    y = x[["Project", "Category", "Name", "Position and company", "Start date", "Start time", "Duration(Minutes)"]]
    y.to_excel("expert_list.xlsx", index=False)

data_manipulating(data)
