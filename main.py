
# Have an excel sheet with tabs for each insurance (payer / templates)
# First tab will be a "summary" page
# includes: payer, billing & rendering provider, and facilities
# Count days. If not enough rows , create rows until satisfactory
# Auto fill days according to member's schedule
# Download claim to folder
# Have a GUI to select excel tab
# Option for "automatically submit" each claim

from datetime import datetime, timedelta

test = ["Test McGee", "ABC123123123", [0, 1, 2, 3, 4]]

def getDatesFromWeekdays(month, year, weekdays):
    numDays = (datetime(year, month % 12 + 1, 1) - timedelta(days=1)).day
    monthDays = []

    for day in range(1, numDays+1):
        date = datetime(year, month, day)
        if date.weekday() in weekdays:
            monthDays.append(date.strftime("%m/%d/%Y"))


    return monthDays

print(getDatesFromWeekdays(4,2024,[0, 1, 2, 3, 4]))