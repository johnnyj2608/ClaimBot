from datetime import datetime, timedelta

def getDatesFromWeekdays(startDate, endDate, weekdays):
    dates = []
    weekdays = set([int(x) for x in weekdays.split(".")])
    delta = timedelta(days=1)
    
    curDate = startDate
    while curDate <= endDate:
        if curDate.weekday() in weekdays:
            dates.append(curDate)
        curDate += delta
    return dates