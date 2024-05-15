from datetime import datetime, timedelta

def getDatesFromWeekdays(startDate, endDate, weekdays, authStart, authEnd):
    authStart = datetime.combine(authStart, datetime.min.time())
    authEnd = datetime.combine(authEnd, datetime.min.time())

    if authStart > startDate:
        startDate = authStart
    if authEnd < endDate:
        endDate = authEnd
    dates = []
    weekdays = set([int(x) for x in weekdays.split(".")])
    delta = timedelta(days=1)
    
    curDate = startDate
    while curDate <= endDate:
        if curDate.weekday() in weekdays:
            dates.append(curDate)
        curDate += delta
    return dates

def intersectVacations():
    pass