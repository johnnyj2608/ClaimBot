from datetime import datetime, timedelta

def getDatesFromWeekdays(startDate, endDate, weekdays, authStart, authEnd, vacaStart, vacaEnd):
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
        if curDate.weekday() + 1 in weekdays:
            dates.append(curDate)
        curDate += delta
    return intersectVacations(dates, vacaStart, vacaEnd)

def intersectVacations(dates, start, end):
    if not start or not end:
        return dates
    
    serviceDates = []
    for date in dates:
        if start <= date <= end:
            continue
        serviceDates.append(date)

    return serviceDates

def stopProcess(stopFlag):
    if stopFlag.value:
        return True
    return False