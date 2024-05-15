import xlwings as xw
import psutil
import pandas as pd

def validateExcelFile(excelFilePath):
    app = xw.App(visible=False)

    closeExcelFile(excelFilePath)

    filteredSheets = []
    summaryValues = {}

    try:
        wb = xw.Book(excelFilePath, ignore_read_only_recommended=True)
        summary = None
        
        for sheet in wb.sheets:
            if "summary" == sheet.name.lower():
                summary = sheet
            elif "history" in sheet.name.lower():
                continue
            else: 
                filteredSheets.append(sheet.name)

        if not summary or summary.range('A1').value != "Claimbot Summary":
            return [], {}
        
        summaryValues = {
                "billingProvider": summary.range('C2').value,
                "renderingProvider": summary.range('C3').value,
                "facilities": summary.range('C4').value,
                "username": summary.range('C6').value,
                "password": summary.range('C7').value
            }
        
    except FileNotFoundError:
        print(f"File not found.")
    
    app.quit()
    return filteredSheets, summaryValues

def getMembersByInsurance(excelFilePath, insurance):
    app = xw.App(visible=False)

    closeExcelFile(excelFilePath)

    try:
        wb = xw.Book(excelFilePath, ignore_read_only_recommended=True)
        ws = None
        for sheet in wb.sheets:
            if sheet.name == insurance:
                ws = sheet
                break

        if not ws:
            return []
        
        members = []
        dataRange = ws.range('A1:H1').expand('down').value
        df = pd.DataFrame(dataRange[1:], columns=dataRange[0])
        df['Birth Date'] = pd.to_datetime(df['Birth Date']).dt.date
        df['Auth Start'] = pd.to_datetime(df['Auth Start']).dt.date
        df['Auth End'] = pd.to_datetime(df['Auth End']).dt.date
        for index, row in df.iterrows():
            members.append(list(row))

    except FileNotFoundError:
        print(f"File not found.")
    
    app.quit()
    return members

def recordClaims(insurance, start, end, claims):
    pass

def closeExcelFile(excelFilePath):
    for proc in psutil.process_iter():
        try:
            if 'EXCEL.EXE' in proc.name():
                for item in proc.open_files():
                    if excelFilePath.lower() in item.path.lower():
                        proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass