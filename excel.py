import xlwings as xw
import psutil

def validateExcelFile(excelFilePath):
    app = xw.App(visible=False)

    closeExcelFile(excelFilePath)

    try:
        wb = xw.Book(excelFilePath, ignore_read_only_recommended=True)
        summary = None
        filteredSheets = []
        for sheet in wb.sheets:
            if "summary" == sheet.name.lower():
                summary = sheet
            elif "history" in sheet.name.lower():
                continue
            else: 
                filteredSheets.append(sheet.name)
                
        if not summary or summary.range('A1').value != "Claimbot Summary":
            return []
        
    except FileNotFoundError:
        print(f"File not found.")
    
    app.quit()
    return filteredSheets

def getMembersByInsurance(excelFilePath, insurance):
    app = xw.App(visible=False)

    closeExcelFile(excelFilePath)

    try:
        wb = xw.Book(excelFilePath, ignore_read_only_recommended=True)
        for sheet in wb.sheets:
            if sheet.name == insurance:
                ws = sheet
                break
        print(ws.range('A1').value)
        
    except FileNotFoundError:
        print(f"File not found.")
    
    app.quit()

def closeExcelFile(excelFilePath):
    for proc in psutil.process_iter():
        try:
            if 'EXCEL.EXE' in proc.name():
                for item in proc.open_files():
                    if excelFilePath.lower() in item.path.lower():
                        proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

