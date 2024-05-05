import xlwings as xw
import psutil

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

def validateExcelFile(excelFilePath):
    app = xw.App(visible=False)

    closeExcelFile(excelFilePath)

    try:
        wb = xw.Book(excelFilePath, ignore_read_only_recommended=True)
        ws = None
        for sheet in wb.sheets:
            if sheet.name == "Summary":
                ws = sheet
                break
        if ws and ws.range('A1').value == "Claimbot Summary":
            return True
        
    except FileNotFoundError:
        print(f"File not found.")
    
    app.quit()
    return False

def closeExcelFile(excelFilePath):
    for proc in psutil.process_iter():
        try:
            if 'EXCEL.EXE' in proc.name():
                for item in proc.open_files():
                    if excelFilePath.lower() in item.path.lower():
                        proc.kill()
                        return 
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

