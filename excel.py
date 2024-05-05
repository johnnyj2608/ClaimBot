import xlwings as xw
import os
import psutil

def getMembersByInsurance(insurance):
    app = xw.App(visible=False)

    current_dir = os.path.dirname(os.path.realpath(__file__))
    excel_file = os.path.join(current_dir, 'test.xlsx')

    closeExcelFile(excel_file)

    try:
        wb = xw.Book(excel_file, ignore_read_only_recommended=True)
        for sheet in wb.sheets:
            if sheet.name == insurance:
                ws = sheet
                break
        print(ws.range('A1').value)
        
    except FileNotFoundError:
        print(f"File not found.")
    
    app.quit()

def closeExcelFile(excelFile):
    for proc in psutil.process_iter():
        try:
            if 'EXCEL.EXE' in proc.name():
                for item in proc.open_files():
                    if excelFile.lower() in item.path.lower():
                        proc.kill()
                        return 
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

getMembersByInsurance("Insurance")