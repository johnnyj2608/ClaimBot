import xlwings as xw
import psutil
import pandas as pd

def validateExcelFile(excelFilePath):
    app = xw.App(visible=False)

    closeExcelFile(excelFilePath)

    memberSheet = None
    summaryValues = {}

    try:
        wb = xw.Book(excelFilePath, ignore_read_only_recommended=True)
        summary = None
        
        for sheet in wb.sheets:
            if "summary" == sheet.name.lower():
                summary = sheet
            elif "claim" in sheet.name.lower():
                continue
            else: 
                memberSheet = sheet

        if summary and summary.range('B2').value == 'Professional (CMS)':
            summaryValues = {
                "insurance": summary.range('B3').value,
                "username": summary.range('B4').value,
                "password": summary.range('B5').value,
                "billingProvider": summary.range('B6').value,
                "renderingProvider": summary.range('B9').value,
                "facilities": summary.range('B10').value,
                "servicePlace": summary.range('B11').value,
                "cptCode": summary.range('B12').value,
                "modifier": summary.range('B13').value,
                "diagnosis": summary.range('B14').value,
                "charges": summary.range('B15').value,
                "units": summary.range('B16').value,
            }
        elif summary and summary.range('B2').value == 'Institutional (UB)':
            summaryValues = {
                "insurance": summary.range('B3').value,
                "username": summary.range('B4').value,
                "password": summary.range('B5').value,
                "billingProvider": summary.range('B6').value,
                "physician": summary.range('B9').value,
                "revenueCode": summary.range('B10').value,
                "description": summary.range('B11').value,
                "cptCodeSDC": summary.range('B12').value,
                "cptCodeTrans": summary.range('B13').value,
                "chargesSDC": summary.range('B14').value,
                "chargesTrans": summary.range('B15').value,
                "unitsSDC": summary.range('B16').value,
                "unitsTrans": summary.range('B17').value,
            }
        else:
            return [], {}
        
        members = []
        dataRange = memberSheet.range('B1:M1').expand('down').value
        df = pd.DataFrame(dataRange[1:], columns=dataRange[0])
        df['Birth Date'] = pd.to_datetime(df['Birth Date']).dt.date
        df['Start'] = pd.to_datetime(df['Start']).dt.date
        df['End'] = pd.to_datetime(df['End']).dt.date
        df['Vacation'] = pd.to_datetime(df['Vacation']).dt.date
        df.rename(columns={df.columns[10]: 'VacationEnd'}, inplace=True)
        df['VacationEnd'] = pd.to_datetime(df['VacationEnd']).dt.date
        
        for index, row in df.iterrows():
            members.append(list(row))
        
    except FileNotFoundError:
        print(f"File not found.")
    
    app.quit()
    return members, summaryValues

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