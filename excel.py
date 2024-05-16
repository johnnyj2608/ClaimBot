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
        sumary = None
        
        for sheet in wb.sheets:
            if "summary" == sheet.name.lower():
                summary = sheet
            elif "claim" in sheet.name.lower():
                continue
            else: 
                memberSheet = sheet

        if summary and summary.range('B2').value == 'Professional (CMS)':
            summaryValues = {
                "form": summary.range('B2').value,
                "insurance": summary.range('B3').value,
                "username": summary.range('B4').value,
                "password": summary.range('B5').value,
                "payer": summary.range('B6').value,
                "billingProvider": summary.range('B7').value,
                "renderingProvider": summary.range('B10').value,
                "facilities": summary.range('B11').value,
                "servicePlace": summary.range('B12').value,
                "cptCode": summary.range('B13').value,
                "modifier": summary.range('B14').value,
                "diagnosis": summary.range('B15').value,
                "charges": summary.range('B16').value,
                "units": summary.range('B17').value,
            }
        elif summary and summary.range('B2').value == 'Institutional (UB)':
            summaryValues = {
                "form": summary.range('B2').value,
                "insurance": summary.range('B3').value,
                "username": summary.range('B4').value,
                "password": summary.range('B5').value,
                "payer": summary.range('B6').value,
                "billingProvider": summary.range('B7').value,
                "physician": summary.range('B10').value,
                "revenueCode": summary.range('B11').value,
                "description": summary.range('B12').value,
                "cptCodeSDC": summary.range('B13').value,
                "cptCodeTrans": summary.range('B14').value,
                "chargesSDC": summary.range('B15').value,
                "chargesTrans": summary.range('B16').value,
                "unitsSDC": summary.range('B17').value,
                "unitsTrans": summary.range('B18').value,
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

        for _, row in df.iterrows():
            member = {
                'lastName': row['Last Name'],
                'firstName': row['First Name'],
                'birthDate': row['Birth Date'],
                'medicaid': row['Medicaid'],
                'authID': row['Auth #'],
                'dxCode': row['Dx Code'],
                'schedule': row['Schedule'],
                'authStart': row['Start'],
                'authEnd': row['End'],
                'vacationStart': row['Vacation'],
                'vacationEnd': row['VacationEnd'],
                'exclude': row['Exclude']
            }
            members.append(member)

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