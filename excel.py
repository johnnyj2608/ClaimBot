import xlwings as xw
import psutil
import pandas as pd

def validateExcelFile(excelFilePath):
    app = xw.App(visible=False)

    closeExcelFile(excelFilePath)

    try:
        wb = xw.Book(excelFilePath, ignore_read_only_recommended=True)
        
        memberSheet = None
        summarySheet = None
        
        for sheet in wb.sheets:
            if "summary" == sheet.name.lower():
                summarySheet = sheet
            elif "claim" in sheet.name.lower():
                continue
            else: 
                memberSheet = sheet

        if summarySheet.range('A1').value == 'Claimbot Summary':
            summary = {
                "form": summarySheet.range('B2').value,
                "username": summarySheet.range('B3').value,
                "password": summarySheet.range('B4').value,
                "payer": summarySheet.range('B5').value,
                "billingProvider": summarySheet.range('B6').value,
                }
            if summary['form'] == 'Professional (CMS)':
                summary.update({
                    "renderingProvider": summarySheet.range('B9').value,
                    "facilities": summarySheet.range('B10').value,
                    "servicePlace": summarySheet.range('B11').value,
                    "cptCode": summarySheet.range('B12').value,
                    "modifier": summarySheet.range('B13').value,
                    "diagnosis": summarySheet.range('B14').value,
                    "charges": summarySheet.range('B15').value,
                    "units": summarySheet.range('B16').value,
                })
            elif summary['form'] == 'Institutional (UB)':
                summary.update({
                    "physician": summarySheet.range('B9').value,
                    "billType": summarySheet.range('B10').value,
                    "revenueCode": summarySheet.range('B11').value,
                    "description": summarySheet.range('B12').value,
                    "cptCodeSDC": summarySheet.range('B13').value,
                    "cptCodeTrans": summarySheet.range('B14').value,
                    "chargesSDC": summarySheet.range('B15').value,
                    "chargesTrans": summarySheet.range('B16').value,
                    "unitsSDC": summarySheet.range('B17').value,
                    "unitsTrans": summarySheet.range('B18').value,
                })
        else:
            return [], {}
        
        for key in summary:
            try:
                if key.startswith('charges'):
                    summary[key] = "{:.2f}".format(summary[key])
                else:
                    summary[key] = "{:.0f}".format(summary[key])
            except (ValueError, TypeError):
                pass

        members = []
        dataRange = memberSheet.range('B1:M1').expand('down').value
        df = pd.DataFrame(dataRange[1:], columns=dataRange[0])
        df['Birth Date'] = pd.to_datetime(df['Birth Date'], errors='coerce')
        df['Start'] = pd.to_datetime(df['Start'], errors='coerce')
        df['End'] = pd.to_datetime(df['End'], errors='coerce')
        df['Vacation'] = pd.to_datetime(df['Vacation'], errors='coerce')
        df.rename(columns={df.columns[10]: 'VacationEnd'}, inplace=True)
        df['VacationEnd'] = pd.to_datetime(df['VacationEnd'], errors='coerce')

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
    return members, summary

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