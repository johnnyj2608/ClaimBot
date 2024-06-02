import xlwings as xw
import psutil
import pandas as pd
from datetime import date

def validateExcelFile(excelFilePath):
    app = xw.App(visible=False)

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
                    "descriptionSDC": summarySheet.range('B12').value,
                    "cptCodeSDC": summarySheet.range('B13').value,
                    "chargesSDC": summarySheet.range('B14').value,
                    "unitsSDC": summarySheet.range('B15').value,
                    "descriptionTrans": summarySheet.range('B16').value,
                    "cptCodeTrans": summarySheet.range('B17').value,
                    "chargesTrans": summarySheet.range('B18').value,
                    "unitsTrans": summarySheet.range('B19').value,
                })
        else:
            return {}, {}
        
        for key in summary:
            try:
                if key.startswith('charges'):
                    summary[key] = "{:.2f}".format(summary[key])
                else:
                    summary[key] = "{:.0f}".format(summary[key])
            except (ValueError, TypeError):
                pass

        dataRange = memberSheet.range('B1:L1').expand('down').value
        if not any(isinstance(i, list) for i in dataRange):
            return [], {}
        df = pd.DataFrame(dataRange[1:], columns=dataRange[0])
        df['Birth Date'] = pd.to_datetime(df['Birth Date'], errors='coerce')
        df['Start'] = pd.to_datetime(df['Start'], errors='coerce')
        df['End'] = pd.to_datetime(df['End'], errors='coerce')
        df['Vacation'] = pd.to_datetime(df['Vacation'], errors='coerce')
        df.rename(columns={df.columns[9]: 'VacationEnd'}, inplace=True)
        df['VacationEnd'] = pd.to_datetime(df['VacationEnd'], errors='coerce')

        members = []
        for _, row in df.iterrows():
            member = {
                'lastName': row['Last Name'],
                'firstName': row['First Name'],
                'birthDate': row['Birth Date'],
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

def recordClaims(filePath, name, range, total):
    app = xw.App(visible=False)
    today = date.today().strftime("%#m/%#d/%y")

    try:
        wb = xw.Book(filePath, ignore_read_only_recommended=True)
        
        claimsSheet = None
        
        for sheet in wb.sheets:
            if "claims" == sheet.name.lower():
                claimsSheet = sheet
        
        nextRow = claimsSheet.range("A"+str(claimsSheet.cells.last_cell.row)).end("up").row + 1
        claimsSheet.range("A" + str(nextRow)).value = [today, name, range, total]

        wb.save()
        wb.close()

    except Exception as e:
        print("Error:", e)
    finally:
        app.quit()

def ifExcelFileOpen(excelFile):
    for proc in psutil.process_iter():
        try:
            if 'EXCEL.EXE' in proc.name():
                for item in proc.open_files():
                    if excelFile.lower() in item.path.lower():
                        return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False

# def closeExcelFile(excelFile):
#     for proc in psutil.process_iter():
#         try:
#             if 'EXCEL.EXE' in proc.name():
#                 for item in proc.open_files():
#                     if excelFile.lower() in item.path.lower():
#                         proc.kill()
#         except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#             pass