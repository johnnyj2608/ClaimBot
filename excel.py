import xlwings as xw
import psutil
import pandas as pd
from datetime import date

def validateExcelFile(excelFilePath):
    app = xw.App(visible=False)

    try:
        wb = xw.Book(excelFilePath, ignore_read_only_recommended=True)
        
        memberSheet = None
        formSheet = None
        
        for sheet in wb.sheets:
            if "claim" == sheet.name.lower():
                formSheet = sheet
            elif "submitted" in sheet.name.lower():
                continue
            elif "members" in sheet.name.lower(): 
                memberSheet = sheet
            else:
                return {}, {}

        if formSheet.range('A1').value == 'Claimbot':
            form = {
                "form": formSheet.range('B2').value,
                "username": formSheet.range('B3').value,
                "password": formSheet.range('B4').value,
                "payer": formSheet.range('B5').value,
                "billingProvider": formSheet.range('B6').value,
                }
            if form['form'] == 'Professional (CMS)':
                form.update({
                    "renderingProvider": formSheet.range('B9').value,
                    "facilities": formSheet.range('B10').value,
                    "servicePlace": formSheet.range('B11').value,
                    "cptCode": formSheet.range('B12').value,
                    "modifier": formSheet.range('B13').value,
                    "diagnosis": formSheet.range('B14').value,
                    "charges": formSheet.range('B15').value,
                    "units": formSheet.range('B16').value,
                })
            elif form['form'] == 'Institutional (UB)':
                form.update({
                    "physician": formSheet.range('B9').value,
                    "billType": formSheet.range('B10').value,
                    "revenueCode": formSheet.range('B11').value,
                    "descriptionSDC": formSheet.range('B12').value,
                    "cptCodeSDC": formSheet.range('B13').value,
                    "chargesSDC": formSheet.range('B14').value,
                    "unitsSDC": formSheet.range('B15').value,
                    "descriptionTrans": formSheet.range('B16').value,
                    "cptCodeTrans": formSheet.range('B17').value,
                    "chargesTrans": formSheet.range('B18').value,
                    "unitsTrans": formSheet.range('B19').value,
                })
        else:
            return {}, {}
        
        for key in form:
            try:
                if key.startswith('charges'):
                    form[key] = "{:.2f}".format(form[key])
                else:
                    form[key] = "{:.0f}".format(form[key])
            except (ValueError, TypeError):
                pass

        dataRange = memberSheet.range('A1:L1').expand('down').value
        if not any(isinstance(i, list) for i in dataRange):
            return [], {}
        df = pd.DataFrame(dataRange[1:], columns=dataRange[0])
        df['Birth Date'] = pd.to_datetime(df['Birth Date'], errors='coerce')
        df['Start'] = pd.to_datetime(df['Start'], errors='coerce')
        df['End'] = pd.to_datetime(df['End'], errors='coerce')
        df['Vacation'] = pd.to_datetime(df['Vacation'], errors='coerce')
        df.rename(columns={df.columns[10]: 'VacationEnd'}, inplace=True)
        df['VacationEnd'] = pd.to_datetime(df['VacationEnd'], errors='coerce')

        members = []
        for _, row in df.iterrows():
            if row['Exclude']:
                continue
            member = {
                'id': str(int(row['ID'])),
                'lastName': row['Last Name'],
                'firstName': row['First Name'],
                'birthDate': row['Birth Date'],
                'authID': row['Auth #'],
                'dxCode': row['Dx Code'],
                'schedule': str(row['Schedule']),
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
    return members, form

def recordClaims(filePath, name, range, total):
    app = xw.App(visible=False)
    today = date.today().strftime("%#m/%#d/%y")

    try:
        wb = xw.Book(filePath, ignore_read_only_recommended=True)
        
        claimsSheet = None
        
        for sheet in wb.sheets:
            if "submitted" == sheet.name.lower():
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