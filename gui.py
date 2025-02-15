import customtkinter as ctk
from customtkinter import filedialog 
from customtkinter import CTkToplevel
from tkinter import Listbox
from tkcalendar import Calendar
from excel import validateExcelFile, ifExcelFileOpen
import os
from datetime import datetime, timedelta
from PIL import Image
from officeAllyBilling.officeAlly import officeAllyAutomate
from threading import Thread
import ctypes
import time
ctypes.windll.shcore.SetProcessDpiAwareness(2)

class ProcessStop:
    def __init__(self):
        self.value = False

class ClaimbotGUI:
    def __init__(self, datePickerIcon):
        self.datePickerIcon = datePickerIcon
        self.root = ctk.CTk()
        self.root.title("Claimbot")
        self.stopFlag = ProcessStop()
        self.runningFlag = False
        self.prevDir = None
        self.autoDownloadPath = ''
        self.selectedMembers = []
        self.membersIndexMap = {}
        self.membersList = set()
        self.members = []
        self.form = {}

        self.root.after(1, lambda: self.root.attributes("-topmost", True))
        self.root.geometry(self.rightAlignWindow(self.root, 350, 625, self.root._get_window_scaling()))

        self.tabView = ctk.CTkTabview(master=self.root)
        self.tabView.pack(pady=(0, 20), padx=20, fill="both", expand=True)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.initAutomateTab()
        self.initSummaryTab()

    def initAutomateTab(self):
        self.automateTab = self.tabView.add("     Automate     ")
        self.automateTab.grid_columnconfigure(0, weight=1)
        
        self.titleLabel = ctk.CTkLabel(master=self.automateTab, text="Automate", font=(None, 25, "bold"))
        self.titleLabel.grid(row=0, column=0, pady=(0, 12), padx=10, sticky="ew")

        self.filePath = ""
        self.folderLabel = ctk.CTkLabel(master=self.automateTab, text="No Excel File Selected")
        self.folderLabel.grid(row=1, column=0, pady=0, padx=10)

        self.browseButton = ctk.CTkButton(master=self.automateTab, text="Select Excel File", command=self.browseFolder)
        self.browseButton.grid(row=2, column=0, pady=(0, 6), padx=10)

        self.initSelectFrame()
        self.initDateFrame()

        self.checkboxFrame = ctk.CTkFrame(master=self.automateTab, fg_color="gray17")
        self.checkboxFrame.grid(row=5, column=0, pady=0, padx=0, sticky="ew")
        self.selectFrame.grid_columnconfigure(0, weight=1)

        self.autoSubmit = ctk.BooleanVar()
        self.autoSubmitCheckbox = ctk.CTkCheckBox(master=self.checkboxFrame, 
                                                  text="Enable auto submit", 
                                                  variable=self.autoSubmit, 
                                                  state="disabled")
        self.autoSubmitCheckbox.grid(row=0, column=0, pady=6, padx=60, sticky="w")

        self.autoDownload = ctk.BooleanVar()
        self.autoDownloadCheckbox = ctk.CTkCheckBox(master=self.checkboxFrame, 
                                                    text="Enable auto download", 
                                                    variable=self.autoDownload, 
                                                    state="disabled", 
                                                    command=lambda: self.autoDownloadToggled())
        self.autoDownloadCheckbox.grid(row=1, column=0, pady=6, padx=60, sticky="w")

        self.automateButton = ctk.CTkButton(master=self.automateTab, text="Automate", command=self.automate, state="disabled")
        self.automateButton.grid(row=7, column=0, pady=(6, 6), padx=10)

        self.statusLabel = ctk.CTkLabel(master=self.automateTab, text="")
        self.statusLabel.grid(row=8, column=0, pady=0, padx=10)

    def initSelectFrame(self):
        self.selectFrame = ctk.CTkFrame(master=self.automateTab, fg_color="gray17")
        self.selectFrame.grid(row=3, column=0, pady=0, padx=10)
        self.selectFrame.grid_columnconfigure(0, weight=1)

        self.selectLabel = ctk.CTkLabel(master=self.selectFrame, text="Select Members")
        self.selectLabel.grid(row=0, column=0, pady=0, padx=10, sticky="ew")

        self.listBoxFrame = ctk.CTkFrame(master=self.selectFrame, fg_color="gray17")
        self.listBoxFrame.grid(row=1, column=0, pady=0, padx=0)
        self.listBoxFrame.grid_columnconfigure(0, weight=1)

        self.searchEntry = ctk.CTkEntry(self.listBoxFrame, width=25, state="disabled")
        self.searchEntry.grid(row=0, column=0, sticky="ew", padx=(20, 0), pady=2)
        self.searchEntry.bind("<KeyRelease>", self.searchList)

        self.listbox = Listbox(self.listBoxFrame, 
                               height=4, 
                               width=25,
                               selectmode="multiple", 
                               fg="gray84", 
                               bg="gray14", 
                               activestyle="none",
                               highlightcolor="#800000",
                               highlightthickness=0,
                               relief='ridge',
                               exportselection=False,
                               state="disabled",)
        self.listbox.grid(row=1, column=0, sticky="ew", padx=(20, 0), pady=0)
        self.listbox.configure(font=("Arial", 10, "bold"))
        self.listbox.bind("<MouseWheel>", self.mouseScrollEvent)
        self.listbox.bind("<<ListboxSelect>>", self.onSelect)
        self.curSelection = set()

        self.scrollbar = ctk.CTkScrollbar(self.listBoxFrame, height=4)
        self.scrollbar.grid(row=1, column=1, sticky="ns")

        self.listbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.configure(command=self.listbox.yview)

        self.listButtonFrame = ctk.CTkFrame(master=self.selectFrame, fg_color="gray17")
        self.listButtonFrame.grid(row=3, column=0, columnspan=1, pady=6, padx=10)
        self.listButtonFrame.grid_columnconfigure((0, 1), weight=1)

        self.allButton = ctk.CTkButton(master=self.listButtonFrame, text="All", width=50, state="disabled")
        self.allButton.grid(row=0, column=0, columnspan=1, pady=0, padx=5, sticky="e")
        self.allButton.configure(command=self.selectAll)

        self.clearButton = ctk.CTkButton(master=self.listButtonFrame, text="Clear", width=50, state="disabled")
        self.clearButton.grid(row=0, column=1, columnspan=1, pady=0, padx=5, sticky="w")
        self.clearButton.configure(command=self.clearSelection)

    def selectAll(self):
        self.listbox.selection_set(0, 'end')
        self.curSelection = set(self.listbox.get(0, 'end'))

    def clearSelection(self):
        self.listbox.selection_clear(0, 'end')
        self.curSelection.clear()

    def searchList(self, event):
        searchText = self.searchEntry.get().lower()
        self.listbox.delete(0, 'end')

        for member in self.members:
            memberName = f" {member['id']}. {member['lastName']}, {member['firstName']}"
            if searchText in memberName.lower():
                self.listbox.insert('end', memberName)
                if memberName in self.curSelection:
                    self.listbox.selection_set('end')

    def onSelect(self, event):
        subset = set(self.listbox.get(0, 'end'))
        filtered = self.membersList - subset

        curSelected = [self.listbox.get(index) for index in self.listbox.curselection()]
        curSelected = set(curSelected)

        deselected = self.curSelection - curSelected - filtered
        selected = curSelected - self.curSelection - filtered

        if deselected:
            self.curSelection.remove(deselected.pop())
        else:
            self.curSelection.add(selected.pop())

        self.listbox.after(100, lambda: self.searchEntry.focus_set())
        
    def mouseScrollEvent(self, event):
        if event.delta > 0:
            self.listbox.yview_scroll(-1, "units")
        elif event.delta < 0:
            self.listbox.yview_scroll(1, "units")
        return "break"

    def initDateFrame(self):
        self.dateFrame = ctk.CTkFrame(master=self.automateTab, fg_color="gray17")
        self.dateFrame.grid(row=4, column=0, columnspan=3, pady=6, padx=10)

        self.startDateLabel = ctk.CTkLabel(master=self.dateFrame, text="Start Date")
        self.startDateLabel.grid(row=0, column=0, columnspan=5, pady=0, padx=10, sticky="ew")

        self.startMonthEntry = ctk.CTkEntry(master=self.dateFrame, width=30)
        self.startMonthEntry.grid(row=1, column=0, pady=0, padx=1, sticky="e")
        self.startMonthEntry.configure(validate="key", validatecommand=(self.automateTab.register(self.validateMonth), "%P"), state="disabled")

        self.startDayEntry = ctk.CTkEntry(master=self.dateFrame, width=30)
        self.startDayEntry.grid(row=1, column=1, pady=0, padx=1, sticky="e")
        self.startDayEntry.configure(validate="key", validatecommand=(self.automateTab.register(self.validateDay), "%P"), state="disabled")

        self.startYearEntry = ctk.CTkEntry(master=self.dateFrame, width=45)
        self.startYearEntry.grid(row=1, column=2, pady=0, padx=1, sticky="e")
        self.startYearEntry.configure(validate="key", validatecommand=(self.automateTab.register(self.validateYear), "%P"), state="disabled")

        img = ctk.CTkImage(dark_image=Image.open(self.datePickerIcon))
        self.startDatePickerButton = ctk.CTkButton(master=self.dateFrame, 
                                                image=img, 
                                                text="", 
                                                command=lambda: self.toggleCalendar('start'),
                                                width=32, 
                                                height=32, 
                                                state="disabled")
        self.startDatePickerButton.grid(row=1, column=4, pady=0, padx=(5, 10), sticky="w")

        self.endDateLabel = ctk.CTkLabel(master=self.dateFrame, text="End Date")
        self.endDateLabel.grid(row=2, column=0, columnspan=5, pady=(10, 0), padx=10, sticky="ew")

        self.endMonthEntry = ctk.CTkEntry(master=self.dateFrame, width=30)
        self.endMonthEntry.grid(row=3, column=0, pady=0, padx=1, sticky="e")
        self.endMonthEntry.configure(validate="key", validatecommand=(self.automateTab.register(self.validateMonth), "%P"), state="disabled")

        self.endDayEntry = ctk.CTkEntry(master=self.dateFrame, width=30)
        self.endDayEntry.grid(row=3, column=1, pady=0, padx=1, sticky="e")
        self.endDayEntry.configure(validate="key", validatecommand=(self.automateTab.register(self.validateDay), "%P"), state="disabled")

        self.endYearEntry = ctk.CTkEntry(master=self.dateFrame, width=45)
        self.endYearEntry.grid(row=3, column=2, pady=0, padx=1, sticky="e")
        self.endYearEntry.configure(validate="key", validatecommand=(self.automateTab.register(self.validateYear), "%P"), state="disabled")

        self.endDatePickerButton = ctk.CTkButton(master=self.dateFrame, 
                                                image=img, 
                                                text="", 
                                                command=lambda: self.toggleCalendar('end'),
                                                width=32, 
                                                height=32, 
                                                state="disabled")
        self.endDatePickerButton.grid(row=3, column=4, pady=0, padx=(5, 10), sticky="w")

        self.dateFrame.grid_columnconfigure((0, 4), weight=1)

    def initSummaryTab(self):
        self.summaryTab = self.tabView.add("     Summary     ")
        self.summaryTab.grid_columnconfigure(0, weight=1)

        self.titleLabel = ctk.CTkLabel(master=self.summaryTab, text="Summary", font=(None, 25, "bold"))
        self.titleLabel.grid(row=0, column=0, columnspan=2, pady=(0, 12), padx=10)

        self.summaryFrame = ctk.CTkFrame(master=self.summaryTab, fg_color=self.summaryTab.cget("fg_color"))
        self.summaryFrame.grid(row=1, column=0, pady=(0, 12), padx=10)
        self.summaryFrame.grid_columnconfigure((0, 1), weight=1)

        font = ("Helvetica", 24)
        self.submittedLabel = ctk.CTkLabel(self.summaryFrame, text="Submitted:", font=font)
        self.submittedLabel.grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.membersLabel = ctk.CTkLabel(self.summaryFrame, text="0 / 0", font=font)
        self.membersLabel.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        self.totalLabel = ctk.CTkLabel(self.summaryFrame, text="Total ($):", font=font)
        self.totalLabel.grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.reimbursementLabel = ctk.CTkLabel(self.summaryFrame, text="0.00", font=font)
        self.reimbursementLabel.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        self.unsubmittedLabel = ctk.CTkLabel(master=self.summaryTab, text="Unsubmitted: (0)", font=font, anchor="w")
        self.unsubmittedLabel.grid(row=2, column=0, columnspan=2, pady=(20, 5), padx=10)

        font = ("Helvetica", 14)
        self.scrollFrame = ctk.CTkScrollableFrame(master=self.summaryTab, height=250)
        self.scrollFrame.grid(row=3, column=0, columnspan=2, pady=0, padx=10, sticky="ew")        
        self.scrollFrame.grid_columnconfigure(0, weight=1)

        self.detailsLabel = ctk.CTkLabel(self.scrollFrame, text="", font=font, wraplength=240, justify="left")
        self.detailsLabel.grid(row=4, column=0, padx=2, pady=0, sticky="w")

        font = ("Helvetica", 24)
        self.timeLabel = ctk.CTkLabel(master=self.summaryTab, text="", font=font)
        self.timeLabel.grid(row=4, column=0, pady=10, padx=10)

    def rightAlignWindow(self, Screen: ctk, width: int, height: int, scale_factor: float = 1.0):
        screen_width = Screen.winfo_screenwidth()
        screen_height = Screen.winfo_screenheight()
        x = int(screen_width - width * scale_factor) - 50
        y = int((screen_height - height * scale_factor) / 2) - 100
        return f"{width}x{height}+{x}+{y}"

    def browseFolder(self):
        self.filePath = filedialog.askopenfilename(title="Select a File", filetypes=[("Excel files", "*.xlsx")])
        if self.filePath:
            self.members, self.form = validateExcelFile(self.filePath)
            if self.members:
                self.enableUserInteraction()
                fileName = os.path.basename(self.filePath)
                self.folderLabel.configure(text=fileName, text_color="gray84")

                self.clearSelection()
                self.listbox.delete(0, 'end')
                for index, member in enumerate(self.members):
                    memberName = f" {member['id']}. {member['lastName']}, {member['firstName']}"
                    self.listbox.insert('end', memberName)
                    self.membersList.add(memberName)
                    self.membersIndexMap[memberName] = index
                self.selectAll()

                today = datetime.now()
                firstDay = today.replace(day=1)
                
                self.startMonthEntry.delete(0, "end")
                self.startMonthEntry.insert(0, firstDay.month)

                self.startDayEntry.delete(0, "end")
                self.startDayEntry.insert(0, 1)

                self.startYearEntry.delete(0, "end")
                self.startYearEntry.insert(0, firstDay.year)

                self.endMonthEntry.delete(0, "end")
                self.endMonthEntry.insert(0, today.month)

                self.endDayEntry.delete(0, "end")
                self.endDayEntry.insert(0, today.day)

                self.endYearEntry.delete(0, "end")
                self.endYearEntry.insert(0, today.year)
            else:
                self.folderLabel.configure(text="No members in template", text_color="red")
                self.disableUserInteraction()
                self.browseButton.configure(state="normal")

    def toggleCalendar(self, range):
        calendarWindow = CTkToplevel()
        calendarWindow.title('Choose date')
        calendarWindow.geometry('300x200')
        calendarWindow.grab_set()
        calendarWindow.attributes("-topmost", True)

        x = self.root.winfo_rootx()
        if range == 'start':
            buttonWidget = self.startDatePickerButton
        elif range == 'end':
            buttonWidget = self.endDatePickerButton
        else:
            return
        y = buttonWidget.winfo_rooty() + buttonWidget.winfo_height()

        calendarWindow.geometry(f'+{x}+{y}')
        
        self.cal = Calendar(calendarWindow, 
                       selectmode="day", 
                       date_pattern="m/d/y", 
                       firstweekday="sunday",
                       font=("Arial", 14),
                       showweeknumbers=False)
        self.cal.grid(row=0, column=0, sticky="nswe")

        calendarWindow.grid_columnconfigure(0, weight=1)
        calendarWindow.grid_rowconfigure(0, weight=1)

        if range == 'start':
            self.cal.bind("<<CalendarSelected>>", self.startDateSelected)
            try:
                startMonth = int(self.startMonthEntry.get())
                startDay = int(self.startDayEntry.get())
                startYear = int(self.startYearEntry.get())
                startDate = datetime(startYear, startMonth, startDay)
            except ValueError:
                return False, "Start date does not exist"
            self.cal.selection_set(startDate.date())
        if range == 'end':
            self.cal.bind("<<CalendarSelected>>", self.endDateSelected)
            try:
                endMonth = int(self.endMonthEntry.get())
                endDay = int(self.endDayEntry.get())
                endYear = int(self.endYearEntry.get())
                endDate = datetime(endYear, endMonth, endDay)
            except ValueError:
                return False, "End date does not exist"
            self.cal.selection_set(endDate.date())
        return

    def startDateSelected(self, event=None):
        selectedDate = self.cal.get_date()
        month, day, year = selectedDate.split('/')

        self.startMonthEntry.delete(0, "end")
        self.startMonthEntry.insert(0, month)

        self.startDayEntry.delete(0, "end")
        self.startDayEntry.insert(0, day)

        self.startYearEntry.delete(0, "end")
        self.startYearEntry.insert(0, year)

        self.cal.grid_remove()
        self.cal.master.destroy()

    def endDateSelected(self, event=None):
        selectedDate = self.cal.get_date()
        month, day, year = selectedDate.split('/')

        self.endMonthEntry.delete(0, "end")
        self.endMonthEntry.insert(0, month)

        self.endDayEntry.delete(0, "end")
        self.endDayEntry.insert(0, day)

        self.endYearEntry.delete(0, "end")
        self.endYearEntry.insert(0, year)
        
        self.cal.grid_remove()
        self.cal.master.destroy()

    def autoDownloadToggled(self, *args):
        if self.autoDownload.get():
            initialDir = self.prevDir
            self.autoDownloadPath = filedialog.askdirectory(initialdir = initialDir)
            if self.autoDownloadPath:
                self.prevDir = os.path.dirname(self.autoDownloadPath)
                folderName = os.path.basename(self.autoDownloadPath)

                if len(folderName) > 15:
                    folderName = folderName[:15] + '...'

                self.autoDownloadCheckbox.configure(text="Saving to: "+folderName)
            else:
                self.autoDownload.set(False)
        else:
            self.autoDownloadCheckbox.configure(text="Enable auto download")
            self.autoDownloadPath = ''

    def validateMember(self, val):
        return val == "" or (val.isdigit() and len(val) <= 3)

    def validateMonth(self, val):
        return val == "" or (val.isdigit() and len(val) <= 2)

    def validateDay(self, val):
        return val == "" or (val.isdigit() and len(val) <= 2)

    def validateYear(self, val):
        return val == "" or (val.isdigit() and len(val) <= 4)
    
    def validateInputs(self):
        if not self.curSelection:
            return False, "No members selected"

        selectedSet = self.curSelection.intersection(self.membersList)
        self.selectedMembers = []
        for member in self.members:
            memberName = f" {member['id']}. {member['lastName']}, {member['firstName']}"
            if memberName in selectedSet:
                self.selectedMembers.append(member)

        if not self.selectedMembers:
            return False, "No members in range"
        
        try:
            startMonth = int(self.startMonthEntry.get())
            startDay = int(self.startDayEntry.get())
            startYear = int(self.startYearEntry.get())
            startDate = datetime(startYear, startMonth, startDay)
        except ValueError:
            return False, "Start date does not exist"
        
        try:
            endMonth = int(self.endMonthEntry.get())
            endDay = int(self.endDayEntry.get())
            endYear = int(self.endYearEntry.get())
            endDate = datetime(endYear, endMonth, endDay)
        except ValueError:
            return False, "End date does not exist"
        
        if endDate < startDate:
            return False, "End date set before start date"
    
        if (endDate - startDate).days > 31:
            return False, "Range of dates are too large"
        
        if endDate > datetime.today():
            return False, "End date is after today's date"
        
        if ifExcelFileOpen(self.folderLabel.cget("text")):
            return False, "Must close selected Excel file"

        return True, (startDate, endDate)

    def automate(self):
        if self.runningFlag:
            self.stopFlag.value = True
            self.automateButton.configure(text="Stopping...")
            return
        valid, response = self.validateInputs()
        if not valid:
            self.statusLabel.configure(text=response, text_color="red")
            return
        self.runningFlag = True
        self.automateButton.configure(text="Stop", fg_color='#800000', hover_color='#98423d')
        self.statusLabel.configure(text="", text_color="gray84")
        self.disableUserInteraction()
        self.automateButton.configure(state="normal")  

        self.startTime = time.time()

        thread = Thread(target = officeAllyAutomate, args=(
            self.form, 
            self.selectedMembers,
            response[0],
            response[1],
            self.filePath,
            self.autoSubmit.get(),
            self.autoDownloadPath,
            self.statusLabel,
            self.stopFlag,
            self.updateSummary,
            self.automationCallback))
        
        thread.start()

    def automationCallback(self):
        self.runningFlag = False
        self.stopFlag.value = False
        self.enableUserInteraction()
        self.automateButton.configure(text="Automate", fg_color='#1f538d', hover_color='#14375e')

        elapsedTime = time.time() - self.startTime
        elapsedTime = timedelta(seconds=int(elapsedTime))
        parts = []
         
        if elapsedTime.seconds >= 3600:
            hours = elapsedTime.seconds // 3600
            parts.append(f"{hours}h")

        if elapsedTime.seconds >= 60:
            minutes = (elapsedTime.seconds // 60) % 60
            parts.append(f"{minutes}m")

        seconds = elapsedTime.seconds % 60
        parts.append(f"{seconds}s")

        formattedTime = " ".join(parts)

        self.timeLabel.configure(text="Completed in "+formattedTime)
        self.timeLabel.update()

        self.statusLabel.configure(text="Completed in "+formattedTime)
        self.statusLabel.update()

        self.tabView.set("     Summary     ")

    def updateSummary(self, summary):
        successRatio = str(summary.get("success", 0)) + ' / ' + str(summary.get("members", 0))
        reimbursement = "{:,.2f}".format(summary.get("total", 0))
        unsubmittedCount = str(len(summary.get('unsubmitted', [])))
        unsubmittedText = f"Unsubmitted: ({unsubmittedCount})"
        unsubmittedJoined = '\n'.join(summary['unsubmitted'])

        self.membersLabel.configure(text=successRatio)
        self.reimbursementLabel.configure(text=reimbursement)
        self.unsubmittedLabel.configure(text=unsubmittedText)
        self.detailsLabel.configure(text=unsubmittedJoined)

        self.membersLabel.update()
        self.reimbursementLabel.update()
        self.unsubmittedLabel.update()
        self.detailsLabel.update()

    def disableUserInteraction(self):
        self.browseButton.configure(state="disabled")

        self.searchEntry.configure(state="disabled")
        self.listbox.configure(state="disabled")
        self.allButton.configure(state="disabled")
        self.clearButton.configure(state="disabled")

        self.startMonthEntry.configure(state="disabled")
        self.startDayEntry.configure(state="disabled")
        self.startYearEntry.configure(state="disabled")
        self.startDatePickerButton.configure(state="disabled")

        self.endMonthEntry.configure(state="disabled")
        self.endDayEntry.configure(state="disabled")
        self.endYearEntry.configure(state="disabled")
        self.endDatePickerButton.configure(state="disabled")

        self.autoSubmitCheckbox.configure(state="disabled")
        self.autoDownloadCheckbox.configure(state="disabled")
        self.automateButton.configure(state="disabled")

    def enableUserInteraction(self):
        self.browseButton.configure(state="normal")

        self.searchEntry.configure(state="normal")
        self.listbox.configure(state="normal")
        self.allButton.configure(state="normal")
        self.clearButton.configure(state="normal")

        self.startMonthEntry.configure(state="normal")
        self.startDayEntry.configure(state="normal")
        self.startYearEntry.configure(state="normal")
        self.startDatePickerButton.configure(state="normal")

        self.endMonthEntry.configure(state="normal")
        self.endDayEntry.configure(state="normal")
        self.endYearEntry.configure(state="normal")
        self.endDatePickerButton.configure(state="normal")

        self.autoSubmitCheckbox.configure(state="normal")
        self.autoDownloadCheckbox.configure(state="normal")
        self.automateButton.configure(state="normal")

    def run(self):
        self.root.mainloop()
