import customtkinter as ctk
from customtkinter import filedialog 
from customtkinter import CTkToplevel
from tkcalendar import Calendar
from excel import validateExcelFile
import os
from datetime import datetime
from PIL import Image

class ClaimbotGUI:

    today = datetime.today()

    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Claimbot")
        today = datetime.today()

        self.root.geometry(self.centerWindow(self.root, 500, 600, self.root._get_window_scaling()))
        self.frame = ctk.CTkFrame(master=self.root)
        self.frame.pack(pady=20, padx=70, fill="both", expand=True)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.titleLabel = ctk.CTkLabel(master=self.frame, text="Automate Claims")
        self.titleLabel.grid(row=0, column=0, columnspan=5, pady=12, padx=10)

        self.folderLabel = ctk.CTkLabel(master=self.frame, text="No Excel File Selected")
        self.folderLabel.grid(row=1, column=0, columnspan=5, pady=0, padx=10)

        self.browseButton = ctk.CTkButton(master=self.frame, text="Select Excel File", command=self.browseFolder)
        self.browseButton.grid(row=2, column=0, columnspan=5, pady=(0, 10), padx=10)

        self.startDateLabel = ctk.CTkLabel(master=self.frame, text="Start Date")
        self.startDateLabel.grid(row=4, column=0, columnspan=5, pady=(10, 0), padx=10, sticky="ew")

        self.startMonthEntry = ctk.CTkEntry(master=self.frame, width=30)
        self.startMonthEntry.grid(row=5, column=0, pady=0, padx=1, sticky="e")
        self.startMonthEntry.configure(validate="key", validatecommand=(self.frame.register(self.validateMonthDay), "%P"))
        self.startMonthEntry.insert(0, today.month)

        self.startDayEntry = ctk.CTkEntry(master=self.frame, width=30)
        self.startDayEntry.grid(row=5, column=1, pady=0, padx=1, sticky="e")
        self.startDayEntry.configure(validate="key", validatecommand=(self.frame.register(self.validateMonthDay), "%P"))
        self.startDayEntry.insert(0, 1)

        self.startYearEntry = ctk.CTkEntry(master=self.frame, width=45)
        self.startYearEntry.grid(row=5, column=2, pady=0, padx=1, sticky="e")
        self.startYearEntry.configure(validate="key", validatecommand=(self.frame.register(self.validateYear), "%P"))
        self.startYearEntry.insert(0, today.year)

        img = ctk.CTkImage(dark_image=Image.open(r"./assets/datePickerIcon.png"))
        self.startDatePickerButton = ctk.CTkButton(self.frame, 
                                                    image=img, 
                                                    text="", 
                                                    command=lambda: self.toggleCalendar('start'),
                                                    width=32, 
                                                    height=32)
        self.startDatePickerButton.grid(row=5, column=4, pady=0, padx=(5, 10), sticky="w")

        self.endDateLabel = ctk.CTkLabel(master=self.frame, text="End Date")
        self.endDateLabel.grid(row=7, column=0, columnspan=5, pady=(10, 0), padx=10, sticky="ew")

        self.endMonthEntry = ctk.CTkEntry(master=self.frame, width=30)
        self.endMonthEntry.grid(row=8, column=0, pady=0, padx=1, sticky="e")
        self.endMonthEntry.configure(validate="key", validatecommand=(self.frame.register(self.validateMonthDay), "%P"))
        self.endMonthEntry.insert(0, today.month)

        self.endDayEntry = ctk.CTkEntry(master=self.frame, width=30)
        self.endDayEntry.grid(row=8, column=1, pady=0, padx=1, sticky="e")
        self.endDayEntry.configure(validate="key", validatecommand=(self.frame.register(self.validateMonthDay), "%P"))
        self.endDayEntry.insert(0, today.day)

        self.endYearEntry = ctk.CTkEntry(master=self.frame, width=45)
        self.endYearEntry.grid(row=8, column=2, pady=0, padx=1, sticky="e")
        self.endYearEntry.configure(validate="key", validatecommand=(self.frame.register(self.validateYear), "%P"))
        self.endYearEntry.insert(0, today.year)

        self.endDatePickerButton = ctk.CTkButton(self.frame, 
                                                    image=img, 
                                                    text="", 
                                                    command=lambda: self.toggleCalendar('end'),
                                                    width=32, 
                                                    height=32)
        self.endDatePickerButton.grid(row=8, column=4, pady=0, padx=(5, 10), sticky="w")

        self.frame.grid_columnconfigure((0, 4), weight=1)

    def centerWindow(self, Screen: ctk, width: int, height: int, scale_factor: float = 1.0):
        screen_width = Screen.winfo_screenwidth()
        screen_height = Screen.winfo_screenheight()
        x = int(((screen_width/2) - (width/2)) * scale_factor)
        y = int(((screen_height/2) - (height/1.5)) * scale_factor)
        return f"{width}x{height}+{x}+{y}"
    
    def browseFolder(self):
        filePath = filedialog.askopenfilename(title="Select a File", filetypes=[("Excel files", "*.xlsx")])
        if filePath and validateExcelFile(filePath):
            fileName = os.path.basename(filePath)
            self.folderLabel.configure(text=fileName, text_color="gray84")
        else:
            self.folderLabel.configure(text="Invalid Excel Template", text_color="red")

    def toggleCalendar(self, range):
        calendarWindow = CTkToplevel()
        calendarWindow.title('Choose date')
        calendarWindow.geometry('240x200')
        calendarWindow.grab_set()

        root_x = self.root.winfo_rootx()
        rootWidth = self.root.winfo_width()
        windowWidth = calendarWindow.winfo_width()
        x = root_x + (rootWidth - windowWidth) // 2

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
        if range == 'end':
            self.cal.bind("<<CalendarSelected>>", self.endDateSelected)
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

    def validatePickedDate(self):
        pass

    def validateTypedDate(self):
        pass

    def validateMonth(self, val):
        return val == "" or (val.isdigit() and int(val) <= 12)

    def validateMonthDay(self, val):
        return val == "" or (val.isdigit() and int(val) <= 31)

    def validateYear(self, val):
        return val == "" or (val.isdigit() and len(val) <= 4)

    def run(self):
        self.root.mainloop()

app = ClaimbotGUI()
app.run()