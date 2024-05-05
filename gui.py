import customtkinter as ctk
from customtkinter import filedialog 
from tkcalendar import Calendar
from excel import validateExcelFile
import os
from datetime import datetime
from PIL import Image

class ClaimbotGUI:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Claimbot")

        self.root.geometry(self.centerWindow(self.root, 500, 400, self.root._get_window_scaling()))
        self.frame = ctk.CTkFrame(master=self.root)
        self.frame.pack(pady=20, padx=70, fill="both", expand=True)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.titleLabel = ctk.CTkLabel(master=self.frame, text="Automate Claims")
        self.titleLabel.grid(row=0, column=0, columnspan=2, pady=12, padx=10)

        self.folderLabel = ctk.CTkLabel(master=self.frame, text="No Excel File Selected")
        self.folderLabel.grid(row=1, column=0, columnspan=2, pady=0, padx=10)

        self.browseButton = ctk.CTkButton(master=self.frame, text="Select Excel File", command=self.browseFolder)
        self.browseButton.grid(row=2, column=0, columnspan=2, pady=(0, 12), padx=10)

        self.folderLabel = ctk.CTkLabel(master=self.frame, text="Select Date Range")
        self.folderLabel.grid(row=3, column=0, columnspan=2, pady=10, padx=10)

        img = ctk.CTkImage(dark_image=Image.open(r"./assets/datePickerIcon.png"))
        self.startDatePickerButton = ctk.CTkButton(self.frame, image=img, text="", command=self.toggleStartCalendar)
        self.startDatePickerButton.grid(row=4, column=0, columnspan=2, pady=0, padx=10)

        today = datetime.today()
        self.startCalendar = Calendar(self.frame, selectmode="day", year=today.year, month=today.month, day=1)
        self.startCalendar.grid(row=5, column=0, columnspan=2, pady=10, padx=10)
        self.startCalendar.grid_remove()

        self.frame.grid_columnconfigure((0, 1), weight=1)

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

    def toggleStartCalendar(self):
        if self.startCalendar.winfo_viewable():
            self.startCalendar.grid_remove()
        else:
            self.startCalendar.grid()

    def run(self):
        self.root.mainloop()

app = ClaimbotGUI()
app.run()