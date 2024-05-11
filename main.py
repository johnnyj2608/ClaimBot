from gui import ClaimbotGUI
import os
import sys

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def main():
    datePickerIcon = resource_path("./assets/datePickerIcon.png")
    ClaimbotGUI(datePickerIcon).run()

if __name__ == "__main__":
    main()