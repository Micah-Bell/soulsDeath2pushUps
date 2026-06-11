# Excel
from openpyxl import Workbook, load_workbook

# File handling
import os


class ExcelManager:
    
    def __init__(self, filename):
        self.filename = filename

        if os.path.exists(self.filename):
            self.workbook = load_workbook(self.filename)
            self.sheet = self.workbook.active
        else:
            self.workbook = Workbook()
            self.sheet = self.workbook.active

        self.sheet.append([
            "Game",
            "Death Number"
        ])

        self.workbook.save(self.filename)


    def load_excel(self):
        """Opens excel"""

    def add_death(self, death_num, timestamp):
        """Append a death record"""

    def save(self):
        """saves"""