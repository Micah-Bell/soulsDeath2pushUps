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

            games = (
                "Elden Ring",
                "Dark Souls III",
                "Dark Souls II",
                "Dark Souls",
                "Sekiro",
                "Demon's Souls",
            )

            for game in games:
                self.sheet.append([game, 0])

        self.workbook.save(self.filename)


    def load_excel(self):
        """Opens excel"""
        self.workbook = load_workbook(self.filename)
        self.sheet = self.workbook.active


    def add_death(self, game):
        """Append a death record"""

        for row in range(1, self.sheet.max_row + 1):
            game_name = self.sheet.cell(row=row, column=1).value

            if game_name == game:
                current_count = self.sheet.cell(row=row, column=2).value
                self.sheet.cell(row=row, column=2).value = current_count + 1

                self.save()
                return


    def save(self):
        """saves"""
        self.workbook.save(self.filename)