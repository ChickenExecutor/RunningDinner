import openpyxl
import os

class RunningDinnerHelper:
    def __init__(self):
        path = os.path.join(os.path.dirname(__file__), "RunningDinner.xlsx")
        workbook = openpyxl.load_workbook(path)
        self.sheet = workbook.active

        self.team_count = 0
        self.row_vorspeise = self.find_cell_in_column(self.sheet, 1, "Vorspeise")
        self.row_hauptgang = self.find_cell_in_column(self.sheet, 1, "Hauptgang")
        self.row_dessert = self.find_cell_in_column(self.sheet, 1, "Dessert")
        self.end_row = self.find_end()

    def find_cell_in_column(self, sheet, column_number, search_value):
        for row in range(1, sheet.max_row + 1):
            cell_value = sheet.cell(row=row, column=column_number).value
            if cell_value == search_value:
                return row
        return None
    

    def find_end(self):
        for row in range(self.row_dessert, self.sheet.max_row + 1):
            cell_value = self.sheet.cell(row=row, column=1).value
            if cell_value is None:
                return row
        return self.sheet.max_row
    
    def count_teams(self):
        for row in range(self.row_vorspeise, self.end_row+1):
            cell_value = self.sheet.cell(row=row, column=1).value
            if cell_value != None and cell_value != "Vorspeise" and cell_value != "Hauptgang" and cell_value != "Dessert":
                self.team_count += 1
                pass  # Placeholder for counting logic
        pass
    def create_teams(self):
        self.count_teams()
        teams = []
        for x in range(self.team_count):
            teams.append([x])
        
        
            

if __name__ == "__main__":
    helper = RunningDinnerHelper()
    helpers = helper.create_teams()

    




