import openpyxl
import json
import player

class ExcelInterface:
    def __init__(self, path_to_workbook = ""):
        if len(path_to_workbook) == 0:
            raise Exception("No workbook loaded")
        self.wb = openpyxl.load_workbook(path_to_workbook)
    
    def load_sheets(self):
        self.sheets = self.wb.sheetnames

    def get_entries(self): 
        
        self.entries = dict()

        for name in self.sheets:
            ws = self.wb[name]
            entries_weekly = ws["Q"]

            skipHeader = True
            for val in entries_weekly:
                # ignore the 'Entries' 
                if skipHeader :
                    skipHeader = False
                # only add into dictionary if name is
                elif val.value is None:
                    break
                elif str.upper(val.value) not in self.entries:
                    self.entries[str.upper(val.value)] =  player.Player(name = str.upper(val.value))
    
    def playersToJSON(self):
        players = []
        for key in self.entries.keys():
            players.append(json.dumps(self.entries[key].__dict__))
        return players

if __name__ == "__main__":
    xl = ExcelInterface("Brackets.xlsx")
    xl.load_sheets()
    xl.get_entries()
    print(xl.playersToJSON())



        




