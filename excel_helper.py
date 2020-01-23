import openpyxl
import player

class ExcelInterface:
    """
    This class is used to interface with the excel document cointaining all the players
    matches and win loss ratios.

    args: path_to_workbook --path to the file being taken in.
    """
    def __init__(self, path_to_workbook):
        if len(path_to_workbook) == 0:
            raise Exception("No workbook loaded")
        self.wb = openpyxl.load_workbook(path_to_workbook)
    
    #stores the sheets in the current instance
    def load_sheets(self):
        self.sheets = self.wb.sheetnames

    """
    gets a list of all players from all tabs in the excel doc
    and adds them to a dictionary. Names are forced to be uppercase 
    so they are case Insensitive in the word doc

    creates the self.entries dictionary a dictionary 
    containing a key value pair like <String, Player>
    """
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
    

    """
    used to create a list of all the players converted to a 
    json object. 
    """
    def playersToJSON(self):
        players = []
        for key in self.entries.keys():
            players.append(player.PlayerToJSON.PlayerToJSON(self.entries[key]))
        return players

if __name__ == "__main__":
    xl = ExcelInterface("Brackets.xlsx")
    xl.load_sheets()
    xl.get_entries()
    print(xl.playersToJSON())



        




