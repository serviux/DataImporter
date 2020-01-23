
from excel_helper import ExcelInterface

xl = ExcelInterface("Brackets.xlsx")
xl.load_sheets();
xl.get_entries();

print(xl.entries)



