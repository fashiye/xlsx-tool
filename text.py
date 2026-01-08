from core.comparator import ExcelComparator
from core.excel_reader import load_workbook_all_sheets
a="D:/code/xlsx/2402大一下成绩1.xlsx"
sheets1 = load_workbook_all_sheets(a)
excel_comparator = ExcelComparator()
sheets2 = excel_comparator.load_workbook(a)
#print(sheets1)
#print(sheets2)
#print(excel_comparator.workbooks)
print(excel_comparator.list_sheets(a))