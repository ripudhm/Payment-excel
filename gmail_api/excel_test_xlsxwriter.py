import xlsxwriter
from datetime import datetime


timestamp = datetime.now().strftime("%d.%m.%Y-%H.%M")

workbook = xlsxwriter.Workbook(timestamp + '.xlsx')
worksheet = workbook.add_worksheet("Expense Report")

expenses = [("Rent", 1500), ("Misc", 15), ("Food", 1), ("Air", 0)]

row = 0
col = 0

for item, cost in expenses:
	worksheet.write(row, col, item)
	worksheet.write(row, col + 1, cost)
	row += 1

worksheet.write(row, 0, 'Total')
worksheet.write(row, 1, '=SUM(B1:B4)')

workbook.close()
