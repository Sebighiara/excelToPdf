#this app reads from an excel file, for each row which contains in a column the value "Selected"
#and geenerates a pdf file with the values from that row

import xlrd
from fpdf import FPDF

path = "Book1.xlsx"

excelWorkbook = xlrd.open_workbook(path)
excelWorksheet = excelWorkbook.sheet_by_index(0)


#create FPDF object
pdf = FPDF('P', 'mm', 'A4')

#add a page
pdf.add_page()

#set font for first row
pdf.set_font('times', 'B', 14)

#add text
for column in range(excelWorksheet.ncols):
	pdf.cell(40, 10, excelWorksheet.cell_value(0, column)) #type the row with titles
pdf.cell(10, 10, '', ln = True)

#set font for the next rows
pdf.set_font('times', '', 11)

for row in range(excelWorksheet.nrows):
	for column in range(excelWorksheet.ncols):
		if excelWorksheet.cell_value(row, column) == "Selected":
			for column in range(excelWorksheet.ncols):
				pdf.cell(40, 10, excelWorksheet.cell_value(row, column))
			pdf.cell(10, 10, '', ln = True)
			break

pdf.output('pdf_conv.pdf')