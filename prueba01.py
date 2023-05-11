import openpyxl

wb = openpyxl.load_workbook("bajada_sap.xlsx")
print(type(wb))
sheets = wb.sheetnames
print(sheets)

wb.active.title = "Boca"

wb.save("bajada_sap2.xlsx")
