import pyxlsb
import openpyxl

def convertir_xlsb_a_xlsx(xlsb_path, xlsx_path):
    wb_xlsx = openpyxl.Workbook(write_only=True)
    with pyxlsb.open_workbook(xlsb_path) as wb_xlsb:
        for sheet_name in wb_xlsb.sheets:
            ws_xlsx = wb_xlsx.create_sheet(title=sheet_name)
            with wb_xlsb.get_sheet(sheet_name) as sheet:
                for row in sheet.rows():
                    ws_xlsx.append([c.v for c in row])
    wb_xlsx.save(xlsx_path)

print("Function compiled correctly!")
