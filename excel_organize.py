import xlwings as xw

def excel_organize(rows):
    wb = xw.Book('Movies.xlsx')
    for value in wb.sheets[0].range(f'A2:A{rows}'):
        path = value.value
        value.value = (f'=HYPERLINK("{path}")')
    wb.sheets[0].range(f"B1:B{rows}").columns.autofit()
    wb.save('Movies.xlsx')
