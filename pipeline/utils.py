import openpyxl


def fix_encoding(text):
    try:
        return text.encode('latin1').decode('cp1251')
    except:
        return text

def make_all_sheets_visible(wb: openpyxl.Workbook):
    hidden = [s.title for s in wb.worksheets if s.sheet_state == 'hidden']
    if hidden:
        print(f"Найдены скрытые листы: {hidden}. их не видно в excel.")
    else:
        print("Все листы видны")
        
    for sheet in wb.worksheets:
        sheet.sheet_state = 'visible'
        
    if hidden:
        print(f"Найдены скрытые листы: {hidden}. их не видно в excel.")
    else:
        print("Все листы видны")