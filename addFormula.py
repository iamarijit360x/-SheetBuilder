def add_formula_to_column(sheet, start_row=13):
    """Add formula to column E based on comparison of columns F and H starting from a specific row."""
    for row in range(start_row, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=5)  # Column E is the 5th column
        cell.value = f'=IF(F{row}<>H{row},"√"," ")'

def add_formula_to_column2(sheet, start_row=13,col='F',col_place=8):
    """Add formula to column E based on comparison of columns F and H starting from a specific row."""
    for row in range(start_row, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=col_place) 
        cell.value = f'=IF({col}{row}>55,{col}{row}*8,IF({col}{row}>30,450,IF({col}{row}>20,350,IF({col}{row}>10,250,IF({col}{row}>0,150,IF({col}{row}<=0,0))))))'

def add_formula_to_column3(sheet, start_row=13,col_place='H'):
    """Add formula to column E based on comparison of columns F and H starting from a specific row."""
    max_row = sheet.max_row
    formula = f'=SUM({col_place}{start_row}:{col_place}{max_row})'
    
    # Add the formula to the cell in column H of the next row after the last data row
    sheet[f'{col_place}{max_row + 1}'] = formula

