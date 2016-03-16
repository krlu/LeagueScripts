import xlsxwriter
from DraftData import championData

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('championColor.xlsx')
worksheet = workbook.add_worksheet()


# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

championToColorMap = championData.colorMap

for key in championData.fullNames:
    color_format = workbook.add_format()
    if row%2 == 0:
        color_format.set_font_color('red')
    else:
        color_format.set_font_color('blue')
    color_format.set_pattern(1)  # This is optional when using a solid fill.
    color_format.set_bg_color(championToColorMap[key])
    # Write a total using a formula.
    worksheet.write(row, col, key, color_format)
    row += 1
workbook.close()