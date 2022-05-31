import openpyxl

efforts =[]

def CalculateEffort(filename,stylerow,effortrow,character_code_row,characternamerow):
    file = openpyxl.load_workbook(filename)
    sheet = file.active
    rows= sheet.max_row

    for row in range(2,rows):
        rowstring = str(row)
        effort = round(int(sheet[effortrow+rowstring].value)*1.89**(4-int(sheet[stylerow+rowstring].value)))
        efforts.append([sheet[character_code_row+rowstring].value,effort,str(sheet[characternamerow+rowstring].value)])




CalculateEffort("calculatefile.xlsx","F","W","A","D")
sorted=sorted(efforts, key=lambda row: (row[1]),reverse=True)

for card in sorted[0:10]:
    print(f"[{card[2]}] {card[0]} {card[1]}")
