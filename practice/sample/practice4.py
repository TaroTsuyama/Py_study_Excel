import glob
import openpyxl
from openpyxl.styles import Border, Side
from openpyxl.styles.alignment import Alignment

files = glob.glob("./fruits/*")
xlsx_files = [file for file in files if ".xlsx" in file and not "~$" in file]

RULE_LINE = Border(
    outline=True,
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style="thin", color='000000')
)
ALIGN_V_CENTER = Alignment(vertical="center")


statistics_dict = {} # ディクショナリの定義

for xlsx_file in xlsx_files:
    wb = openpyxl.load_workbook(xlsx_file)

    if "統計表" in wb:
        shipment_dict = {} # ディクショナリの初期化
        ws = wb["統計表"]
        rows = ws.rows

        for row in rows:
            if row[0].value == "品目":
                item = row[1].value
            elif type(row[4].value) is int:
                shipment_dict[row[0].value] = row[4].value

        statistics_dict[item] = shipment_dict

new_wb = openpyxl.Workbook()
new_wb.worksheets[0].title = "果物出荷量統計"
new_wb.create_sheet(title="果物出荷量ランキング")

ws = new_wb.worksheets[0]
ws.cell(1,1).value = "令和元年 果物出荷量統計"
ws.cell(3,1).value = "品目"
ws.cell(3,1).border = RULE_LINE
ws.cell(3,2).value = "出荷量(t)"
ws.cell(3,2).border = RULE_LINE
count = 0
for item in statistics_dict.items():
    ws.cell(4 + count,1).value = item[0]
    ws.cell(4 + count,1).border = RULE_LINE
    ws.cell(4 + count,2).value = sum(item[1].values())
    ws.cell(4 + count,2).border = RULE_LINE
    count += 1

ws = new_wb.worksheets[1]
ws.cell(1,1).value = "令和元年 果物出荷量ランキング"
ws.cell(3,1).value = "品目"
ws.cell(3,1).border = RULE_LINE
ws.cell(3,2).value = "順位"
ws.cell(3,2).border = RULE_LINE
ws.cell(3,3).value = "都道府県"
ws.cell(3,3).border = RULE_LINE
ws.cell(3,4).value = "出荷量(t)"
ws.cell(3,4).border = RULE_LINE

count = 0
rank_num = 3
for item in statistics_dict.items():
    sorted_dict = dict(sorted(item[1].items(), key=lambda x:x[1], reverse=True)) # 出荷量で降順ソートしたディクショナリ
    max_num = min(rank_num,len(sorted_dict))

    ws.cell(4+count,1).value = item[0]
    ws.cell(4+count,1).alignment = ALIGN_V_CENTER
    ws.merge_cells(start_row = 4 + count, start_column = 1, end_row = 4 + count + max_num - 1, end_column = 1)
    for i in range(max_num):
        ws.cell(4+count+i,1).border = RULE_LINE
        ws.cell(4+count+i,2).value = i + 1
        ws.cell(4+count+i,3).value = list(sorted_dict.keys())[i]
        ws.cell(4+count+i,4).value = list(sorted_dict.values())[i]
        ws.cell(4+count+i,2).border = RULE_LINE
        ws.cell(4+count+i,3).border = RULE_LINE
        ws.cell(4+count+i,4).border = RULE_LINE

    count += max_num


new_wb.save("令和元年_果物出荷量統計.xlsx")