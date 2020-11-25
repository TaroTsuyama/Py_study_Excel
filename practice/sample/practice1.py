import openpyxl

wb = openpyxl.Workbook()
ws = wb.worksheets[0]
ws.title = "20201127"
ws.sheet_properties.tabColor = "red"

member_list = [
    "津山 太郎",
    "清水 雄亮",
    "田辺 秀哉",
    "常磐 治希",
    "舘 見菜子",
    "手嶋 洋介",
    "岩井 万理子",
    "駒澤 裕"
]

ws.cell(1,1).value = "No."
ws.cell(1,2).value = "氏名"

count = 0
for member in member_list:
    ws.cell(2+count,1).value = count+1
    ws.cell(2+count,2).value = member
    count += 1

wb.save("Python勉強会参加者リスト1.xlsx")
wb.close()