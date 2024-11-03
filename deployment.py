from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "전개"

# 0번에 카드명, 1번에 방법
data = [[], []]
while True:
    t = input("카드명과 방법을 입력하세요 (종료하려면 빈 입력): ").split()
    if not t:
        break
    for idx, value in enumerate(t):
        if idx % 2 == 0:
            data[0].append(value)
        else:
            data[1].append(value)
    data[0].append(">")
    data[1].append("")

# 5번 행부터 시작 (1~4번 행에 미리 카드 가져다 놓기위해 비워둠)
for col_idx, col_data in enumerate(data, start=5):
    for row_idx, row_data in enumerate(col_data):
        ws.cell(row=col_idx, column=row_idx + 1, value=row_data)

# 열의 너비를 15로 설정
for col_letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
    ws.column_dimensions[col_letter].width = 15

# 카드 이미지 들어갈 위치만 130사이즈로 조절
for row in range(1, ws.max_row + 1, 2):
    ws.row_dimensions[row].height = 130

wb.save("전개법.xlsx")
print('저장 완료')