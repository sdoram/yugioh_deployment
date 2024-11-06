from openpyxl import Workbook, load_workbook

file_name = input("수정할 파일명을 입력하세요 (확장자 제외): ") + ".xlsx"

try:
    # 기존 파일 열기
    wb = load_workbook(file_name)
    print(f"'{file_name}' 파일을 불러왔습니다.")
except FileNotFoundError:
    # 파일이 없을 경우 새 파일 생성
    wb = Workbook()
    print(f"'{file_name}' 파일이 없어 새 파일을 생성합니다.")

# 기존 시트 선택 또는 새로운 시트 생성
sheet_name = input("수정할 시트명을 입력하세요: ")
if sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
else:
    ws = wb.create_sheet(sheet_name)

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

wb.save(file_name)
print('저장 완료')