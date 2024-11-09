# t열까지 가면 알아서 줄바꿈 처리하기

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
import os

file_name = input("수정할 파일명을 입력하세요 (확장자 제외): ") + ".xlsx"
image_folder = f"C:/Users/Administrator/Desktop/yugioh_deployment/이미지/{file_name.split('.xlsx')[0]}/"

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
    
# 시트명을 A1에 입력
ws.cell(row=1, column=1, value=sheet_name)
# 핸드 내용을 입력
for col_idx ,hand in enumerate(sheet_name.split(), start=2):
    ws.cell(row=2, column=col_idx, value=hand)
    image_file_path = os.path.join(image_folder, f"{hand}.png")
        # 이미지 파일이 존재하면 엑셀에 추가
    if os.path.exists(image_file_path):
        # 이미지 객체 생성
        img = Image(image_file_path)

        # 이미지 크기 조정
        cell_width = 15
        cell_height = 130
        img.width = cell_width * 8
        img.height = cell_height * 1.33
            
        # 이미지 삽입
        cell_location = f"{chr(65 + col_idx - 1)}{2}"
        ws.add_image(img, cell_location)
    else:
        print(f"이미지를 찾을 수 없습니다: {image_file_path}")

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



for row_idx, col_data in enumerate(data, start=3):
    for  col_idx, row_data in enumerate(col_data, start=2):
        image_file_path = os.path.join(image_folder, f"{row_data}.png")
        # 이미지 파일이 존재하면 엑셀에 추가
        if os.path.exists(image_file_path):
            # 이미지 객체 생성
            img = Image(image_file_path)

            # 이미지 크기 조정
            cell_width = 15
            cell_height = 130
            img.width = cell_width * 8
            img.height = cell_height * 1.33
            
            # 이미지 삽입
            cell_location = f"{chr(65 + col_idx - 1)}{3}"
            ws.add_image(img, cell_location)
        else:
            print(f"이미지를 찾을 수 없습니다: {image_file_path}")
        ws.cell(row=row_idx, column=col_idx, value=row_data)

    # 열의 너비를 15로 설정
for col_letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
    ws.column_dimensions[col_letter].width = 15

# 카드 이미지 들어갈 위치만 130사이즈로 조절
for row in range(3, ws.max_row + 1, 2):
    ws.row_dimensions[row].height = 130
ws.row_dimensions[2].height = 130

wb.save(file_name)
print('저장 완료')