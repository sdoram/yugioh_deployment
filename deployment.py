# 전개 끝난 후 필드상황 어떻게 만들지
# 상대턴 움직임 만들기

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from pathlib import Path

file_name = input("수정할 파일명을 입력하세요 (확장자 제외): ") + ".xlsx"
big_data = [[], []]
image_folder = Path.cwd() / "이미지" / file_name.split(".xlsx")[0]
alignment_style = Alignment(horizontal='center', vertical='center')
font_style_card = Font(bold=True, size=30)
font_style_text = Font(bold=True, size=11)

extra_color = PatternFill(start_color="B3CEFB", end_color="B3CEFB", fill_type="solid")
monster_color = PatternFill(start_color="FFC599", end_color="FFC599", fill_type="solid")
magic_trap_color = PatternFill(start_color="A6E3B7", end_color="A6E3B7", fill_type="solid")
tomb_color = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

border_style = Border(
    left=Side(border_style="thin", color="000000"),
    right=Side(border_style="thin", color="000000"),
    top=Side(border_style="thin", color="000000"),
    bottom=Side(border_style="thin", color="000000")
)

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
ws.cell(row=2, column=1, value='핸드')
ws.cell(row=3, column=1, value='전개')
# 핸드 내용을 입력
for col_idx, value in enumerate(sheet_name.split(), start=2):
    ws.cell(row=2, column=col_idx, value=value)
    image_file_path = image_folder / f"{value}.png"
    print(f"이미지 경로: {image_file_path}")
        # 이미지 파일이 존재하면 엑셀에 추가
    if image_file_path.exists():
        # 이미지 객체 생성
        img = Image(str(image_file_path))

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

while True:
    # 0번에 카드명, 1번에 방법
    data = input("카드명과 방법을 입력하세요 (종료하려면 빈 입력): ").split()
    if not data:
        break
    for idx, value in enumerate(data):
        if idx % 2 == 0:
            big_data[0].append(value)
        else:
            big_data[1].append(value)
    big_data[0].append(">")
    big_data[1].append(" ")

# O열 넘기는지 체크용
current_location = big_data[0].index('>')
try:
    next_location = big_data[0].index('>', current_location+1)
except ValueError:
    next_location = current_location

for i, data in enumerate(big_data):
    if i == 0:
        row_idx, col_idx = 3, 2
    max_col = 15
    
    for idx, value in enumerate(data):
        if idx+1 != len(data):
            if value == '>' and col_idx + next_location - current_location > 16:
                col_idx = 1
                row_idx += 2
                current_location = next_location
                try:
                    next_location = big_data[0].index('>', current_location+1)
                except ValueError:
                    next_location = current_location
            image_file_path = image_folder / f"{value}.png"
            # 이미지 파일이 존재하면 엑셀에 추가
            if image_file_path.exists():
                # 이미지 객체 생성
                img = Image(str(image_file_path))

                # 이미지 크기 조정
                cell_width = 15
                cell_height = 130
                img.width = cell_width * 8
                img.height = cell_height * 1.33
                
                # 이미지 삽입
                cell_location = f"{chr(65 + col_idx - 1)}{row_idx}"
                ws.add_image(img, cell_location)
            else:
                if i == 1 and value == ' ' and col_idx + next_location - current_location > 16:
                    col_idx = 1
                    row_idx += 2
                    current_location = next_location
                    try:
                        next_location = big_data[0].index('>', current_location+1)
                    except ValueError:
                        next_location = current_location
                print(f"이미지를 찾을 수 없습니다: {image_file_path}")
                ws.cell(row=row_idx, column=col_idx, value=value)
        col_idx += 1
    # 첫번째 for문이 끝나고 방법에 해당하는 반복문이 진행되기 전에 row, col의 정보를 갱신
    row_idx, col_idx = 4, 2

    # 열의 너비를 15로 설정
for col_letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
    ws.column_dimensions[col_letter].width = 15

# 카드 이미지 들어갈 위치만 130사이즈로 조절
for row in range(3, ws.max_row + 5):
    if row > ws.max_row or row % 2 != 0:
        ws.row_dimensions[row].height = 130
ws.row_dimensions[2].height = 130

# 결과 필드 만들기

ws.cell(row=ws.max_row+1, column=1, value='결과')
result_row = ws.max_row
# 엑스트라 몬스터
ws.cell(row=result_row, column=4).fill = extra_color
ws.cell(row=result_row, column=4).border = border_style
ws.cell(row=result_row, column=6).fill = extra_color
ws.cell(row=result_row, column=6).border = border_style
# 몬스터
for col in ws.iter_rows(min_row=result_row+1, max_row=result_row+1, min_col=3, max_col=7):
    for cell in col:
        cell.fill = monster_color
        cell.border = border_style
# 마함
for col in ws.iter_rows(min_row=result_row+2, max_row=result_row+2, min_col=3, max_col=7):
    for cell in col:
        cell.fill = magic_trap_color
        cell.border = border_style
# 필마
ws.cell(row=result_row+1, column=2).fill = magic_trap_color
ws.cell(row=result_row+1, column=2).border = border_style

# 묘지&제외
for row in ws.iter_rows(min_row=result_row, max_row=result_row+1, min_col=8, max_col=8):
    for cell in row:
        cell.fill = tomb_color
        cell.border = border_style

# 폰트 조정
for enu, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=1):
    for cell in row:
        cell.alignment = alignment_style
        if enu >= 4 and enu % 2 == 0:
            cell.font = font_style_text
        else:
            cell.font = font_style_card

wb.save(file_name)
print('저장 완료')