import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import ipywidgets as widgets
from IPython.display import display, clear_output
from datetime import datetime

# ==========================================================================
# 1. 도면 기준서 데이터베이스 (기본 세팅 + 신규 등록용 통합)
# ==========================================================================
STANDARDS_DB = {
    "A타입 베어링 도면": [
        {"no": "1", "method": "V/C", "lsl": 10.0, "usl": 10.2},
        {"no": "2", "method": "Micrometer", "lsl": 5.45, "usl": 5.55},
        {"no": "3", "method": "PJ", "lsl": 20.0, "usl": 20.5},
    ],
    "B타입 샤프트 도면": [
        {"no": "1", "method": "Micrometer", "lsl": 14.95, "usl": 15.05},
        {"no": "2", "method": "V/C", "lsl": 4.98, "usl": 5.02},
    ],
    "C타입 실린더 도면": [
        {"no": "1", "method": "마이크로미터", "lsl": 35.15, "usl": 35.25},
        {"no": "2", "method": "버니어 캘리퍼스", "lsl": 8.02, "usl": 8.10}
    ]
}

# ==========================================================================
# 2. 엑셀 생성 핵심 함수
# ==========================================================================
def generate_excel(meta_data, cvt_settings, standards_key):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "치수측정성적서"
    ws.views.sheetView[0].showGridLines = True

    # 스타일 세팅
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    sub_header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    summary_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    meta_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    code_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # 우측 코드표 초록색 강조

    font_title = Font(name="맑은 고딕", size=16, bold=True)
    font_header = Font(name="맑은 고딕", size=10, bold=True)
    font_data = Font(name="맑은 고딕", size=10)
    font_bold = Font(name="맑은 고딕", size=10, bold=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    standards = STANDARDS_DB.get(standards_key, [])

    # 열 수 계산 ('측정부위' 헤더 제외)
    base_headers = ["검사번호", "측정방법", "하한치\n(LSL)", "상한치\n(USL)"]
    start_data_col = len(base_headers) + 1

    current_col = start_data_col
    for cvt_name, count in cvt_settings.items():
        current_col += count
    end_data_col = current_col - 1

    # 통계 요약 헤더 정의 (사용자 요청: 종합판정 옆에 비고란 추가)
    summary_headers = ["최소값\n(MIN)", "최대값\n(MAX)", "평균값\n(AVG)", "종합판정\n(RESULT)", "비고\n(REMARKS)"]
    total_cols = end_data_col + len(summary_headers)

    # 타이틀 영역 (사용자 요청: 상단 고정 타이틀 적용)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    ws["A1"] = "시사출 치수 측정 결과"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center
    ws.row_dimensions[1].height = 40

    # 메타 정보 영역 작성 (3행 ~ 4행, 기준서 제외하고 압축 배치)
    meta_labels = [
        ("차종", meta_data["car_model"], "품명", meta_data["part_name"]),
        ("품번", meta_data["part_number"], "LOT NO", meta_data["lot_no"]),
        ("측정일자", meta_data["date"], "", "")  # 기준서 항목 제거
    ]

    for i, (l1, v1, l2, v2) in enumerate(meta_labels):
        r = 3 + i
        ws.cell(row=r, column=1, value=l1).fill = meta_fill
        ws.cell(row=r, column=2, value=v1)

        # 세 번째 줄에서 우측 빈칸인 경우는 제외
        if l2:
            ws.cell(row=r, column=3, value=l2).fill = meta_fill
            ws.cell(row=r, column=4, value=v2)
        else:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)

        for c in range(1, 5):
            cell = ws.cell(row=r, column=c)
            cell.font = font_data
            cell.border = thin_border
            cell.alignment = align_center

    # 상단표와 치수입력표 사이에 도면 사진용 20칸 공간 여유 확보 (메인 표는 27행부터 시작)
    header_start_row = 27

    # 본문 헤더 영역 배치 (27행 ~ 28행)
    for col_idx, h_text in enumerate(base_headers, 1):
        ws.merge_cells(start_row=header_start_row, start_column=col_idx, end_row=header_start_row+1, end_column=col_idx)
        cell = ws.cell(row=header_start_row, column=col_idx, value=h_text)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = align_center

    current_col = start_data_col
    for cvt_name, count in cvt_settings.items():
        ws.merge_cells(start_row=header_start_row, start_column=current_col, end_row=header_start_row, end_column=current_col + count - 1)
        cell = ws.cell(row=header_start_row, column=current_col, value=cvt_name)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = align_center

        for idx in range(count):
            c_cell = ws.cell(row=header_start_row+1, column=current_col + idx, value=f"#{idx+1}")
            c_cell.font = font_header
            c_cell.fill = sub_header_fill
            c_cell.alignment = align_center
        current_col += count

    for col_idx, s_text in enumerate(summary_headers, end_data_col + 1):
        ws.merge_cells(start_row=header_start_row, start_column=col_idx, end_row=header_start_row+1, end_column=col_idx)
        cell = ws.cell(row=header_start_row, column=col_idx, value=s_text)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = align_center

    for r in [header_start_row, header_start_row+1]:
        for c in range(1, total_cols + 1):
            ws.cell(row=r, column=c).border = thin_border

    # 데이터 행 작성
    start_col_letter = get_column_letter(start_data_col)
    end_col_letter = get_column_letter(end_data_col)

    for i, std in enumerate(standards):
        r_idx = header_start_row + 2 + i

        # [no, method, lsl, usl] 필드 매핑 - 신규 브래킷 누락 오류 수정
        base_values = [std["no"], std["method"], std["lsl"], std["usl"]]
        if "part" in std:
            # 만약 'part' (측정부위) 항목이 등록되어 있다면 비고에 자동 주석 처리
            f_remark = std["part"]
        else:
            f_remark = ""

        for col_idx, val in enumerate(base_values, 1):
            cell = ws.cell(row=r_idx, column=col_idx, value=val)
            cell.font = font_data
            cell.alignment = align_center
            cell.border = thin_border

        for col_idx in range(start_data_col, end_data_col + 1):
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.border = thin_border
            cell.alignment = align_center
            cell.font = font_data

        # 수식 계산 컬럼들의 주소 매핑 (하한치: C열, 상한치: D열)
        f_min = f"=MIN({start_col_letter}{r_idx}:{end_col_letter}{r_idx})"
        f_max = f"=MAX({start_col_letter}{r_idx}:{end_col_letter}{r_idx})"
        f_avg = f"=AVERAGE({start_col_letter}{r_idx}:{end_col_letter}{r_idx})"
        f_result = f'=IF(COUNT({start_col_letter}{r_idx}:{end_col_letter}{r_idx})=0, "", IF(AND({get_column_letter(end_data_col+1)}{r_idx}>=C{r_idx}, {get_column_letter(end_data_col+2)}{r_idx}<=D{r_idx}), "OK", "NG"))'

        formulas = [f_min, f_max, f_avg, f_result, f_remark]
        for idx, formula in enumerate(formulas, end_data_col + 1):
            cell = ws.cell(row=r_idx, column=idx, value=formula)
            cell.font = Font(name="맑은 고딕", size=10, bold=(idx == (total_cols - 1)))
            cell.alignment = align_center
            cell.border = thin_border
            cell.fill = summary_fill

        ws.row_dimensions[r_idx].height = 25

    # 우측 측정 방식 코드 표 작성
    code_start_col = total_cols + 3

    ws.merge_cells(start_row=header_start_row, start_column=code_start_col, end_row=header_start_row, end_column=code_start_col+1)
    lbl_cell = ws.cell(row=header_start_row, column=code_start_col, value="측정 방식 코드 표")
    lbl_cell.font = font_header
    lbl_cell.alignment = align_center
    lbl_cell.fill = code_fill

    col_code = ws.cell(row=header_start_row+1, column=code_start_col, value="코드")
    col_desc = ws.cell(row=header_start_row+1, column=code_start_col+1, value="상세 측정 방법")
    for h_cell in [col_code, col_desc]:
        h_cell.font = font_bold
        h_cell.alignment = align_center
        h_cell.fill = sub_header_fill
        h_cell.border = thin_border
    ws.cell(row=header_start_row, column=code_start_col+1).border = thin_border

    measurement_methods = [
        ("V/C", "버니어 캘리퍼스"),
        ("CMM", "3차원 측정기"),
        ("H/G", "하이트 게이지"),
        ("Micrometer", "마이크로미터"),
        ("PJ", "투영기")
    ]

    for idx, (code, desc) in enumerate(measurement_methods):
        curr_row = header_start_row + 2 + idx
        c_cell = ws.cell(row=curr_row, column=code_start_col, value=code)
        d_cell = ws.cell(row=curr_row, column=code_start_col+1, value=desc)
        for cell in [c_cell, d_cell]:
            cell.font = font_data
            cell.alignment = align_center
            cell.border = thin_border

    # 너비 맞춤 조절
    for c in range(1, total_cols + 1):
        col_letter = get_column_letter(c)
        if c <= len(base_headers):
            ws.column_dimensions[col_letter].width = 15 if c != 2 else 20
        elif c <= end_data_col:
            ws.column_dimensions[col_letter].width = 9
        else:
            ws.column_dimensions[col_letter].width = 14

    ws.column_dimensions[get_column_letter(code_start_col)].width = 15
    ws.column_dimensions[get_column_letter(code_start_col+1)].width = 20

    filename = f"성적서_{meta_data['part_name']}_{meta_data['lot_no']}.xlsx"
    wb.save(filename)
    return filename

# ==========================================================================
# 3. UI 제어 컴포넌트 구성 및 렌더링
# ==========================================================================
output_area = widgets.Output()

select_standard = widgets.Dropdown(options=list(STANDARDS_DB.keys()), description="기준서 선택:")
txt_car_model = widgets.Text(placeholder="예: CVT-Project", description="차종:")
txt_part_name = widgets.Text(placeholder="예: 샤프트 쉴드", description="품명:")
txt_part_number = widgets.Text(placeholder="예: P12345-ABC", description="품번:")
txt_lot_no = widgets.Text(placeholder="예: LOT-2023A", description="LOT NO:")
txt_date = widgets.DatePicker(value=datetime.now(), description="측정일자:")

cvt_count_slider = widgets.IntSlider(value=2, min=1, max=5, description="CVT 개수:")
cvt_samples_container = widgets.VBox()

def update_cvt_inputs(*args):
    items = []
    for idx in range(1, cvt_count_slider.value + 1):
        items.append(widgets.IntText(value=3, min=1, max=20, description=f"CVT-{idx} 샘플수:"))
    cvt_samples_container.children = items

cvt_count_slider.observe(update_cvt_inputs, 'value')
update_cvt_inputs()

btn_generate = widgets.Button(description="엑셀 성적서 자동 생성", button_style="success", icon="check")

def on_btn_click(b):
    with output_area:
        clear_output()

        cvt_settings = {}
        for widget in cvt_samples_container.children:
            cvt_name = widget.description.replace(" 샘플수:", "")
            cvt_settings[cvt_name] = widget.value

        meta_data = {
            'car_model': txt_car_model.value or "N/A",
            'part_name': txt_part_name.value or "N/A",
            'part_number': txt_part_number.value or "N/A",
            'lot_no': txt_lot_no.value or "N/A",
            'date': txt_date.value.strftime("%Y-%m-%d") if txt_date.value else "N/A"
        }

        try:
            filename = generate_excel(meta_data, cvt_settings, select_standard.value)
            print(f"🎉 성공적으로 생성되었습니다! 파일명: {filename}")
            print("\n좌측 파일 탭을 새로고침하여 다운로드 받으세요.")
        except Exception as e:
            print(f"❌ 오류 발생: {e}")

btn_generate.on_click(on_btn_click)

display(
    widgets.VBox([
        widgets.HTML("<h3>📝 엑셀 성적서 자동화 제어반 (통합본)</h3>"),
        select_standard,
        widgets.HBox([txt_car_model, txt_part_name]),
        widgets.HBox([txt_part_number, txt_lot_no]),
        txt_date,
        cvt_count_slider,
        cvt_samples_container,
        widgets.Label(""),
        btn_generate,
        output_area
    ])
)
