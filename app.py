import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import streamlit as st
import io
from datetime import datetime

# ==========================================================================
# 1. 도면 기준서 데이터베이스
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
    code_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") 

    font_title = Font(name="맑은 고딕", size=16, bold=True)
    font_header = Font(name="맑은 고딕", size=10, bold=True)
    font_data = Font(name="맑은 고딕", size=10)
    font_bold = Font(name="맑은 고딕", size=10, bold=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    standards = STANDARDS_DB.get(standards_key, [])

    base_headers = ["검사번호", "측정방법", "하한치\n(LSL)", "상한치\n(USL)"]
    start_data_col = len(base_headers) + 1

    current_col = start_data_col
    for cvt_name, count in cvt_settings.items():
        current_col += count
    end_data_col = current_col - 1

    summary_headers = ["최소값\n(MIN)", "최대값\n(MAX)", "평균값\n(AVG)", "종합판정\n(RESULT)", "비고\n(REMARKS)"]
    total_cols = end_data_col + len(summary_headers)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    ws["A1"] = "시사출 치수 측정 결과"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center
    ws.row_dimensions[1].height = 40

    meta_labels = [
        ("차종", meta_data["car_model"], "품명", meta_data["part_name"]),
        ("품번", meta_data["part_number"], "LOT NO", meta_data["lot_no"]),
        ("측정일자", meta_data["date"], "", "")  
    ]

    for i, (l1, v1, l2, v2) in enumerate(meta_labels):
        r = 3 + i
        ws.cell(row=r, column=1, value=l1).fill = meta_fill
        ws.cell(row=r, column=2, value=v1)

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

    header_start_row = 27

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

    start_col_letter = get_column_letter(start_data_col)
    end_col_letter = get_column_letter(end_data_col)

    for i, std in enumerate(standards):
        r_idx = header_start_row + 2 + i

        base_values = [std["no"], std["method"], std["lsl"], std["usl"]]
        f_remark = std.get("part", "")

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
        h_cell.border = thin_border

    # 메모리 버퍼로 엑셀 반환
    excel_data = io.BytesIO()
    wb.save(excel_data)
    excel_data.seek(0)
    return excel_data

# ==========================================================================
# 3. Streamlit 웹 인터페이스 화면 구성
# ==========================================================================
st.set_page_config(page_title="치수성적서 생성기", layout="centered")
st.title("📊 치수 측정 성적서 자동 생성기")
st.write("정보를 입력하고 하단의 버튼을 누르면 성적서 양식이 자동 생성됩니다.")

st.markdown("---")
st.subheader("1. 기본 정보 입력")

col1, col2 = st.columns(2)
with col1:
    car_model = st.text_input("차종", value="A-CAR")
    part_number = st.text_input("품번", value="P12345-ABC")
with col2:
    part_name = st.text_input("품명", value="BEARING ASSEMBLY")
    lot_no = st.text_input("LOT NO", value="LOT-" + datetime.today().strftime('%Y%m%d'))

check_date = st.date_input("측정일자", datetime.today())

st.markdown("---")
st.subheader("2. 도면 선택 및 조건 세팅")

standards_key = st.selectbox("기준 도면 선택", list(STANDARDS_DB.keys()))

col3, col4 = st.columns(2)
with col3:
    cvt1_sample = st.number_input("CVT-1 샘플 수", min_value=0, max_value=30, value=3)
with col4:
    cvt2_sample = st.number_input("CVT-2 샘플 수", min_value=0, max_value=30, value=3)

# 매핑 데이터 가공
meta_data = {
    "car_model": car_model,
    "part_name": part_name,
    "part_number": part_number,
    "lot_no": lot_no,
    "date": check_date.strftime('%Y-%m-%d')
}

cvt_settings = {}
if cvt1_sample > 0:
    cvt_settings["CVT-1"] = cvt1_sample
if cvt2_sample > 0:
    cvt_settings["CVT-2"] = cvt2_sample

st.markdown("---")

# 엑셀 파일 생성 및 즉시 다운로드 버튼 배치
if cvt1_sample == 0 and cvt2_sample == 0:
    st.warning("CVT 샘플 수를 최소 1개 이상 입력해 주세요.")
else:
    try:
        processed_excel = generate_excel(meta_data, cvt_settings, standards_key)
        
        st.download_button(
            label="🟢 치수측정성적서 엑셀 다운로드",
            data=processed_excel,
            file_name=f"치수측정성적서_{part_number}_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        st.success("양식이 준비되었습니다! 위의 초록색 버튼을 눌러 엑셀 파일을 다운로드하세요.")
    except Exception as e:
        st.error(f"엑셀 생성 중 에러가 발생했습니다: {e}")
