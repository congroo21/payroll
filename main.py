# 1. 라이브러리 설치 (Google Colab 환경에서 필요시 실행)
# 이 셀들을 가장 먼저 실행해주세요. (세션당 한 번 또는 런타임 재시작 시)
# !pip install tabula-py pandas fpdf2 PyMuPDF

# 2. 한글 폰트 설치 (Google Colab 환경에서 필요시 실행)
# !apt-get install -y fonts-nanum*
# import matplotlib.font_manager as fm
# try:
#     fm._rebuild()
# except AttributeError:
#     pass


# 3. 필요한 라이브러리 임포트
import tabula
import pandas as pd
import re
from fpdf import FPDF
from fpdf.enums import XPos, YPos # new_x, new_y 사용을 위해 임포트
import fitz # PyMuPDF (지급일 추출용)
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys # resource_path 함수 및 OS 확인용
# from google.colab import files # Colab에서 파일 다운로드용 (필요시 주석 해제)

# 4. 데이터 정제 및 추출 함수들 (이전 버전에서 검증됨)

def clean_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, str):
        cleaned_value = value.strip().replace(',', '')
        cleaned_value = cleaned_value.replace(' ', '')
        if re.match(r"^-?\d+\.000$", cleaned_value):
            cleaned_value = cleaned_value.split('.')[0]
        if cleaned_value.isdigit():
            return int(cleaned_value)
        elif re.match(r"^-?\d+$", cleaned_value):
            return int(cleaned_value)
        return cleaned_value
    if isinstance(value, (int, float)):
        return value
    return value

def verify_employee_totals(record):
    if record.get('구분') != '직원':
        return
    emp_name = record.get('성명', 'N/A')
    error_messages = []
    payment_item_keys = ['기본급', '식대', '상여']
    calculated_payment_total = 0
    for key in payment_item_keys:
        value = record.get(key)
        if isinstance(value, (int, float)):
            calculated_payment_total += value
    record_payment_total = record.get('지급합계')
    expected_payment_total = record_payment_total if isinstance(record_payment_total, (int, float)) else 0
    if calculated_payment_total != expected_payment_total:
        error_messages.append(f"    - 지급합계 불일치: 계산된 값({calculated_payment_total:,}) != 추출된 값({expected_payment_total:,})")

    deduction_item_keys = ['국민연금', '건강보험', '고용보험', '장기요양보험료', '소득세', '지방소득세']
    calculated_deduction_total = 0
    for key in deduction_item_keys:
        value = record.get(key)
        if isinstance(value, (int, float)):
            calculated_deduction_total += value
    record_deduction_total = record.get('공제합계')
    expected_deduction_total = record_deduction_total if isinstance(record_deduction_total, (int, float)) else 0
    if calculated_deduction_total != expected_deduction_total:
        error_messages.append(f"    - 공제합계 불일치: 계산된 값({calculated_deduction_total:,}) != 추출된 값({expected_deduction_total:,})")

    calculated_net_pay = expected_payment_total - expected_deduction_total
    record_net_pay = record.get('차인지급액')
    expected_net_pay = record_net_pay if isinstance(record_net_pay, (int, float)) else 0
    if calculated_net_pay != expected_net_pay:
        error_messages.append(f"    - 차인지급액 불일치: 계산된 값({calculated_net_pay:,}) != 추출된 값({expected_net_pay:,})")

    if error_messages:
        print(f"Warning: {emp_name}님 데이터 검증 오류:")
        for msg in error_messages:
            print(msg)

def parse_payroll_data_from_raw_table(raw_df):
    processed_records = []
    data_start_row_index = 5
    rows_per_block = 3
    num_rows = raw_df.shape[0]
    num_cols = raw_df.shape[1]
    for i in range(data_start_row_index, num_rows, rows_per_block):
        if i + rows_per_block > num_rows and (num_rows - i) < 2:
            break
        row1_data = raw_df.iloc[i]
        row2_data = raw_df.iloc[i+1] if i+1 < num_rows else pd.Series([None]*num_cols, index=raw_df.columns)
        row3_data = raw_df.iloc[i+2] if i+2 < num_rows else pd.Series([None]*num_cols, index=raw_df.columns)
        record = {}
        is_total_row = False
        temp_id = clean_value(row1_data.iloc[0])
        if isinstance(temp_id, str) and '합계' in temp_id:
            record['구분'] = '합계'; record['성명'] = temp_id; record['사원번호'] = None; record['입사일'] = None
            is_total_row = True
        else:
            record['구분'] = '직원'; record['사원번호'] = temp_id
            record['성명'] = clean_value(row1_data.iloc[1])
            record['입사일'] = clean_value(row2_data.iloc[0])
        if is_total_row:
            record['기본급'] = clean_value(row1_data.iloc[1]); record['식대'] = clean_value(row2_data.iloc[0]); record['상여'] = None
            record['국민연금'] = clean_value(row1_data.iloc[8]); record['건강보험'] = clean_value(row1_data.iloc[9])
            record['고용보험'] = clean_value(row1_data.iloc[10]); record['장기요양보험료'] = clean_value(row1_data.iloc[11])
            record['소득세'] = clean_value(row1_data.iloc[12]); record['지방소득세'] = clean_value(row1_data.iloc[13])
            record['공제합계'] = clean_value(row2_data.iloc[12]); record['지급합계'] = clean_value(row3_data.iloc[6])
            record['차인지급액'] = clean_value(row3_data.iloc[12])
        else:
            record['기본급'] = clean_value(row1_data.iloc[2]); record['상여'] = clean_value(row1_data.iloc[3])
            record['식대'] = clean_value(row2_data.iloc[2])
            record['국민연금'] = clean_value(row1_data.iloc[9]); record['건강보험'] = clean_value(row1_data.iloc[10])
            record['고용보험'] = clean_value(row1_data.iloc[11]); record['장기요양보험료'] = clean_value(row1_data.iloc[12])
            record['소득세'] = clean_value(row1_data.iloc[13]); record['지방소득세'] = clean_value(row1_data.iloc[14])
            record['공제합계'] = clean_value(row2_data.iloc[14]); record['지급합계'] = clean_value(row3_data.iloc[8])
            record['차인지급액'] = clean_value(row3_data.iloc[14])
            verify_employee_totals(record)
        processed_records.append(record)
    return processed_records

def extract_and_process_payroll_with_tabula(pdf_path):
    try:
        tables = tabula.read_pdf(pdf_path, pages='1', lattice=True, pandas_options={'header': None}, multiple_tables=True)
        if not tables or len(tables) < 2:
            print("알림: PDF에서 충분한 테이블을 찾지 못했습니다. (최소 2개 예상)")
            return None
        raw_main_df = tables[1]
        final_data_list = parse_payroll_data_from_raw_table(raw_main_df.copy())
        return final_data_list
    except Exception as e:
        if "java.io.IOException: Cannot run program \"java\": error=2" in str(e) or \
           "JavaNotFoundError" in str(e) or \
           "FileNotFoundError: [Errno 2] No such file or directory: 'java'" in str(e) :
            print("Error: Java가 설치되어 있지 않거나 Java 경로가 올바르게 설정되지 않았습니다.")
            print("       tabula-py를 사용하려면 Java가 필요합니다. 시스템에 Java를 설치하고 PATH를 설정해주세요.")
        else:
            print(f"Error: 테이블 추출 및 처리 중 예상치 못한 오류가 발생했습니다: {e}")
        return None

def extract_payment_date(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]
        text = page.get_text("text")
        match = re.search(r"\[지급\s*:\s*(\d{4}년\s?\d{1,2}월\s?\d{1,2}일)\]", text)
        if match:
            doc.close()
            return match.group(1).replace(" ", "")
        doc.close()
        print("알림: 지급일을 찾지 못했습니다.")
        return "지급일 정보 없음"
    except Exception as e:
        print(f"Error: 지급일 추출 중 오류 발생: {e}")
        return "지급일 정보 없음"

# 5. FPDF 리소스 경로 함수 및 클래스 정의 (폰트 로드 방식 수정됨)
def resource_path(relative_path):
    """ PyInstaller로 패키징 시 리소스 파일(예: 폰트)의 절대 경로를 반환합니다. """
    try:
        # PyInstaller는 임시 폴더에 파일을 풀고 _MEIPASS라는 경로를 sys에 추가합니다.
        base_path = sys._MEIPASS
    except Exception:
        # PyInstaller로 실행되지 않은 경우(개발 환경) 현재 파일 위치 기준으로 경로 설정
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class PayStubPDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # 업로드된 NanumGothic.ttf 파일 사용 (스크립트와 동일 위치 또는 PyInstaller 번들 내 위치)
        font_file_to_load = resource_path('NanumGothic.ttf')
        try:
            if not os.path.exists(font_file_to_load):
                # Colab 환경에서 /content/ 에 직접 업로드된 경우도 고려 (resource_path가 '.' 기준으로 반환하므로)
                if 'google.colab' in sys.modules and os.path.exists('NanumGothic.ttf'):
                    font_file_to_load = 'NanumGothic.ttf'
                else:
                    raise RuntimeError(f"폰트 파일을 찾을 수 없습니다: {font_file_to_load}")

            self.add_font('NanumGothic', '', font_file_to_load)
            self.add_font('NanumGothic', 'B', font_file_to_load) # 굵은 스타일도 우선 동일 파일로 등록
            self.font_family_regular = 'NanumGothic'
            self.font_family_bold = 'NanumGothic' # 굵은 스타일 사용 시 'B' 지정
        except RuntimeError as e:
            print(f"FPDF 폰트 설정 오류: {e}. 'NanumGothic.ttf' 파일을 올바른 위치에 두었는지 확인해주세요.")
            print(f" (FPDF가 찾으려고 시도한 폰트 경로: {font_file_to_load})")
            self.font_family_regular = 'Arial' # Fallback
            self.font_family_bold = 'Arial'   # Fallback

    def set_regular_font(self, size=10): # 폰트 스타일 설정을 위한 헬퍼 메소드
        self.set_font(self.font_family_regular, '', size)

    def set_bold_font(self, size=10): # 폰트 스타일 설정을 위한 헬퍼 메소드
        self.set_font(self.font_family_bold, 'B', size) # 'B' 스타일 명시

    # ... (footer, chapter_title, employee_details, payment_details_table,
    #      calculation_methods, work_days_hours, generate_paystub_pdf 메소드는 이전과 동일하게 유지) ...
    # --- 아래는 FPDF의 나머지 메소드들 (이전 버전에서 호환성 업데이트 완료됨) ---
    def footer(self):
        self.set_y(-15)
        self.set_regular_font(8)
        company_name = "히어로 법무사사무소"
        self.cell(0, 10, company_name, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')

    def chapter_title(self, title):
        self.set_bold_font(16)
        self.cell(0, 10, title, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.ln(5)

    def employee_details(self, data, payment_date):
        self.set_regular_font(10)
        page_width = self.w - 2 * self.l_margin
        self.set_x(self.l_margin + page_width - 60)
        self.cell(60, 7, f"지급일 : {payment_date}", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='R')
        self.ln(2)
        col_width1 = 35; col_width2 = 55; line_height = 7
        self.set_bold_font(10)
        self.cell(col_width1, line_height, "성명", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_regular_font(10)
        self.cell(col_width2, line_height, str(data.get('성명', '')), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_bold_font(10)
        self.cell(col_width1, line_height, "생년월일(사번)", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_regular_font(10)
        self.cell(col_width2, line_height, "", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.set_bold_font(10)
        self.cell(col_width1, line_height, "부서", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_regular_font(10)
        self.cell(col_width2, line_height, "", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_bold_font(10)
        self.cell(col_width1, line_height, "직급", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_regular_font(10)
        self.cell(col_width2, line_height, "", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.ln(5)

    def payment_details_table(self, data):
        self.set_bold_font(11); self.cell(0, 7, "세 부 내 역", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C'); self.ln(2)
        header = ["지 급 항 목", "금 액", "공 제 항 목", "금 액"]; col_widths = [55, 35, 55, 35]; line_height = 7
        self.set_bold_font(10)
        for i, header_text in enumerate(header):
            if i == len(header) - 1: self.cell(col_widths[i], line_height, header_text, border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
            else: self.cell(col_widths[i], line_height, header_text, border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_regular_font(9)
        payments = [("기 본 급", data.get('기본급')), ("식    대", data.get('식대')), ("상    여", data.get('상여')), ("", None), ("", None)]
        deductions = [("국민 연금", data.get('국민연금')), ("건강 보험", data.get('건강보험')), ("고용 보험", data.get('고용보험')), ("장기요양 보험료", data.get('장기요양보험료')), ("소 득 세", data.get('소득세')), ("지방 소득세", data.get('지방소득세'))]
        max_rows = max(len(payments), len(deductions))
        for i in range(max_rows):
            pay_item = payments[i][0] if i < len(payments) else ""; pay_val = payments[i][1] if i < len(payments) else None
            pay_amount = f"{pay_val:,}" if pay_val is not None else ""
            ded_item = deductions[i][0] if i < len(deductions) else ""; ded_val = deductions[i][1] if i < len(deductions) else None
            ded_amount = f"{ded_val:,}" if ded_val is not None else ""
            self.cell(col_widths[0], line_height, pay_item, border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='L')
            self.cell(col_widths[1], line_height, pay_amount, border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='R')
            self.cell(col_widths[2], line_height, ded_item, border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='L')
            self.cell(col_widths[3], line_height, ded_amount, border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='R')
        self.set_bold_font(9)
        total_payment_val = data.get('지급합계'); total_deduction_val = data.get('공제합계'); net_pay_val = data.get('차인지급액')
        total_payment_str = f"{total_payment_val:,}" if total_payment_val is not None else "0"
        total_deduction_str = f"{total_deduction_val:,}" if total_deduction_val is not None else "0"
        net_pay_str = f"{net_pay_val:,}" if net_pay_val is not None else "0"
        self.cell(col_widths[0], line_height, "지급액 계", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.cell(col_widths[1], line_height, total_payment_str, border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='R')
        self.cell(col_widths[2], line_height, "공제액 계", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.cell(col_widths[3], line_height, total_deduction_str, border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='R')
        self.cell(col_widths[0] + col_widths[1], line_height, "실 수 령 액", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.cell(col_widths[2] + col_widths[3], line_height, net_pay_str, border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='R')
        self.ln(5)

    def calculation_methods(self):
        self.set_bold_font(10); self.cell(0, 7, "계 산 방 법", border="B", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C'); self.ln(3)
        self.set_regular_font(8)
        methods = ["  · 근로소득세: 간이세액표 적용", "  · 지방소득세: 근로소득세 × 10%", "  · 국민연금: 취득신고 월 보수 × 4.5%", "  · 고용보험: 취득신고 월 보수 × 0.8%", "  · 건강보험: 취득신고 월 보수 × 3.43%", "  · 장기요양보험: 건강보험료 × 11.52%"]
        for method in methods: self.multi_cell(0, 5, method, border=0, align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_regular_font(7); self.multi_cell(0, 5, "  ※ 해당 사업장 상황에 따라 기재가 필요없는 항목이 있을 수 있습니다.", border=0, align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT); self.ln(5)

    def work_days_hours(self):
        self.set_regular_font(8); line_height = 6
        headers = ["근로일수", "총 근로시간수", "연장근로시간수", "야간근로시간수", "휴일근로시간수"]
        col_width = (self.w - 2 * self.l_margin) / len(headers)
        for i, header in enumerate(headers):
            if i == len(headers) - 1: self.cell(col_width, line_height, header, border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
            else: self.cell(col_width, line_height, header, border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        for i, _ in enumerate(headers):
            if i == len(headers) - 1: self.cell(col_width, line_height, "", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
            else: self.cell(col_width, line_height, "", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.ln(10)

    def generate_paystub_pdf(self, employee_data, payment_date, filename="급여명세서.pdf"):
        self.add_page()
        self.chapter_title("임  금  명  세  서")
        self.employee_details(employee_data, payment_date)
        self.payment_details_table(employee_data)
        self.work_days_hours()
        self.calculation_methods()
        self.output(filename, 'F')
        print(f"Info: '{filename}' 파일이 생성되었습니다.")

# 6. Tkinter UI 클래스 및 실행 코드
class PayrollApp:
    def __init__(self, master):
        self.master = master; master.title("급여 명세서 자동 생성 프로그램 v0.3"); master.geometry("550x300")
        self.input_pdf_path = ""; self.output_dir = "generated_paystubs"
        input_frame = tk.Frame(master); input_frame.pack(pady=10, padx=10, fill=tk.X)
        self.btn_select_file = tk.Button(input_frame, text="1. 급여대장 PDF 선택", command=self.select_input_file, width=20)
        self.btn_select_file.pack(side=tk.LEFT, padx=5)
        self.label_file = tk.Label(input_frame, text="선택된 파일: 없음", anchor="w", justify=tk.LEFT)
        self.label_file.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.btn_generate = tk.Button(master, text="2. 급여 명세서 생성 시작", command=self.generate_paystubs, state=tk.DISABLED, height=2, bg="lightblue")
        self.btn_generate.pack(pady=10, padx=10, fill=tk.X)
        self.status_text = tk.Text(master, height=6, wrap=tk.WORD, state=tk.DISABLED)
        self.status_text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        self.btn_open_folder = tk.Button(master, text="저장 폴더 열기", command=self.open_output_folder, state=tk.DISABLED)
        self.btn_open_folder.pack(pady=5, padx=10)

    def _update_status(self, message, is_error=False):
        self.status_text.config(state=tk.NORMAL)
        if is_error: self.status_text.insert(tk.END, f"오류: {message}\n", "error"); self.status_text.tag_config("error", foreground="red")
        else: self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END); self.status_text.config(state=tk.DISABLED); self.master.update_idletasks()

    def select_input_file(self):
        filepath = filedialog.askopenfilename(title="급여대장 PDF 파일을 선택하세요", filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")))
        if filepath:
            self.input_pdf_path = filepath; self.label_file.config(text=f"선택: {os.path.basename(filepath)}"); self.btn_generate.config(state=tk.NORMAL)
            self._update_status(f"'{os.path.basename(filepath)}' 파일이 선택되었습니다. '생성 시작' 버튼을 누르세요.")
        else:
            self.input_pdf_path = ""; self.label_file.config(text="선택된 파일: 없음"); self.btn_generate.config(state=tk.DISABLED)
            self._update_status("파일 선택이 취소되었습니다.")

    def generate_paystubs(self):
        if not self.input_pdf_path: messagebox.showerror("오류", "먼저 급여대장 PDF 파일을 선택해주세요."); self._update_status("급여대장 PDF 파일이 선택되지 않았습니다.", is_error=True); return
        self._update_status("급여 명세서 생성을 시작합니다... (잠시 기다려주세요)"); self.btn_generate.config(state=tk.DISABLED); self.btn_select_file.config(state=tk.DISABLED)
        try:
            payroll_data_list = extract_and_process_payroll_with_tabula(self.input_pdf_path)
            payment_date_on_ledger = extract_payment_date(self.input_pdf_path)
            if not payroll_data_list: messagebox.showerror("데이터 추출 오류", "급여 데이터를 PDF에서 추출하지 못했습니다."); self._update_status("급여 데이터 추출 실패.", is_error=True); self.btn_generate.config(state=tk.NORMAL); self.btn_select_file.config(state=tk.NORMAL); return
            if not os.path.exists(self.output_dir): os.makedirs(self.output_dir)
            self._update_status(f"'{self.output_dir}' 폴더에 생성된 파일을 저장합니다.")
            num_generated = 0; generated_files_info = []
            for employee_record in payroll_data_list:
                if employee_record.get('구분') == '직원':
                    emp_name = str(employee_record.get('성명', '정보없음')).replace(" ", "_"); emp_id = str(employee_record.get('사원번호', 'ID없음'))
                    output_filename = os.path.join(self.output_dir, f"{emp_name}_{emp_id}_급여명세서.pdf")
                    pdf = PayStubPDF(); pdf.generate_paystub_pdf(employee_record, payment_date_on_ledger, output_filename)
                    generated_files_info.append(output_filename); num_generated += 1
            if num_generated > 0:
                final_message = f"{num_generated}명의 급여 명세서가 성공적으로 생성되었습니다:\n" + "\n".join([f" - {os.path.basename(f)}" for f in generated_files_info])
                messagebox.showinfo("성공", final_message.split('\n')[0]); self._update_status(final_message); self.btn_open_folder.config(state=tk.NORMAL)
            else: messagebox.showwarning("알림", "처리할 직원 데이터가 없습니다."); self._update_status("처리할 직원 데이터 없음.")
        except Exception as e: messagebox.showerror("치명적 오류", f"명세서 생성 중 예외 발생: {e}"); self._update_status(f"예외 발생: {e}", is_error=True)
        finally: self.btn_generate.config(state=tk.NORMAL); self.btn_select_file.config(state=tk.NORMAL)

    def open_output_folder(self):
        abs_output_dir = os.path.abspath(self.output_dir)
        if os.path.exists(abs_output_dir):
            try:
                if os.name == 'nt': os.startfile(abs_output_dir)
                elif os.name == 'posix': import subprocess; subprocess.call(['open' if sys.platform == 'darwin' else 'xdg-open', abs_output_dir])
                self._update_status(f"'{abs_output_dir}' 폴더를 열었습니다.")
            except Exception as e: self._update_status(f"폴더 열기 실패: {e}", is_error=True); messagebox.showerror("오류", f"저장 폴더를 여는 데 실패했습니다: {e}\n경로: {abs_output_dir}")
        else: messagebox.showwarning("알림", f"저장 폴더 ('{abs_output_dir}')가 아직 생성되지 않았습니다."); self._update_status(f"저장 폴더 '{abs_output_dir}' 없음.", is_error=True)

# --- 메인 실행 부분 ---
if __name__ == "__main__":
    # Colab 환경인지 로컬 환경인지 감지하여 실행 방식 결정
    if 'google.colab' in sys.modules:
        print("Colab 환경 감지: UI 없이 데이터 처리 및 PDF 생성 테스트를 진행합니다.")
        # 0. Colab에 'NanumGothic.ttf' 와 급여대장 PDF 파일 업로드 필요
        font_file_colab = 'NanumGothic.ttf'
        payroll_pdf_colab = '25.05_히어로법무사사무소_급여대장.pdf'

        if not os.path.exists(font_file_colab):
            print(f"Error: '{font_file_colab}' 파일이 현재 Colab 세션 디렉토리에 없습니다. 업로드해주세요.")
        elif not os.path.exists(payroll_pdf_colab):
            print(f"Error: '{payroll_pdf_colab}' 파일이 현재 Colab 세션 디렉토리에 없습니다. 업로드해주세요.")
        else:
            output_dir_colab = "generated_paystubs_colab"
            if not os.path.exists(output_dir_colab): os.makedirs(output_dir_colab)

            print(f"'{payroll_pdf_colab}' 파일에서 급여 정보 추출 및 정제를 시작합니다...")
            payroll_data_list = extract_and_process_payroll_with_tabula(payroll_pdf_colab)
            payment_date_on_ledger = extract_payment_date(payroll_pdf_colab)

            if payroll_data_list:
                print(f"\n--- 추출된 급여 데이터 (딕셔너리 리스트) ---")
                for record_idx, record_data in enumerate(payroll_data_list): print(f"레코드 {record_idx}: {record_data}")
                print(f"\n--- 추출된 지급일: {payment_date_on_ledger} ---")

                for employee_record in payroll_data_list:
                    if employee_record.get('구분') == '직원':
                        emp_name = str(employee_record.get('성명', '정보없음')).replace(" ", "_"); emp_id = str(employee_record.get('사원번호', 'ID없음'))
                        output_filename = os.path.join(output_dir_colab, f"{emp_name}_{emp_id}_급여명세서.pdf")
                        pdf = PayStubPDF(); pdf.generate_paystub_pdf(employee_record, payment_date_on_ledger, output_filename)
                print(f"\n직원별 급여 명세서 PDF 생성이 완료되었습니다. '{output_dir_colab}' 폴더를 확인하세요.")
                print("Colab 왼쪽 파일 탐색기에서 새로고침 후 폴더 및 파일 확인 가능합니다.")
            else:
                print("\n급여 데이터 추출에 실패하여 명세서를 생성할 수 없습니다.")
    else: # 로컬 환경에서 실행 시 Tkinter UI 실행
        root = tk.Tk()
        app = PayrollApp(root)
        root.mainloop()
