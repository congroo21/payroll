# main.py
# ============================================================
# 이 파일 하나로 모든 기능을 관리하며,
# requirements.txt에 명시된 라이브러리를 설치한 뒤 실행합니다.
#
# 1) 먼저 프로젝트 루트에서 아래 명령을 사용해 의존성 설치:
#      py -3.11 -m venv venv
#      .\venv\Scripts\Activate.ps1
#      python -m pip install --upgrade pip
#      pip install -r requirements.txt
#
# 2) 실행:
#      python main.py
#
# 3) 이후, exe로 패키징할 때는 'jre' 폴더와 NanumGothic.ttf를 함께 포함해야 합니다.
# ============================================================

import os
import sys
import re
import pandas as pd             # pip install pandas
import tabula                    # pip install tabula-py
from fpdf import FPDF           # pip install fpdf2
from fpdf.enums import XPos, YPos
import fitz                      # pip install PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox

# ============================================================
# 1. Embedded JRE 세팅
# ============================================================
def _setup_embedded_jre():
    """
    - PyInstaller로 묶인 경우: sys._MEIPASS 아래에 'jre' 폴더가 풀려 있습니다.
    - 개발(로컬) 환경: main.py가 있는 디렉터리의 jre 폴더를 참조합니다.
    """
    MEIPASS = getattr(sys, '_MEIPASS', None)
    if MEIPASS:
        embedded_jre_dir = os.path.join(MEIPASS, 'jre')
    else:
        embedded_jre_dir = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'jre')

    java_bin_dir = os.path.join(embedded_jre_dir, 'bin')
    java_exe = os.path.join(java_bin_dir, 'java.exe')

    if os.path.exists(java_exe):
        # PATH에 embedded JRE/bin을 맨 앞에 추가
        prev_path = os.environ.get('PATH', '')
        os.environ['PATH'] = java_bin_dir + os.pathsep + prev_path

        # tabula-py가 참조할 JAR 경로 설정 (jre/lib 밑에 있는 tabula JAR)
        jar_path = os.path.join(embedded_jre_dir, 'lib', 'tabula-1.0.5-jar-with-dependencies.jar')
        if os.path.exists(jar_path):
            os.environ['TABULA_JAR_PATH'] = jar_path

        # 디버깅용 출력 (콘솔 모드일 때 확인)
        print(f"[Info] Embedded JRE 사용 경로: {java_exe}")
    else:
        print("[Warning] Embedded JRE를 찾지 못했습니다. 시스템 설치 Java 사용 예상.")
    return embedded_jre_dir

# JRE 세팅 후, JAVA_HOME 환경 변수 지정
jre_dir = _setup_embedded_jre()
os.environ['JAVA_HOME'] = jre_dir
os.environ['PATH'] = os.path.join(jre_dir, 'bin') + os.pathsep + os.environ.get('PATH', '')

# ============================================================
# 2. 데이터 정제 및 추출 함수들
# ============================================================
def clean_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, str):
        cleaned = value.strip().replace(',', '').replace(' ', '')
        if re.match(r"^-?\d+\.000$", cleaned):
            cleaned = cleaned.split('.')[0]
        if cleaned.isdigit():
            return int(cleaned)
        elif re.match(r"^-?\d+$", cleaned):
            return int(cleaned)
        return cleaned
    if isinstance(value, (int, float)):
        return value
    return value

def verify_employee_totals(record):
    if record.get('구분') != '직원':
        return
    emp_name = record.get('성명', '정보없음')
    errors = []

    # 지급합계 검증
    payment_items = ['기본급', '식대', '상여']
    calc_pay = sum([record.get(k, 0) for k in payment_items if isinstance(record.get(k), (int, float))])
    rec_pay_total = record.get('지급합계') or 0
    if calc_pay != rec_pay_total:
        errors.append(f"    - 지급합계 불일치: 계산된({calc_pay:,}) vs 추출된({rec_pay_total:,})")

    # 공제합계 검증
    deduction_items = ['국민연금', '건강보험', '고용보험', '장기요양보험료', '소득세', '지방소득세']
    calc_ded = sum([record.get(k, 0) for k in deduction_items if isinstance(record.get(k), (int, float))])
    rec_ded_total = record.get('공제합계') or 0
    if calc_ded != rec_ded_total:
        errors.append(f"    - 공제합계 불일치: 계산된({calc_ded:,}) vs 추출된({rec_ded_total:,})")

    # 실수령액 검증
    calc_net = rec_pay_total - rec_ded_total
    rec_net = record.get('차인지급액') or 0
    if calc_net != rec_net:
        errors.append(f"    - 차인지급액 불일치: 계산된({calc_net:,}) vs 추출된({rec_net:,})")

    if errors:
        print(f"Warning: {emp_name}님 데이터 검증 오류:")
        for msg in errors:
            print(msg)

def parse_payroll_data_from_raw_table(raw_df):
    processed = []
    start_row = 5
    block = 3
    total_rows = raw_df.shape[0]
    total_cols = raw_df.shape[1]
    for i in range(start_row, total_rows, block):
        if i + block > total_rows and (total_rows - i) < 2:
            break
        row1 = raw_df.iloc[i]
        row2 = raw_df.iloc[i + 1] if i + 1 < total_rows else pd.Series([None]*total_cols, index=raw_df.columns)
        row3 = raw_df.iloc[i + 2] if i + 2 < total_rows else pd.Series([None]*total_cols, index=raw_df.columns)

        rec = {}
        is_total = False
        temp_id = clean_value(row1.iloc[0])
        if isinstance(temp_id, str) and '합계' in temp_id:
            rec['구분'] = '합계'
            rec['성명'] = temp_id
            rec['사원번호'] = None
            rec['입사일'] = None
            is_total = True
        else:
            rec['구분'] = '직원'
            rec['사원번호'] = temp_id
            rec['성명'] = clean_value(row1.iloc[1])
            rec['입사일'] = clean_value(row2.iloc[0])

        if is_total:
            rec['기본급'] = clean_value(row1.iloc[1])
            rec['식대'] = clean_value(row2.iloc[0])
            rec['상여'] = None
            rec['국민연금'] = clean_value(row1.iloc[8])
            rec['건강보험'] = clean_value(row1.iloc[9])
            rec['고용보험'] = clean_value(row1.iloc[10])
            rec['장기요양보험료'] = clean_value(row1.iloc[11])
            rec['소득세'] = clean_value(row1.iloc[12])
            rec['지방소득세'] = clean_value(row1.iloc[13])
            rec['공제합계'] = clean_value(row2.iloc[12])
            rec['지급합계'] = clean_value(row3.iloc[6])
            rec['차인지급액'] = clean_value(row3.iloc[12])
        else:
            rec['기본급'] = clean_value(row1.iloc[2])
            rec['상여'] = clean_value(row1.iloc[3])
            rec['식대'] = clean_value(row2.iloc[2])
            rec['국민연금'] = clean_value(row1.iloc[9])
            rec['건강보험'] = clean_value(row1.iloc[10])
            rec['고용보험'] = clean_value(row1.iloc[11])
            rec['장기요양보험료'] = clean_value(row1.iloc[12])
            rec['소득세'] = clean_value(row1.iloc[13])
            rec['지방소득세'] = clean_value(row1.iloc[14])
            rec['공제합계'] = clean_value(row2.iloc[14])
            rec['지급합계'] = clean_value(row3.iloc[8])
            rec['차인지급액'] = clean_value(row3.iloc[14])
            verify_employee_totals(rec)

        processed.append(rec)
    return processed

def extract_and_process_payroll_with_tabula(pdf_path):
    try:
        tables = tabula.read_pdf(
            pdf_path,
            pages='1',
            lattice=True,
            encoding='CP949',                 # 한글 디코딩 오류 방지
            pandas_options={'header': None},
            multiple_tables=True
        )
        if not tables or len(tables) < 2:
            print("알림: PDF에서 충분한 테이블을 찾지 못했습니다. (최소 2개 예상)")
            return None
        raw_main_df = tables[1]
        return parse_payroll_data_from_raw_table(raw_main_df.copy())
    except Exception as e:
        msg = str(e)
        if ("java.io.IOException: Cannot run program \"java\": error=2" in msg
            or "JavaNotFoundError" in msg
            or "FileNotFoundError: [Errno 2] No such file or directory: 'java'" in msg):
            print("Error: Java가 설치되어 있지 않거나 Java 경로가 올바르지 않습니다.")
            print("       tabula-py를 사용하려면 Java가 필요합니다.")
        else:
            print(f"Error: 테이블 추출 및 처리 중 예상치 못한 오류 발생: {msg}")
        return None

def extract_payment_date(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]
        text = page.get_text("text")
        doc.close()
        match = re.search(r"\[지급\s*:\s*(\d{4}년\s?\d{1,2}월\s?\d{1,2}일)\]", text)
        if match:
            return match.group(1).replace(" ", "")
        else:
            print("알림: 지급일을 찾지 못했습니다.")
            return "지급일 정보 없음"
    except Exception as e:
        print(f"Error: 지급일 추출 중 오류 발생: {e}")
        return "지급일 정보 없음"


# ============================================================
# 3. FPDF 리소스 경로 헬퍼 & PayStubPDF 클래스
# ============================================================
def resource_path(relative_path):
    """ PyInstaller 번들 시 리소스 경로 반환 """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class PayStubPDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        font_file = resource_path('NanumGothic.ttf')
        try:
            if not os.path.exists(font_file):
                if 'google.colab' in sys.modules and os.path.exists('NanumGothic.ttf'):
                    font_file = 'NanumGothic.ttf'
                else:
                    raise RuntimeError(f"폰트 파일을 찾을 수 없습니다: {font_file}")
            self.add_font('NanumGothic', '', font_file)
            self.add_font('NanumGothic', 'B', font_file)
            self.font_family_regular = 'NanumGothic'
            self.font_family_bold = 'NanumGothic'
        except RuntimeError as e:
            print(f"FPDF 폰트 설정 오류: {e}.")
            self.font_family_regular = 'Arial'
            self.font_family_bold = 'Arial'

    def set_regular_font(self, size=10):
        self.set_font(self.font_family_regular, '', size)
    def set_bold_font(self, size=10):
        self.set_font(self.font_family_bold, 'B', size)

    def footer(self):
        self.set_y(-15)
        self.set_regular_font(8)
        self.cell(0, 10, "히어로 법무사사무소", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')

    def chapter_title(self, title):
        self.set_bold_font(16)
        self.cell(0, 10, title, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.ln(5)

    def employee_details(self, data, payment_date):
        self.set_regular_font(10)
        w = self.w - 2 * self.l_margin
        self.set_x(self.l_margin + w - 60)
        self.cell(60, 7, f"지급일 : {payment_date}", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='R')
        self.ln(2)
        col1, col2, lh = 35, 55, 7
        self.set_bold_font(10)
        self.cell(col1, lh, "성명", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_regular_font(10)
        self.cell(col2, lh, str(data.get('성명', '')), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_bold_font(10)
        self.cell(col1, lh, "생년월일(사번)", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_regular_font(10)
        self.cell(col2, lh, "", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.set_bold_font(10)
        self.cell(col1, lh, "부서", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_regular_font(10)
        self.cell(col2, lh, "", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_bold_font(10)
        self.cell(col1, lh, "직급", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_regular_font(10)
        self.cell(col2, lh, "", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.ln(5)

    def payment_details_table(self, data):
        self.set_bold_font(11)
        self.cell(0, 7, "세 부 내 역", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.ln(2)
        header = ["지 급 항 목", "금 액", "공 제 항 목", "금 액"]
        col_w = [55, 35, 55, 35]
        lh = 7
        self.set_bold_font(10)
        for i, text in enumerate(header):
            if i == len(header) - 1:
                self.cell(col_w[i], lh, text, border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
            else:
                self.cell(col_w[i], lh, text, border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.set_regular_font(9)

        payments = [
            ("기 본 급", data.get('기본급')),
            ("식    대", data.get('식대')),
            ("상    여", data.get('상여')),
            ("", None),
            ("", None)
        ]
        deductions = [
            ("국민 연금", data.get('국민연금')),
            ("건강 보험", data.get('건강보험')),
            ("고용 보험", data.get('고용보험')),
            ("장기요양 보험료", data.get('장기요양보험료')),
            ("소 득 세", data.get('소득세')),
            ("지방 소득세", data.get('지방소득세'))
        ]
        max_rows = max(len(payments), len(deductions))
        for i in range(max_rows):
            pay_item, pay_val = payments[i] if i < len(payments) else ("", None)
            pay_str = f"{pay_val:,}" if pay_val is not None else ""
            ded_item, ded_val = deductions[i] if i < len(deductions) else ("", None)
            ded_str = f"{ded_val:,}" if ded_val is not None else ""
            self.cell(col_w[0], lh, pay_item, border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='L')
            self.cell(col_w[1], lh, pay_str,   border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='R')
            self.cell(col_w[2], lh, ded_item,  border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='L')
            self.cell(col_w[3], lh, ded_str,   border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='R')
        self.set_bold_font(9)

        total_pay   = data.get('지급합계') or 0
        total_ded   = data.get('공제합계') or 0
        net_pay     = data.get('차인지급액') or 0
        tot_pay_str = f"{total_pay:,}"
        tot_ded_str = f"{total_ded:,}"
        net_pay_str = f"{net_pay:,}"

        self.cell(col_w[0], lh, "지급액 계",  border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.cell(col_w[1], lh, tot_pay_str, border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='R')
        self.cell(col_w[2], lh, "공제액 계",  border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.cell(col_w[3], lh, tot_ded_str, border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='R')
        self.cell(col_w[0] + col_w[1], lh, "실 수 령 액", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        self.cell(col_w[2] + col_w[3], lh, net_pay_str, border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='R')
        self.ln(5)

    def calculation_methods(self):
        self.set_bold_font(10)
        self.cell(0, 7, "계 산 방 법", border="B", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.ln(3)
        self.set_regular_font(8)
        methods = [
            "  · 근로소득세: 간이세액표 적용",
            "  · 지방소득세: 근로소득세 × 10%",
            "  · 국민연금: 취득신고 월 보수 × 4.5%",
            "  · 고용보험: 취득신고 월 보수 × 0.8%",
            "  · 건강보험: 취득신고 월 보수 × 3.43%",
            "  · 장기요양보험: 건강보험료 × 11.52%"
        ]
        for m in methods:
            self.multi_cell(0, 5, m, border=0, align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_regular_font(7)
        self.multi_cell(
            0, 5,
            "  ※ 해당 사업장 상황에 따라 기재가 필요없는 항목이 있을 수 있습니다.",
            border=0, align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT
        )
        self.ln(5)

    def work_days_hours(self):
        self.set_regular_font(8)
        lh = 6
        headers = ["근로일수", "총 근로시간수", "연장근로시간수", "야간근로시간수", "휴일근로시간수"]
        col_w = (self.w - 2 * self.l_margin) / len(headers)
        for i, head in enumerate(headers):
            if i == len(headers) - 1:
                self.cell(col_w, lh, head, border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
            else:
                self.cell(col_w, lh, head, border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
        for _ in headers:
            if _ == headers[-1]:
                self.cell(col_w, lh, "", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
            else:
                self.cell(col_w, lh, "", border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align='C')
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


# ============================================================
# 4. Tkinter UI 클래스 (PayrollApp)
# ============================================================
class PayrollApp:
    def __init__(self, master):
        self.master = master
        master.title("급여 명세서 자동 생성 프로그램 v0.3")
        master.geometry("550x300")

        self.input_pdf_path = ""
        self.output_dir = "generated_paystubs"

        input_frame = tk.Frame(master)
        input_frame.pack(pady=10, padx=10, fill=tk.X)

        self.btn_select_file = tk.Button(
            input_frame,
            text="1. 급여대장 PDF 선택",
            command=self.select_input_file,
            width=20
        )
        self.btn_select_file.pack(side=tk.LEFT, padx=5)

        self.label_file = tk.Label(
            input_frame,
            text="선택된 파일: 없음",
            anchor="w",
            justify=tk.LEFT
        )
        self.label_file.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_generate = tk.Button(
            master,
            text="2. 급여 명세서 생성 시작",
            command=self.generate_paystubs,
            state=tk.DISABLED,
            height=2,
            bg="lightblue"
        )
        self.btn_generate.pack(pady=10, padx=10, fill=tk.X)

        self.status_text = tk.Text(master, height=6, wrap=tk.WORD, state=tk.DISABLED)
        self.status_text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        self.btn_open_folder = tk.Button(
            master,
            text="저장 폴더 열기",
            command=self.open_output_folder,
            state=tk.DISABLED
        )
        self.btn_open_folder.pack(pady=5, padx=10)

    def _update_status(self, message, is_error=False):
        self.status_text.config(state=tk.NORMAL)
        if is_error:
            self.status_text.insert(tk.END, f"오류: {message}\n", "error")
            self.status_text.tag_config("error", foreground="red")
        else:
            self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
        self.master.update_idletasks()

    def select_input_file(self):
        filepath = filedialog.askopenfilename(
            title="급여대장 PDF 파일을 선택하세요",
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
        )
        if filepath:
            self.input_pdf_path = filepath
            self.label_file.config(text=f"선택: {os.path.basename(filepath)}")
            self.btn_generate.config(state=tk.NORMAL)
            self._update_status(f"'{os.path.basename(filepath)}' 파일이 선택되었습니다. '생성 시작'을 누르세요.")
        else:
            self.input_pdf_path = ""
            self.label_file.config(text="선택된 파일: 없음")
            self.btn_generate.config(state=tk.DISABLED)
            self._update_status("파일 선택이 취소되었습니다.")

    def generate_paystubs(self):
        if not self.input_pdf_path:
            messagebox.showerror("오류", "먼저 급여대장 PDF가 선택되어야 합니다.")
            self._update_status("급여대장 PDF 파일이 선택되지 않았습니다.", is_error=True)
            return

        self._update_status("급여 명세서 생성을 시작합니다... (잠시만요)")
        self.btn_generate.config(state=tk.DISABLED)
        self.btn_select_file.config(state=tk.DISABLED)

        try:
            payroll_data_list = extract_and_process_payroll_with_tabula(self.input_pdf_path)
            payment_date_on_ledger = extract_payment_date(self.input_pdf_path)

            if not payroll_data_list:
                messagebox.showerror("데이터 추출 오류", "PDF에서 급여 데이터를 추출하지 못했습니다.")
                self._update_status("급여 데이터 추출 실패.", is_error=True)
                self.btn_generate.config(state=tk.NORMAL)
                self.btn_select_file.config(state=tk.NORMAL)
                return

            if not os.path.exists(self.output_dir):
                os.makedirs(self.output_dir)

            self._update_status(f"'{self.output_dir}' 폴더에 파일을 저장합니다.")
            num_gen = 0
            gen_files = []

            for rec in payroll_data_list:
                if rec.get('구분') == '직원':
                    name_safe = str(rec.get('성명', '정보없음')).replace(" ", "_")
                    emp_id   = str(rec.get('사원번호', 'ID없음'))
                    out_fn   = os.path.join(self.output_dir, f"{name_safe}_{emp_id}_급여명세서.pdf")
                    pdf = PayStubPDF()
                    pdf.generate_paystub_pdf(rec, payment_date_on_ledger, out_fn)
                    gen_files.append(out_fn)
                    num_gen += 1

            if num_gen > 0:
                msg = f"{num_gen}명의 급여명세서가 생성되었습니다:\n" + "\n".join([f" - {os.path.basename(f)}" for f in gen_files])
                messagebox.showinfo("성공", msg.split('\n')[0])
                self._update_status(msg)
                self.btn_open_folder.config(state=tk.NORMAL)
            else:
                messagebox.showwarning("알림", "처리할 직원 데이터가 없습니다.")
                self._update_status("직원 데이터 없음.")
        except Exception as e:
            messagebox.showerror("치명적 오류", f"명세서 생성 중 예외 발생:\n{e}")
            self._update_status(f"예외 발생: {e}", is_error=True)
        finally:
            self.btn_generate.config(state=tk.NORMAL)
            self.btn_select_file.config(state=tk.NORMAL)

    def open_output_folder(self):
        abs_dir = os.path.abspath(self.output_dir)
        if os.path.exists(abs_dir):
            try:
                if os.name == 'nt':
                    os.startfile(abs_dir)
                else:
                    import subprocess
                    subprocess.call(['open' if sys.platform == 'darwin' else 'xdg-open', abs_dir])
                self._update_status(f"'{abs_dir}' 폴더를 열었습니다.")
            except Exception as e:
                self._update_status(f"폴더 열기 실패: {e}", is_error=True)
                messagebox.showerror("오류", f"폴더 열기 실패:\n{e}\n경로: {abs_dir}")
        else:
            messagebox.showwarning("알림", f"저장 폴더 '{abs_dir}'가 없습니다.")
            self._update_status(f"저장 폴더 없음: {abs_dir}", is_error=True)


# ============================================================
# 5. 메인 실행
# ============================================================
if __name__ == "__main__":
    # Colab 환경 여부 체크 (로컬에서는 tkinter GUI 실행)
    if 'google.colab' in sys.modules:
        print("Colab 환경 감지: UI 없이 PDF 처리 테스트 모드")
        font_file_colab = 'NanumGothic.ttf'
        payroll_pdf_colab = '25.05_히어로법무사사무소_급여대장.pdf'
        if not os.path.exists(font_file_colab):
            print(f"Error: '{font_file_colab}' 파일이 없습니다.")
        elif not os.path.exists(payroll_pdf_colab):
            print(f"Error: '{payroll_pdf_colab}' 파일이 없습니다.")
        else:
            out_dir = "generated_paystubs_colab"
            if not os.path.exists(out_dir):
                os.makedirs(out_dir)
            print(f"'{payroll_pdf_colab}' 데이터 추출 시작…")
            data_list = extract_and_process_payroll_with_tabula(payroll_pdf_colab)
            pay_date = extract_payment_date(payroll_pdf_colab)
            if data_list:
                print("\n--- 추출된 데이터 ---")
                for idx, r in enumerate(data_list):
                    print(f"레코드 {idx}: {r}")
                print(f"\n--- 지급일: {pay_date} ---")
                for r in data_list:
                    if r.get('구분') == '직원':
                        nm = str(r.get('성명','정보없음')).replace(" ","_")
                        eid = str(r.get('사원번호','ID없음'))
                        fn = os.path.join(out_dir, f"{nm}_{eid}_급여명세서.pdf")
                        pdf = PayStubPDF()
                        pdf.generate_paystub_pdf(r, pay_date, fn)
                print(f"\n파일 생성 완료: '{out_dir}' 폴더 확인")
            else:
                print("PDF 데이터 추출 실패.")
    else:
        # 로컬 환경: Tkinter GUI 실행
        root = tk.Tk()
        app = PayrollApp(root)
        root.mainloop()
