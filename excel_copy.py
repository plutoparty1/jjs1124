# -*- coding: utf-8 -*-
import os
import re
import shutil
import tempfile
import threading
from urllib.parse import unquote

import tkinter as tk
from tkinter import messagebox

from tkinterdnd2 import TkinterDnD, DND_FILES  # pip install tkinterdnd2
import win32com.client as win32               # pip install pywin32


# -------------------- 공통 유틸 -------------------- #

def extract_year_month_from_filename(filename: str) -> tuple[int, int, str, str] | None:
    """
    파일명에서 'YY년MM월', 'YYYY년MM월', 'YY.MM', 'YYYY.MM',
    'YYMM', 'YYYYMM' 패턴을 찾아서 (year_full, month, 사용한_정규식패턴, 매칭된_전체문자열) 형태로 반환.
    year_full은 4자리 정수(예: 2025).
    """
    name, _ = os.path.splitext(os.path.basename(filename))
    patterns = [
        r'(?P<year>\d{4})\s*년\s*(?P<month>\d{1,2})\s*월',
        r'(?P<year>\d{2})\s*년\s*(?P<month>\d{1,2})\s*월',
        r'(?P<year>\d{4})\.\s*(?P<month>\d{1,2})(?=\D|$)',
        r'(?P<year>\d{2})\.\s*(?P<month>\d{1,2})(?=\D|$)',
        r'(?P<year>\d{4})(?P<month>\d{1,2})(?=\D|$)',
        r'(?P<year>\d{2})(?P<month>\d{1,2})(?=\D|$)',
    ]
    for pattern in patterns:
        m = re.search(pattern, name)
        if m:
            year_str = m.group("year")
            month_str = m.group("month")
            year = int(year_str)
            month = int(month_str)
            if len(year_str) == 2:
                year += 2000
            matched_text = m.group(0)  # 매칭된 전체 문자열
            return year, month, pattern, matched_text
    return None


def get_next_year_month(year: int, month: int) -> tuple[int, int]:
    """year, month에서 다음 달의 (year, month)를 반환."""
    if month == 12:
        return year + 1, 1
    else:
        return year, month + 1




def make_unique_path(path: str) -> str:
    """
    같은 이름의 파일이 이미 있으면 _copy1, _copy2... 붙이면서
    겹치지 않는 경로를 찾아서 반환.
    """
    base, ext = os.path.splitext(path)
    candidate = path
    i = 1
    while os.path.exists(candidate):
        candidate = f"{base}_copy{i}{ext}"
        i += 1
    return candidate


# -------------------- xls → xlsx 변환 (Excel COM 객체 재사용) -------------------- #

def normalize_path(path: str) -> str:
    """경로 정규화: URL 디코딩 및 경로 정리."""
    # URL 인코딩된 경로 디코딩
    try:
        path = unquote(path)
    except Exception:
        pass
    
    # 경로 정규화 (슬래시 통일, 중복 제거 등)
    path = os.path.normpath(path)
    
    # 절대 경로로 변환
    if not os.path.isabs(path):
        path = os.path.abspath(path)
    
    return path


def convert_xls_to_xlsx_with_excel(xls_path: str, dest_xlsx_path: str, excel_app=None) -> str:
    """
    Excel COM 객체를 사용하여 .xls를 .xlsx로 변환.
    excel_app이 제공되면 재사용하고, None이면 새로 생성 (호출자가 종료해야 함).
    """
    # 경로 정규화
    xls_path = normalize_path(xls_path)
    dest_xlsx_path = normalize_path(dest_xlsx_path)
    dest_xlsx_path = make_unique_path(dest_xlsx_path)
    
    # Excel 객체가 없으면 새로 생성 (호출자가 종료해야 함)
    excel = excel_app
    should_quit = False
    if excel is None:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False  # 화면 업데이트 비활성화로 성능 향상
        excel.EnableEvents = False  # 이벤트 비활성화로 성능 향상
        # Calculation 속성은 일부 Excel 버전에서 지원하지 않을 수 있음
        try:
            excel.Calculation = -4105  # xlCalculationManual - 자동 계산 비활성화
        except Exception:
            pass  # 실패해도 계속 진행
        should_quit = True
    
    wb = None
    try:
        # 파일 경로를 절대 경로로 변환하고 백슬래시 사용
        xls_path_abs = os.path.abspath(xls_path)
        dest_xlsx_path_abs = os.path.abspath(dest_xlsx_path)
        
        # Excel은 백슬래시를 선호하므로 변환
        xls_path_excel = xls_path_abs.replace('/', '\\')
        dest_xlsx_path_excel = dest_xlsx_path_abs.replace('/', '\\')
        
        # 직접 원본 파일을 열어서 목적지로 저장
        wb = excel.Workbooks.Open(xls_path_excel, ReadOnly=True, UpdateLinks=0)
        wb.SaveAs(dest_xlsx_path_excel, FileFormat=51)  # 51 = xlsx
        wb.Close(False)
        wb = None
    except Exception as e:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if should_quit:
            try:
                excel.Quit()
            except Exception:
                pass
        raise RuntimeError(f".xls → .xlsx 변환 중 오류: {e}")
    
    return dest_xlsx_path


# -------------------- 메인 로직: 파일명만 변경 (내용 변경 없음) -------------------- #

def make_next_month_copy(original_path: str, excel_app=None) -> tuple[str, int]:
    """
    - 파일명에서 연/월 추출 ('YY년MM월', 'YYYY년MM월', 'YY.MM', 'YYYY.MM' 형식 지원)
    - 다음달 연/월로 바뀐 이름을 만들고 (확장자는 .xlsx로 통일)
    - 원본이 .xls면: Excel COM으로 .xlsx로 변환 후, 새 이름으로 저장
    - 원본이 .xlsx/.xlsm이면: 파일을 새 이름(.xlsx/.xlsm)으로 복사
    - ⚠️ 엑셀 파일 안의 내용(셀 값, 날짜 등)은 전혀 수정하지 않음
    - excel_app: Excel COM 객체 (재사용 시 제공, None이면 필요시 새로 생성)
    → (최종 사본 경로, 변경된 셀 개수=0) 반환
    """
    info = extract_year_month_from_filename(original_path)
    if not info:
        raise ValueError("파일명에서 'YY년MM월' / 'YYYY년MM월' / 'YY.MM' / 'YYYY.MM' 패턴을 찾지 못함.")

    old_year, old_month, pattern, matched_text = info
    new_year, new_month = get_next_year_month(old_year, old_month)

    dir_name = os.path.dirname(original_path)
    base_name = os.path.basename(original_path)
    name, ext = os.path.splitext(base_name)

    # 원본 매칭 문자열에서 공백 패턴 추출
    # 예: "25년11월" -> 공백 없음, "25년 11월" -> "년" 뒤에 공백 있음
    space_after_year = ""
    if "년" in matched_text:
        # "년" 뒤 공백 추출
        idx_year = matched_text.find("년")
        if idx_year + 1 < len(matched_text):
            # "년" 뒤부터 숫자 시작 전까지의 공백 추출
            after_year = matched_text[idx_year + 1:]
            # 숫자가 시작되기 전까지의 공백만 추출
            space_match = re.match(r'^(\s+)', after_year)
            if space_match:
                space_after_year = space_match.group(1)
    
    # 새 파일 이름 (원본 확장자 유지, 단 .xls는 .xlsx로 변환)
    def repl_func(m):
        year_str = m.group("year")
        month_str = m.group("month")
        
        # 원본 연도 형식 유지
        if len(year_str) == 4:
            year_new_str = f"{new_year}"
        else:
            year_new_str = f"{new_year % 100:02d}"
        
        # 원본 월 형식 유지 (한 자리면 한 자리, 두 자리면 두 자리)
        if len(month_str) == 1:
            month_new_str = f"{new_month}"
        else:
            month_new_str = f"{new_month:02d}"
        
        # 패턴에 따라 형식 결정 (원본 공백 유지)
        if "년" in pattern and "월" in pattern:
            return f"{year_new_str}년{space_after_year}{month_new_str}월"
        elif "." in pattern:
            return f"{year_new_str}.{month_new_str}"
        else:
            return f"{year_new_str}{month_new_str}"

    new_name = re.sub(pattern, repl_func, name, count=1)
    ext_lower = ext.lower()
    
    # 확장자 결정: .xls는 .xlsx로, 나머지는 원본 확장자 유지
    if ext_lower == ".xls":
        new_ext = ".xlsx"
    else:
        new_ext = ext  # .xlsx 또는 .xlsm 유지
    
    new_filename = new_name + new_ext
    dest_path = os.path.join(dir_name, new_filename)
    dest_path = make_unique_path(dest_path)

    if ext_lower == ".xls":
        # xls → xlsx 변환 (Excel 객체 재사용)
        final_path = convert_xls_to_xlsx_with_excel(original_path, dest_path, excel_app)
    elif ext_lower in [".xlsx", ".xlsm"]:
        # 파일 복사 최적화: 메타데이터가 중요하지 않으므로 copy 사용 (copy2보다 빠름)
        shutil.copy(original_path, dest_path)
        final_path = dest_path
    else:
        raise ValueError(".xls / .xlsx / .xlsm 파일만 지원.")

    # 내용은 안 바꾸니까 항상 0
    changed_count = 0

    return final_path, changed_count


# -------------------- GUI (드래그앤드롭 창) -------------------- #

class ExcelDnDApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("엑셀 사본 생성")
        self.geometry("600x500")
        self.attributes("-topmost", True)  # 항상 위에
        self.configure(bg="#f0f0f0")

        # 안내 라벨
        self.label = tk.Label(
            self,
            text="(.xlsx / .xls / .xlsm)",
            bg="#f0f0f0",
            font=("맑은 고딕", 14)
        )
        self.label.pack(fill="x", pady=(10, 5))

        # 로그 영역
        self.log = tk.Text(
            self,
            height=20,
            state="disabled",
            bg="#ffffff",
            font=("Consolas", 10),
        )
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # DnD 설정
        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self.on_drop)
        
        # 처리 중 플래그 (동시 처리 방지용, 하지만 완료 후에는 다시 처리 가능)
        # 메인 스레드에서만 수정하도록 주의
        self._processing = False

    def append_log(self, msg: str):
        """아래 텍스트 박스에 로그 추가 (스레드 안전)."""
        self.after(0, self._append_log_sync, msg)
    
    def _append_log_sync(self, msg: str):
        """로그 추가 (메인 스레드에서 실행)."""
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update()  # GUI 업데이트 강제
    
    def _set_processing_flag_and_log(self, value: bool, msg: str = None):
        """처리 플래그 설정 및 로그 (메인 스레드에서 실행)."""
        self._processing = value
        if msg:
            self._append_log_sync(msg)

    def on_drop(self, event):
        """드롭 이벤트 핸들러 - 백그라운드 스레드에서 처리."""
        # 처리 중이면 무시 (동시 처리 방지)
        if self._processing:
            self.append_log("⚠️ 이미 처리 중입니다. 완료 후 다시 시도해주세요.")
            return
        
        files = self.splitlist(event.data)
        if not files:
            return
        
        # 플래그를 먼저 설정 (메인 스레드에서 동기적으로)
        self._processing = True
        
        # 백그라운드 스레드에서 처리하여 GUI가 멈추지 않도록
        thread = threading.Thread(target=self._process_files, args=(files,), daemon=True)
        thread.start()

    def _process_files(self, files):
        """파일 처리 (백그라운드 스레드에서 실행)."""
        import pythoncom
        excel_app = None
        com_initialized = False
        
        try:
            # 백그라운드 스레드에서 COM 초기화
            pythoncom.CoInitialize()
            com_initialized = True
            
            self.append_log("=" * 60)
            self.append_log("드롭 처리 시작")

            # 파일 목록 정리 및 검증
            valid_files = []
            skipped = []
            
            for file_path in files:
                file_path = file_path.strip()
                if not file_path:
                    continue

                # 경로 정규화 (URL 디코딩 등)
                try:
                    file_path = normalize_path(file_path)
                except Exception:
                    pass  # 정규화 실패해도 원본 경로 사용

                if not os.path.isfile(file_path):
                    skipped.append(f"(파일 아님) {file_path}")
                    continue

                ext = os.path.splitext(file_path)[1].lower()
                if ext not in [".xlsx", ".xls", ".xlsm"]:
                    skipped.append(f"(엑셀 아님) {os.path.basename(file_path)}")
                    continue
                
                valid_files.append(file_path)

            # .xls 파일이 있는지 확인하여 Excel 객체 미리 생성
            has_xls = any(os.path.splitext(f)[1].lower() == ".xls" for f in valid_files)
            
            if has_xls:
                try:
                    self.append_log("Excel 객체 생성 중...")
                    excel_app = win32.DispatchEx("Excel.Application")
                    excel_app.Visible = False
                    excel_app.DisplayAlerts = False
                    excel_app.ScreenUpdating = False  # 화면 업데이트 비활성화
                    excel_app.EnableEvents = False  # 이벤트 비활성화
                    # Calculation 속성은 일부 Excel 버전에서 지원하지 않을 수 있음
                    try:
                        excel_app.Calculation = -4105  # xlCalculationManual - 자동 계산 비활성화
                    except Exception:
                        pass  # 실패해도 계속 진행
                    self.append_log("Excel 객체 생성 완료 (재사용 모드)")
                except Exception as e:
                    self.append_log(f"⚠️ Excel 객체 생성 실패: {e}")

            success_files = []
            failed = []

            # 각 파일 처리
            for i, file_path in enumerate(valid_files, 1):
                file_name = os.path.basename(file_path)
                self.append_log(f"[{i}/{len(valid_files)}] 처리 중: {file_name}")
                
                try:
                    dest, _ = make_next_month_copy(file_path, excel_app)
                    success_files.append(dest)
                    self.append_log(f"  ✅ 완료: {os.path.basename(dest)}")
                except Exception as e:
                    failed.append(f"{file_name} → {e}")
                    self.append_log(f"  ❌ 실패: {e}")

            # 최종 로그 출력
            self.append_log("")
            if success_files:
                self.append_log("✅ 복사 완료된 파일:")
                for p in success_files:
                    self.append_log(f"  - {p}")

            if skipped:
                self.append_log("⏭ 건너뛴 항목:")
                for s in skipped:
                    self.append_log(f"  - {s}")

            if failed:
                self.append_log("❌ 처리 실패한 파일:")
                for f in failed:
                    self.append_log(f"  - {f}")

            self.append_log("드롭 처리 종료")
            self.append_log("")  # 빈 줄

        except Exception as e:
            # 여기서도 팝업 대신 로그만 찍어줌
            self.append_log(f"❌ 전체 처리 중 예외 발생: {e}")
            import traceback
            self.append_log(traceback.format_exc())
        finally:
            # Excel 객체 정리
            if excel_app is not None:
                try:
                    self.append_log("Excel 객체 종료 중...")
                    excel_app.Quit()
                    self.append_log("Excel 객체 종료 완료")
                except Exception as e:
                    self.append_log(f"⚠️ Excel 객체 종료 중 오류: {e}")
            
            # COM 정리
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            
            # 준비 완료 메시지 먼저 표시
            self.append_log("✅ 준비 완료 - 다음 파일을 드롭할 수 있습니다.")
            
            # 처리 완료 플래그 해제 (메인 스레드에서 실행, 약간의 지연을 두어 로그 메시지 표시 후 해제)
            # 200ms 지연으로 로그 메시지가 표시된 후 플래그 해제
            self.after(200, self._set_processing_flag_and_log, False, None)


if __name__ == "__main__":
    app = ExcelDnDApp()
    app.mainloop()
