import re
import os
import json
import shutil
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Set, Optional, Tuple, Callable

import pandas as pd

# 설정 파일 경로
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor_configs.json")


# ----------------------------
# 0) Excel 엔진 유틸
# ----------------------------
def get_excel_engine(file_path: str) -> str:
    """파일 확장자에 따라 pandas engine 반환"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".xls":
        return "xlrd"
    return "openpyxl"


# ----------------------------
# 1) 업체별 설정 (여기만 늘리면 됨)
# ----------------------------
@dataclass
class VendorConfig:
    name: str

    # 전체리스트 쪽
    list_sheet: Optional[str] = None      # None이면 active 시트 사용
    header_row: int = 3                   # "전체리스트" 파일 헤더가 있는 행
    list_id_col: str = "D"                # 전체리스트에서 로그인ID가 있는 열

    # 거래명세서(업체 양식) 쪽
    invoice_sheet: str = "상세내역"        # 업체 양식에서 상세내역 시트명 (단일)
    invoice_sheets: Optional[List[str]] = None  # 여러 시트에 추가할 경우 (예: ["상세내역-BGM", "상세내역-DMB"])
    store_col_letter: str = "C"           # 업체 양식에서 매장명이 있는 컬럼(예: C열)
    table_header_text: str = "매장명"     # 테이블 헤더에서 찾을 텍스트 (동적 감지용)
    
    # ID 시트 설정 (매장명 ↔ 로그인ID 매핑 테이블)
    id_sheet: Optional[str] = None        # 거래명세서에서 ID 시트명 (예: "ID")
    id_list_store_col: str = "A"          # ID 시트에서 전체리스트 매장명 열 (주스샵 매장명)
    id_store_col: str = "B"               # ID 시트에서 명세서 매장명 열
    id_login_col: str = "C"               # ID 시트에서 로그인ID 열
    id_start_row: int = 2                 # ID 시트에서 데이터 시작 행
    
    # 보호 테이블 헤더 (이 텍스트를 찾아서 그 행 위에 삽입)
    protected_table_headers: Optional[List[str]] = None  # 예: ["공급가액", "부가세"]

    # 업체별 필터(전체리스트에서 기업명 매칭)
    company_value: Optional[str] = None   # 예: "할리스커피" / 업체마다 다름
    group_col: Optional[str] = None       # 그룹명 열 (예: "B")
    group_value: Optional[str] = None     # 그룹명 값 (예: "타코벨") - 포함 조건
    group_exclude: Optional[List[str]] = None  # 제외할 그룹명 목록 (예: ["타코벨"])
    
    # 날짜 셀 (첫번째 시트에 오늘 날짜 입력)
    date_cell: Optional[str] = None       # 날짜 셀 (예: "A1", "B2") - 형식: YYYY-MM-DD


VENDOR_CONFIGS: Dict[str, VendorConfig] = {
    "할리스커피": VendorConfig(
        name="할리스커피",
        list_sheet=None,
        header_row=3,
        list_id_col="D",            # 전체리스트 D열 = 로그인ID
        invoice_sheet="상세내역",
        store_col_letter="B",       # 상세내역 B열 = 매장명
        table_header_text="매장명",  # 테이블 헤더 텍스트 (동적 감지)
        id_sheet="ID",              # ID 시트 (매핑 테이블)
        id_store_col="B",           # ID 시트 B열 = 매장명
        id_login_col="C",           # ID 시트 C열 = 로그인ID
        id_start_row=2,             # ID 시트 2행부터 데이터
        protected_table_headers=["공급가액", "부가세"],
        company_value="할리스커피",
    ),
    "카페드롭탑": VendorConfig(
        name="카페드롭탑",
        list_sheet=None,
        header_row=3,
        list_id_col="D",            # 전체리스트 D열 = 로그인ID
        invoice_sheet="상세내역",
        store_col_letter="E",       # E열 = 매장명
        table_header_text="매장명",  # 테이블 헤더 텍스트 (동적 감지)
        id_sheet="ID",              # ID 시트 (매핑 테이블)
        id_store_col="B",           # ID 시트 B열 = 매장명
        id_login_col="C",           # ID 시트 C열 = 로그인ID
        id_start_row=2,             # ID 시트 2행부터 데이터
        protected_table_headers=["공급가액", "부가세"],
        company_value="카페드롭탑",
    ),
    "아웃백스테이크하우스": VendorConfig(
        name="아웃백스테이크하우스",
        list_sheet=None,
        header_row=3,
        list_id_col="D",            # 전체리스트 D열 = 로그인ID
        invoice_sheet="상세내역",
        store_col_letter="C",       # C열 = 매장명
        table_header_text="매장명",  # 테이블 헤더 텍스트 (동적 감지)
        id_sheet="ID",              # ID 시트 (매핑 테이블)
        id_store_col="B",           # ID 시트 B열 = 매장명
        id_login_col="C",           # ID 시트 C열 = 로그인ID
        id_start_row=2,             # ID 시트 2행부터 데이터
        protected_table_headers=["공급가액", "부가세"],
        company_value="아웃백스테이크하우스",
    ),
    "타코벨": VendorConfig(
        name="타코벨",
        list_sheet=None,
        header_row=3,
        list_id_col="D",            # 전체리스트 D열 = 로그인ID
        invoice_sheets=["상세내역-BGM", "상세내역-DMB"],  # 두 시트에 추가
        store_col_letter="C",       # C열 = 매장명
        table_header_text="매장명",  # 테이블 헤더 텍스트 (동적 감지)
        id_sheet="ID",              # ID 시트 (매핑 테이블)
        id_store_col="B",           # ID 시트 B열 = 매장명
        id_login_col="C",           # ID 시트 C열 = 로그인ID
        id_start_row=2,             # ID 시트 2행부터 데이터
        protected_table_headers=["공급가액", "부가세"],
        company_value="KFC",        # 기업명(A열) = KFC
        group_col="B",              # 그룹명 열 = B
        group_value="타코벨",        # 그룹명(B열) = 타코벨
    ),
    "KFC": VendorConfig(
        name="KFC",
        list_sheet=None,
        header_row=3,
        list_id_col="D",            # 전체리스트 D열 = 로그인ID
        invoice_sheet="상세내역",
        store_col_letter="C",       # C열 = 매장명
        table_header_text="매장명",  # 테이블 헤더 텍스트 (동적 감지)
        id_sheet="ID",              # ID 시트 (매핑 테이블)
        id_store_col="B",           # ID 시트 B열 = 매장명
        id_login_col="C",           # ID 시트 C열 = 로그인ID
        id_start_row=2,             # ID 시트 2행부터 데이터
        protected_table_headers=["공급가액", "부가세"],
        company_value="KFC",        # 기업명(A열) = KFC
        group_col="B",              # 그룹명 열 = B
        group_exclude=["타코벨"],    # 그룹명 타코벨 제외
    ),
    # 여기에 업체 계속 추가하면 됨
}


def save_vendor_configs():
    """업체 설정을 JSON 파일로 저장"""
    data = {}
    for name, config in VENDOR_CONFIGS.items():
        config_dict = asdict(config)
        data[name] = config_dict
    
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_vendor_configs():
    """JSON 파일에서 업체 설정 불러오기"""
    global VENDOR_CONFIGS
    
    if not os.path.exists(CONFIG_FILE):
        # 파일이 없으면 기본 설정 저장
        save_vendor_configs()
        return
    
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        VENDOR_CONFIGS.clear()
        for name, config_dict in data.items():
            VENDOR_CONFIGS[name] = VendorConfig(**config_dict)
    except Exception as e:
        print(f"설정 파일 로드 실패: {e}")


def add_vendor_config(config: VendorConfig):
    """업체 설정 추가"""
    VENDOR_CONFIGS[config.name] = config
    save_vendor_configs()


def delete_vendor_config(name: str):
    """업체 설정 삭제"""
    if name in VENDOR_CONFIGS:
        del VENDOR_CONFIGS[name]
        save_vendor_configs()


# 프로그램 시작 시 설정 로드
load_vendor_configs()


# ----------------------------
# 2) 공통 유틸
# ----------------------------
TEST_PAT = re.compile(r"(테스트|test)", re.IGNORECASE)

# 매장명 정규화: "점" 접미사 등 제거
STORE_SUFFIX_PAT = re.compile(r"(점|지점|매장|센터|스토어|store)$", re.IGNORECASE)

def norm_text(x) -> str:
    if x is None:
        return ""
    return str(x).strip()

def normalize_store_name(name: str) -> str:
    """
    매장명 정규화 - 비교용
    예: "(하)선산휴게소점" → "(하)선산휴게소"
    """
    name = norm_text(name)
    # 접미사 제거
    name = STORE_SUFFIX_PAT.sub("", name)
    return name

def is_test_account(a_to_d_values: List[object]) -> bool:
    joined = " ".join(norm_text(v) for v in a_to_d_values if norm_text(v))
    return bool(TEST_PAT.search(joined))

def find_col_idx_by_header(headers: List[str], header_name: str) -> int:
    # pandas의 df.columns는 보통 문자열/NaN 섞일 수 있음
    for i, h in enumerate(headers):
        if norm_text(h) == header_name:
            return i
    raise KeyError(f"헤더 '{header_name}' 컬럼을 못 찾았어. 실제 헤더명을 확인해봐.")


# ----------------------------
# 3) 전체리스트에서 "추가해야 할 매장" 뽑기 (로그인ID → 매장명 딕셔너리)
# ----------------------------
def extract_stores_from_list(
    list_path: str,
    vendor: VendorConfig,
) -> Dict[str, str]:
    """
    Returns: {로그인ID: 매장명} 딕셔너리
    """
    engine = get_excel_engine(list_path)
    df = pd.read_excel(
        list_path,
        sheet_name=vendor.list_sheet if vendor.list_sheet else 0,
        header=vendor.header_row - 1,
        engine=engine,
        dtype=object,
    )

    # 필요한 컬럼 인덱스 찾기
    headers = [norm_text(h) for h in df.columns.tolist()]

    company_idx = find_col_idx_by_header(headers, "기업명")
    store_idx = find_col_idx_by_header(headers, "매장명")
    recent_login_idx = find_col_idx_by_header(headers, "최근로그인시간")

    # 그룹명 열 인덱스 (있는 경우)
    group_idx = None
    if vendor.group_col:
        group_idx = col_letter_to_num(vendor.group_col) - 1  # 0-based index
    
    # 로그인ID 열 인덱스 (D열 = 인덱스 3)
    id_col_idx = col_letter_to_num(vendor.list_id_col) - 1  # 0-based index

    # {로그인ID: 매장명}
    id_to_store: Dict[str, str] = {}

    for _, row in df.iterrows():
        company = norm_text(row.iloc[company_idx])
        store = norm_text(row.iloc[store_idx])
        recent_login = norm_text(row.iloc[recent_login_idx])
        login_id = norm_text(row.iloc[id_col_idx])

        # 업체별: 기업명 매칭
        if vendor.company_value and company != vendor.company_value:
            continue
        
        # 업체별: 그룹명 매칭 (있는 경우)
        if vendor.group_value and group_idx is not None:
            group = norm_text(row.iloc[group_idx])
            if group != vendor.group_value:
                continue
        
        # 업체별: 그룹명 제외 (있는 경우)
        if vendor.group_exclude and group_idx is not None:
            group = norm_text(row.iloc[group_idx])
            if group in vendor.group_exclude:
                continue

        # 공통: 최근로그인시간 비어있으면 제외
        if recent_login == "":
            continue

        # 공통: 테스트 계정 제외 (A~D 검사)
        a_to_d = row.iloc[0:4].tolist()
        if is_test_account(a_to_d):
            continue

        if login_id and store:
            id_to_store[login_id] = store

    return id_to_store


def col_letter_to_num(letter: str) -> int:
    """A=1, B=2, ..., Z=26, AA=27, ..."""
    result = 0
    for char in letter.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


# ----------------------------
# 4) Excel COM을 사용한 거래명세서 처리 (이미지 보존)
# ----------------------------
def find_id_sheet(wb, vendor: VendorConfig) -> Tuple[object, bool]:
    """
    ID 시트 찾기 (숨겨져 있어도 찾음)
    Returns: (시트 객체, 원래 숨김 상태였는지 여부)
    """
    if not vendor.id_sheet:
        return None, False
    
    id_ws = None
    was_hidden = False
    
    for sheet in wb.Sheets:
        if sheet.Name == vendor.id_sheet:
            id_ws = sheet
            # 숨겨져 있으면 임시로 보이게
            if not sheet.Visible:
                was_hidden = True
                sheet.Visible = True
            break
    
    return id_ws, was_hidden


def read_id_sheet_mapping(wb, vendor: VendorConfig) -> Tuple[Dict[str, str], Dict[str, str]]:
    """
    ID 시트에서 매핑 읽기
    Returns: ({매장명: 로그인ID}, {로그인ID: 매장명})
    """
    if not vendor.id_sheet:
        return {}, {}
    
    # ID 시트 찾기 (숨겨져 있어도 찾음)
    id_ws, _ = find_id_sheet(wb, vendor)
    
    if id_ws is None:
        return {}, {}
    
    store_col = col_letter_to_num(vendor.id_store_col)
    login_col = col_letter_to_num(vendor.id_login_col)
    start = vendor.id_start_row
    
    # 벌크로 읽기 (B열과 C열)
    end_row = start + 2000
    range_str = f"{vendor.id_store_col}{start}:{vendor.id_login_col}{end_row}"
    values = id_ws.Range(range_str).Value
    
    store_to_id: Dict[str, str] = {}
    id_to_store_invoice: Dict[str, str] = {}
    
    if values:
        for row in values:
            if not isinstance(row, tuple) or len(row) < 2:
                continue
            store_name = norm_text(row[0])  # B열 (명세서 매장명)
            login_id = norm_text(row[1])    # C열 (로그인ID)
            if store_name == "" or login_id == "":
                continue
            store_to_id[store_name] = login_id
            id_to_store_invoice[login_id] = store_name
    
    return store_to_id, id_to_store_invoice


def get_existing_login_ids_dynamic(
    ws, vendor: VendorConfig, store_to_id: Dict[str, str], 
    data_start_row: int, protected_row: Optional[int]
) -> Tuple[Set[str], List[str]]:
    """
    상세내역 시트의 기존 매장명들을 ID 시트 매핑으로 로그인ID로 변환 (동적 레이아웃)
    Returns: (기존 로그인ID set, 매핑되지 않은 매장명 리스트)
    """
    col_num = col_letter_to_num(vendor.store_col_letter)
    start = data_start_row
    
    # 보호 행까지만 읽기
    end_row = protected_row - 1 if protected_row else start + 1000
    range_str = f"{vendor.store_col_letter}{start}:{vendor.store_col_letter}{end_row}"
    values = ws.Range(range_str).Value
    
    existing_ids: Set[str] = set()
    existing_store_names: List[str] = []  # 기존 매장명 목록
    
    if values:
        for row in values:
            val = row[0] if isinstance(row, tuple) else row
            store_name = norm_text(val)
            if store_name == "":
                break
            existing_store_names.append(store_name)
            # 매장명으로 로그인ID 찾기
            login_id = store_to_id.get(store_name, "")
            if login_id:
                existing_ids.add(login_id)
    
    return existing_ids, existing_store_names


def add_to_id_sheet(wb, vendor: VendorConfig, new_ids: List[str], list_store_names: List[str]):
    """
    ID 시트에 새 매장 추가
    - 명세서 매장명(id_store_col): 비워둠
    - 로그인ID(id_login_col): 입력
    - 전체리스트 매장명(id_list_store_col): 입력 (주스샵 매장명)
    """
    if not vendor.id_sheet or not new_ids:
        return
    
    # ID 시트 찾기 (숨겨져 있어도 찾음)
    id_ws, _ = find_id_sheet(wb, vendor)
    
    if id_ws is None:
        return
    
    store_col = col_letter_to_num(vendor.id_store_col)        # 명세서 매장명 (비워둠)
    login_col = col_letter_to_num(vendor.id_login_col)        # 로그인ID
    list_store_col = col_letter_to_num(vendor.id_list_store_col)  # 전체리스트 매장명
    
    # 마지막 데이터 행 찾기 (로그인ID 열 기준으로 찾기)
    start = vendor.id_start_row
    r = start
    while True:
        val = id_ws.Cells(r, login_col).Value
        if norm_text(val) == "":
            break
        r += 1
    
    # 새 매장 추가
    for i, (login_id, list_store) in enumerate(zip(new_ids, list_store_names)):
        # 전체리스트 매장명 입력 (주스샵 매장명) - A열
        id_ws.Cells(r + i, list_store_col).Value = list_store
        # 명세서 매장명은 비워둠 (id_store_col) - B열
        id_ws.Cells(r + i, store_col).Value = ""
        # 로그인ID 입력 (텍스트 형식) - C열
        login_cell = id_ws.Cells(r + i, login_col)
        login_cell.NumberFormat = "@"  # 텍스트 형식
        login_cell.Value = str(login_id)


def hide_id_sheet(wb, vendor: VendorConfig):
    """
    ID 시트 숨기기
    """
    if not vendor.id_sheet:
        return
    
    for sheet in wb.Sheets:
        if sheet.Name == vendor.id_sheet:
            sheet.Visible = False
            break


def find_supply_amount_cell(ws, vendor: VendorConfig, start_row: int) -> Optional[Tuple[int, int]]:
    """
    공급가액 셀의 위치를 찾기
    Returns: (행, 열) 또는 None
    """
    if not vendor.protected_table_headers:
        return None
    
    # 첫 번째 헤더 텍스트 (보통 "공급가액")
    search_text = vendor.protected_table_headers[0]
    
    # 검색 범위: start_row부터 충분히 큰 범위
    max_search = start_row + 2000
    
    # 사용 범위 확인
    used_range = ws.UsedRange
    last_used_row = used_range.Row + used_range.Rows.Count - 1
    max_search = min(max_search, last_used_row + 1)
    
    # 여러 열에서 헤더 검색
    used_cols = used_range.Column + used_range.Columns.Count - 1
    search_cols = list(range(1, min(used_cols + 1, 15)))  # 1~14열에서 검색
    
    for r in range(start_row, max_search):
        for col in search_cols:
            cell_value = ws.Cells(r, col).Value
            text = norm_text(cell_value)
            if search_text in text:
                return (r, col)
    
    return None


def write_excluded_stores_list(
    ws, vendor: VendorConfig, excluded_stores: List[str], 
    supply_cell_row: int, supply_cell_col: int
):
    """
    제외된 매장 목록을 공급가액 셀 3칸 아래에 작성
    """
    if not excluded_stores:
        return
    
    # 공급가액 셀의 3칸 아래부터 시작
    start_row = supply_cell_row + 3
    
    # 헤더 작성
    ws.Cells(start_row, supply_cell_col).Value = "※ 제외된 매장 (전체리스트에 없음)"
    ws.Cells(start_row, supply_cell_col).Font.Bold = True
    ws.Cells(start_row, supply_cell_col).Font.Color = 0x0000FF  # 빨간색 (BGR 형식)
    
    # 제외된 매장 목록 작성
    for i, store_name in enumerate(sorted(excluded_stores)):
        ws.Cells(start_row + 1 + i, supply_cell_col).Value = f"- {store_name}"
        ws.Cells(start_row + 1 + i, supply_cell_col).Font.Color = 0x0000FF  # 빨간색


def detect_table_layout(ws, vendor: VendorConfig) -> Tuple[int, int, int]:
    """
    테이블 레이아웃 동적 감지
    - 헤더 행 찾기 (table_header_text로 검색)
    - 테이블 너비 감지
    Returns: (데이터 시작 행, 테이블 시작 열, 테이블 끝 열)
    """
    store_col = col_letter_to_num(vendor.store_col_letter)
    header_text = vendor.table_header_text
    
    # 1) 헤더 행 찾기 (store_col에서 header_text 검색)
    header_row = None
    for r in range(1, 100):  # 1~100행에서 검색
        cell_value = ws.Cells(r, store_col).Value
        text = norm_text(cell_value)
        if header_text in text:
            header_row = r
            break
    
    if header_row is None:
        # 못 찾으면 기본값 사용
        header_row = 14  # 데이터 시작 = 15
    
    data_start_row = header_row + 1
    
    # 2) 테이블 너비 감지 (헤더 행에서 데이터가 있는 열 범위)
    start_col = 1
    end_col = store_col  # 최소 store_col까지
    
    # 헤더 행에서 왼쪽으로 첫 데이터 열 찾기
    for c in range(1, store_col + 1):
        val = ws.Cells(header_row, c).Value
        if norm_text(val) != "":
            start_col = c
            break
    
    # 헤더 행에서 오른쪽으로 마지막 데이터 열 찾기
    for c in range(store_col, 30):  # 최대 30열까지 검색
        val = ws.Cells(header_row, c).Value
        if norm_text(val) != "":
            end_col = c
        else:
            # 2개 연속 빈 셀이면 종료
            val2 = ws.Cells(header_row, c + 1).Value
            if norm_text(val2) == "":
                break
    
    return data_start_row, start_col, end_col


def find_protected_row(ws, vendor: VendorConfig, start_row: int) -> Optional[int]:
    """
    보호할 테이블의 시작 행을 찾기 (헤더 텍스트로 검색)
    예: "공급가액", "부가세" 등의 헤더가 있는 행을 찾음
    """
    if not vendor.protected_table_headers:
        return None
    
    # 검색 범위: start_row부터 충분히 큰 범위
    max_search = start_row + 2000
    
    # 사용 범위 확인
    used_range = ws.UsedRange
    last_used_row = used_range.Row + used_range.Rows.Count - 1
    max_search = min(max_search, last_used_row + 1)
    
    # 여러 열에서 헤더 검색
    used_cols = used_range.Column + used_range.Columns.Count - 1
    search_cols = list(range(1, min(used_cols + 1, 15)))  # 1~14열에서 검색
    
    for r in range(start_row, max_search):
        for col in search_cols:
            cell_value = ws.Cells(r, col).Value
            text = norm_text(cell_value)
            for header in vendor.protected_table_headers:
                if header in text:
                    return r
    
    return None


def apply_borders_to_range(ws, start_row: int, end_row: int, last_col: int = 8):
    """
    테이블에 테두리 적용
    - 모든 셀에 작은 점선 테두리
    - 바깥쪽에 굵은 실선 테두리
    last_col: 마지막 열 (기본값 8 = H열, I열 제외)
    """
    # 테이블 범위 설정 (A열 ~ H열)
    table_range = ws.Range(ws.Cells(start_row, 1), ws.Cells(end_row, last_col))
    
    # 테두리 상수
    xlContinuous = 1      # 실선
    xlDot = 4             # 작은 점선 (dotted)
    xlHairline = 1        # 가장 얇은 선
    xlThin = 2            # 얇은 선
    xlMedium = -4138      # 중간 굵기
    xlEdgeLeft = 7
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlEdgeRight = 10
    xlInsideVertical = 11
    xlInsideHorizontal = 12
    
    # 1) 안쪽 테두리: 작은 점선
    for edge in [xlInsideVertical, xlInsideHorizontal]:
        try:
            border = table_range.Borders(edge)
            border.LineStyle = xlDot
            border.Weight = xlHairline
        except:
            pass
    
    # 2) 바깥쪽 테두리: 굵은 실선
    for edge in [xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight]:
        try:
            border = table_range.Borders(edge)
            border.LineStyle = xlContinuous
            border.Weight = xlMedium
        except:
            pass


def read_existing_stores_via_com_dynamic(
    ws, vendor: VendorConfig, data_start_row: int, table_end_col: int
) -> Tuple[Set[str], int, Optional[int]]:
    """
    Excel COM worksheet에서 기존 매장명 읽기 (동적 레이아웃 사용)
    Returns: (정규화된 매장명 set, 마지막 데이터 행 번호, 보호할 행 번호)
    """
    col_num = col_letter_to_num(vendor.store_col_letter)
    start = data_start_row
    
    # 보호할 행 찾기 (동적으로)
    protected_row = find_protected_row(ws, vendor, start)
    
    # protected_row가 있으면 그 전까지만, 없으면 충분히 큰 범위
    end_row = protected_row - 1 if protected_row else start + 1000
    
    # 벌크로 읽기 (훨씬 빠름)
    range_str = f"{vendor.store_col_letter}{start}:{vendor.store_col_letter}{end_row}"
    values = ws.Range(range_str).Value
    
    existing_normalized: Set[str] = set()
    last_data_row = start - 1
    
    if values:
        for i, row in enumerate(values):
            val = row[0] if isinstance(row, tuple) else row
            text = norm_text(val)
            if text == "":
                break
            existing_normalized.add(normalize_store_name(text))
            last_data_row = start + i
    
    return existing_normalized, last_data_row, protected_row


def insert_stores_via_com_dynamic(
    ws, vendor: VendorConfig, new_stores: List[str], 
    last_data_row: int,
    data_start_row: int,
    table_start_col: int,
    table_end_col: int,
    protected_row: Optional[int] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
):
    """
    Excel COM worksheet에 새 매장 추가 (동적 테이블 레이아웃)
    - protected_row가 있으면 그 위에 행 삽입
    - 마지막 데이터 행을 복사 (내용 포함)
    - 매장명만 교체
    - 테두리 적용 (점선 + 굵은 바깥 테두리)
    """
    if not new_stores:
        return

    col_num = col_letter_to_num(vendor.store_col_letter)
    start = data_start_row
    total = len(new_stores)
    
    if progress_callback:
        progress_callback(0, total, f"행 삽입 준비 중... ({total}개)")
    
    # 삽입 위치 결정
    insert_row = last_data_row + 1 if last_data_row >= start else start
    template_row = last_data_row if last_data_row >= start else None
    
    if protected_row:
        # protected_row 위에 행 삽입
        if total > 0:
            insert_range = ws.Range(f"{insert_row}:{insert_row + total - 1}")
            insert_range.Insert(Shift=-4121)  # xlDown = -4121
    
    if progress_callback:
        progress_callback(20, total, f"행 내용 복사 중... ({total}개)")
    
    # 테이블 범위만 복사 (동적으로 감지된 열 범위)
    if template_row:
        # 열 번호를 문자로 변환
        def num_to_col_letter(n):
            result = ""
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                result = chr(65 + remainder) + result
            return result
        
        start_col_letter = num_to_col_letter(table_start_col)
        end_col_letter = num_to_col_letter(table_end_col)
        
        source_range = ws.Range(f"{start_col_letter}{template_row}:{end_col_letter}{template_row}")
        source_range.Copy()
        for i in range(total):
            dest_range = ws.Range(f"{start_col_letter}{insert_row + i}:{end_col_letter}{insert_row + i}")
            # xlPasteAll = -4104 (전체 붙여넣기: 서식 + 값 + 수식)
            dest_range.PasteSpecial(-4104)
    
    # 클립보드 모드 해제
    try:
        ws.Application.CutCopyMode = False
    except:
        pass
    
    if progress_callback:
        progress_callback(50, total, f"매장명 입력 중... ({total}개)")
    
        # 매장명만 교체
    for i, store in enumerate(new_stores):
        ws.Cells(insert_row + i, col_num).Value = store
        
        if progress_callback and (i + 1) % 10 == 0:
            pct = 50 + int((i + 1) / total * 30)
            progress_callback(pct, total, f"매장명 입력 중: {i + 1}/{total}")
    
    if progress_callback:
        progress_callback(80, total, "테두리 적용 중...")
    
    # 테두리 적용 (동적 테이블 범위)
    apply_borders_to_range(ws, start, insert_row + total - 1, last_col=table_end_col)


# ----------------------------
# 5) 실행 함수 (Excel COM 사용)
# ----------------------------
def run_build(
    list_path: str,
    invoice_path: str,
    vendor_key: str,
    output_path: str,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> Tuple[List[str], str, int, List[str]]:
    """
    Returns: (missing_stores, actual_output_path, existing_count, excluded_stores)
    - missing_stores: 새로 추가된 매장 목록
    - excluded_stores: 명세서에 있지만 전체리스트에 없어서 제외된 매장 목록
    """
    import pythoncom
    import win32com.client as win32
    
    # COM 초기화 (스레드에서 호출 시 필요)
    pythoncom.CoInitialize()
    
    if vendor_key not in VENDOR_CONFIGS:
        raise KeyError(f"등록되지 않은 업체야: {vendor_key}")

    vendor = VENDOR_CONFIGS[vendor_key]

    if progress_callback:
        progress_callback(0, 100, "전체리스트 파일 읽는 중...")

    # 1) 전체리스트에서 {로그인ID: 매장명} 추출
    id_to_store = extract_stores_from_list(list_path, vendor)

    if progress_callback:
        progress_callback(20, 100, "거래명세서 파일 여는 중...")

    # 2) Excel COM으로 거래명세서 열기
    excel = None
    wb = None
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False           # 엑셀 창 숨김 (헤드리스)
        excel.DisplayAlerts = False     # 경고창 숨김
        excel.ScreenUpdating = False    # 화면 업데이트 비활성화 (속도 향상)
        
        # 절대 경로로 변환
        invoice_path = os.path.abspath(invoice_path)
        output_path = os.path.abspath(output_path)
        
        # 원본 파일을 Excel로 직접 열기 (복사하지 않음 - Excel이 SaveAs로 저장)
        wb = excel.Workbooks.Open(invoice_path)
        
        # 날짜 셀에 오늘 날짜 입력 (첫번째 시트)
        if vendor.date_cell:
            from datetime import datetime
            today_str = datetime.now().strftime("%Y-%m-%d")
            first_sheet = wb.Sheets(1)  # 첫번째 시트
            first_sheet.Range(vendor.date_cell).Value = today_str
            if progress_callback:
                progress_callback(22, 100, f"날짜 입력: {today_str} → {vendor.date_cell}")
        
        # 처리할 시트 목록 결정 (invoice_sheets가 있으면 여러 시트, 없으면 단일 시트)
        sheet_names_to_process = vendor.invoice_sheets if vendor.invoice_sheets else [vendor.invoice_sheet]
        total_sheets = len(sheet_names_to_process)

        # 4) ID 시트 찾기 (숨겨져 있으면 임시로 보이게)
        id_ws, id_sheet_was_hidden = find_id_sheet(wb, vendor)
        
        # ID 시트에서 매핑 읽기 (양방향)
        # store_to_id: {매장명: 로그인ID}, id_to_store_invoice: {로그인ID: 명세서 매장명}
        store_to_id, id_to_store_invoice = read_id_sheet_mapping(wb, vendor)
        
        # 전체 기존 매장 수, 추가 매장 수, 제외 매장 수 추적
        total_existing_count = 0
        all_missing_stores = []
        all_missing_ids = []
        all_missing_list_stores = []  # 전체리스트 매장명 (주스샵 매장명)
        all_excluded_stores = []  # 제외된 매장 (전체리스트에 없음)
        
        # 각 시트 처리
        for sheet_idx, sheet_name in enumerate(sheet_names_to_process):
            sheet_progress_base = 20 + int(sheet_idx / total_sheets * 60)  # 20% ~ 80%
            
            if progress_callback:
                progress_callback(sheet_progress_base, 100, f"시트 '{sheet_name}' 처리 중... ({sheet_idx + 1}/{total_sheets})")
            
            # 시트 찾기
            ws = None
            for sheet in wb.Sheets:
                if sheet.Name == sheet_name:
                    ws = sheet
                    break
            
            if ws is None:
                all_sheet_names = [s.Name for s in wb.Sheets]
                wb.Close(False)
                raise KeyError(f"시트 '{sheet_name}'가 없어. 현재 시트: {all_sheet_names}")

            # 테이블 레이아웃 동적 감지 (헤더 행, 테이블 너비)
            data_start_row, table_start_col, table_end_col = detect_table_layout(ws, vendor)
            
            if progress_callback:
                progress_callback(sheet_progress_base + 5, 100, f"[{sheet_name}] 테이블: {data_start_row}행, {table_start_col}~{table_end_col}열")

            # 보호 테이블 찾기
            _, last_data_row, protected_row = read_existing_stores_via_com_dynamic(
                ws, vendor, data_start_row, table_end_col
            )

            # 기존 매장의 로그인ID 확인
            existing_ids, existing_store_names = get_existing_login_ids_dynamic(
                ws, vendor, store_to_id, data_start_row, protected_row
            )
            total_existing_count += len(existing_ids)
            
            # 명세서에 있지만 전체리스트에 없는 매장 찾기
            excluded_stores = []
            for store_name in existing_store_names:
                login_id = store_to_id.get(store_name, "")
                if login_id:
                    # 로그인ID가 전체리스트에 있는지 확인
                    if login_id not in id_to_store:
                        excluded_stores.append(store_name)
                else:
                    # ID 시트에도 없는 매장 = 전체리스트에도 없음
                    excluded_stores.append(store_name)

            # 새로 추가할 매장 찾기 (로그인ID로 비교)
            missing_ids = []
            missing_stores = []
            missing_list_stores = []  # 전체리스트 매장명 (주스샵 매장명)
            for login_id, store_name_from_list in id_to_store.items():
                if login_id not in existing_ids:
                    store_name = id_to_store_invoice.get(login_id, store_name_from_list)
                    missing_ids.append(login_id)
                    missing_stores.append(store_name)
                    missing_list_stores.append(store_name_from_list)  # 전체리스트 매장명
            
            # 매장명 기준으로 정렬
            if missing_stores:
                sorted_pairs = sorted(zip(missing_stores, missing_ids, missing_list_stores), key=lambda x: x[0])
                missing_stores, missing_ids, missing_list_stores = zip(*sorted_pairs)
                missing_stores = list(missing_stores)
                missing_ids = list(missing_ids)
                missing_list_stores = list(missing_list_stores)

            if progress_callback:
                progress_callback(sheet_progress_base + 10, 100, f"[{sheet_name}] 추가할 매장: {len(missing_stores)}개")

            # 새 매장 추가 (행 삽입 방식)
            if missing_stores:
                def make_sub_progress(base):
                    def sub_progress(pct, total, msg):
                        overall_pct = base + int(pct / 100 * 10)
                        if progress_callback:
                            progress_callback(overall_pct, 100, f"[{sheet_name}] {msg}")
                    return sub_progress
                
                insert_stores_via_com_dynamic(
                    ws, vendor, missing_stores, last_data_row, 
                    data_start_row, table_start_col, table_end_col,
                    protected_row, make_sub_progress(sheet_progress_base + 10)
                )
                
                # 첫 시트에서만 all_missing 추적 (ID 시트에 한번만 추가하기 위해)
                if sheet_idx == 0:
                    all_missing_stores = missing_stores
                    all_missing_ids = missing_ids
                    all_missing_list_stores = missing_list_stores
            
            # 제외된 매장 목록을 공급가액 셀 아래에 기록
            if excluded_stores:
                # 첫 시트에서만 all_excluded 추적
                if sheet_idx == 0:
                    all_excluded_stores = excluded_stores
                
                supply_cell = find_supply_amount_cell(ws, vendor, data_start_row)
                if supply_cell:
                    supply_row, supply_col = supply_cell
                    write_excluded_stores_list(ws, vendor, excluded_stores, supply_row, supply_col)
                    if progress_callback:
                        progress_callback(
                            sheet_progress_base + 15, 100, 
                            f"[{sheet_name}] 제외 매장 {len(excluded_stores)}개 기록"
                        )

        # ID 시트에 새 매장 추가 (한번만)
        # - 명세서 매장명: 비워둠
        # - 로그인ID: 입력
        # - 전체리스트 매장명 (주스샵 매장명): 입력
        if all_missing_ids:
            if progress_callback:
                progress_callback(85, 100, "ID 시트에 새 매핑 추가 중...")
            
            add_to_id_sheet(wb, vendor, all_missing_ids, all_missing_list_stores)
        
        missing_stores = all_missing_stores  # 결과 반환용

        if progress_callback:
            progress_callback(93, 100, "ID 시트 숨기는 중...")

        # ID 시트 숨기기 (원래 숨겨져 있었든 아니든 항상 숨김)
        hide_id_sheet(wb, vendor)

        if progress_callback:
            progress_callback(95, 100, "저장 중...")

        # 7) 저장 (.xlsx로) - 항상 SaveAs 사용 (원본 보존)
        # FileFormat: 51 = xlsx, 56 = xls
        wb.SaveAs(output_path, FileFormat=51)
        wb.Close(False)
        wb = None  # finally에서 중복 Close 방지

        if progress_callback:
            excluded_msg = f", 제외: {len(all_excluded_stores)}개" if all_excluded_stores else ""
            progress_callback(100, 100, f"완료! 시트 {total_sheets}개, 추가: {len(missing_stores)}개{excluded_msg}")

        return missing_stores, output_path, total_existing_count, all_excluded_stores
        
    finally:
        try:
            if wb:
                wb.Close(False)
        except:
            pass
        try:
            if excel:
                excel.ScreenUpdating = True  # 복원
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()


# ----------------------------
# 6) Tkinter GUI
# ----------------------------
if __name__ == "__main__":
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import threading

    class InvoiceBuilderApp:
        def __init__(self, root):
            self.root = root
            self.root.title("거래명세서 작성")
            self.root.geometry("700x650")
            self.root.resizable(False, False)

            # 변수들
            self.vendor_var = tk.StringVar()
            self.list_path_var = tk.StringVar()
            self.invoice_path_var = tk.StringVar()
            self.progress_var = tk.IntVar(value=0)
            self.status_var = tk.StringVar(value="대기 중")

            self._create_widgets()

        def _create_widgets(self):
            # 탭 컨트롤 생성
            self.notebook = ttk.Notebook(self.root)
            self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

            # 탭 1: 실행
            self.run_tab = ttk.Frame(self.notebook, padding=20)
            self.notebook.add(self.run_tab, text="실행")
            self._create_run_tab()

            # 탭 2: 업체 관리
            self.vendor_tab = ttk.Frame(self.notebook, padding=20)
            self.notebook.add(self.vendor_tab, text="업체 관리")
            self._create_vendor_tab()

        def _create_run_tab(self):
            """실행 탭 생성"""
            frame = self.run_tab

            # 1) 업체명 드롭다운 (가나다순 정렬)
            ttk.Label(frame, text="업체명:").grid(row=0, column=0, sticky="w", pady=10)
            vendor_names = sorted(VENDOR_CONFIGS.keys())
            self.vendor_combo = ttk.Combobox(
                frame, textvariable=self.vendor_var, values=vendor_names, state="readonly", width=40
            )
            self.vendor_combo.grid(row=0, column=1, sticky="w", pady=10)
            if vendor_names:
                self.vendor_combo.current(0)
            
            # 방향키로 업체 선택 가능하도록 바인딩
            self.vendor_combo.bind("<Up>", self._combo_prev)
            self.vendor_combo.bind("<Down>", self._combo_next)
            self.vendor_combo.bind("<Return>", lambda e: self.vendor_combo.event_generate("<Button-1>"))

            # 2) 전체리스트 파일 선택
            ttk.Label(frame, text="전체리스트 파일:").grid(row=1, column=0, sticky="w", pady=10)
            list_frame = ttk.Frame(frame)
            list_frame.grid(row=1, column=1, sticky="w", pady=10)
            ttk.Entry(list_frame, textvariable=self.list_path_var, width=35).pack(side="left")
            ttk.Button(list_frame, text="찾아보기", command=self._select_list_file).pack(side="left", padx=5)

            # 3) 거래명세서 파일 선택
            ttk.Label(frame, text="거래명세서 파일:").grid(row=2, column=0, sticky="w", pady=10)
            invoice_frame = ttk.Frame(frame)
            invoice_frame.grid(row=2, column=1, sticky="w", pady=10)
            ttk.Entry(invoice_frame, textvariable=self.invoice_path_var, width=35).pack(side="left")
            ttk.Button(invoice_frame, text="찾아보기", command=self._select_invoice_file).pack(side="left", padx=5)

            # 4) 실행 버튼
            self.run_button = ttk.Button(frame, text="실행", command=self._run)
            self.run_button.grid(row=3, column=0, columnspan=2, pady=20)

            # 5) 진행률 바
            ttk.Label(frame, text="진행률:").grid(row=4, column=0, sticky="w", pady=5)
            self.progress_bar = ttk.Progressbar(
                frame, variable=self.progress_var, maximum=100, length=400
            )
            self.progress_bar.grid(row=4, column=1, sticky="w", pady=5)

            # 6) 상태 라벨
            self.status_label = ttk.Label(frame, textvariable=self.status_var, wraplength=500)
            self.status_label.grid(row=5, column=0, columnspan=2, sticky="w", pady=10)

        def _create_vendor_tab(self):
            """업체 관리 탭 생성"""
            frame = self.vendor_tab

            # 상단: 업체 목록
            list_frame = ttk.LabelFrame(frame, text="등록된 업체", padding=10)
            list_frame.pack(fill="x", pady=(0, 10))

            # 업체 리스트박스
            self.vendor_listbox = tk.Listbox(list_frame, height=6, width=60)
            self.vendor_listbox.pack(side="left", fill="x", expand=True)
            self.vendor_listbox.bind("<<ListboxSelect>>", self._on_vendor_select)
            
            scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.vendor_listbox.yview)
            scrollbar.pack(side="right", fill="y")
            self.vendor_listbox.config(yscrollcommand=scrollbar.set)

            # 버튼 프레임
            btn_frame = ttk.Frame(frame)
            btn_frame.pack(fill="x", pady=5)
            ttk.Button(btn_frame, text="새 업체", command=self._new_vendor).pack(side="left", padx=5)
            ttk.Button(btn_frame, text="삭제", command=self._delete_vendor).pack(side="left", padx=5)
            ttk.Button(btn_frame, text="저장", command=self._save_vendor).pack(side="right", padx=5)

            # 하단: 업체 설정 폼
            form_frame = ttk.LabelFrame(frame, text="업체 설정", padding=10)
            form_frame.pack(fill="both", expand=True)

            # 입력 필드들
            self.ve_name = tk.StringVar()
            self.ve_company = tk.StringVar()
            self.ve_group_col = tk.StringVar(value="B")
            self.ve_group_value = tk.StringVar()
            self.ve_group_exclude = tk.StringVar()
            self.ve_invoice_sheet = tk.StringVar(value="상세내역")
            self.ve_invoice_sheets = tk.StringVar()
            self.ve_store_col = tk.StringVar(value="C")
            self.ve_header_text = tk.StringVar(value="매장명")
            self.ve_date_cell = tk.StringVar()

            row = 0
            ttk.Label(form_frame, text="업체명*:").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(form_frame, textvariable=self.ve_name, width=30).grid(row=row, column=1, sticky="w", pady=3)
            
            row += 1
            ttk.Label(form_frame, text="기업명 (A열):").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(form_frame, textvariable=self.ve_company, width=30).grid(row=row, column=1, sticky="w", pady=3)
            
            row += 1
            ttk.Label(form_frame, text="그룹명 열:").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(form_frame, textvariable=self.ve_group_col, width=5).grid(row=row, column=1, sticky="w", pady=3)
            
            row += 1
            ttk.Label(form_frame, text="그룹명 포함:").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(form_frame, textvariable=self.ve_group_value, width=30).grid(row=row, column=1, sticky="w", pady=3)
            
            row += 1
            ttk.Label(form_frame, text="그룹명 제외 (,구분):").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(form_frame, textvariable=self.ve_group_exclude, width=30).grid(row=row, column=1, sticky="w", pady=3)
            
            row += 1
            ttk.Label(form_frame, text="시트명 (단일):").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(form_frame, textvariable=self.ve_invoice_sheet, width=30).grid(row=row, column=1, sticky="w", pady=3)
            
            row += 1
            ttk.Label(form_frame, text="시트명 (복수, ,구분):").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(form_frame, textvariable=self.ve_invoice_sheets, width=30).grid(row=row, column=1, sticky="w", pady=3)
            
            row += 1
            ttk.Label(form_frame, text="매장명 열:").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(form_frame, textvariable=self.ve_store_col, width=5).grid(row=row, column=1, sticky="w", pady=3)
            
            row += 1
            ttk.Label(form_frame, text="헤더 텍스트:").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(form_frame, textvariable=self.ve_header_text, width=30).grid(row=row, column=1, sticky="w", pady=3)
            
            row += 1
            ttk.Label(form_frame, text="날짜 셀 (예: A1):").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(form_frame, textvariable=self.ve_date_cell, width=10).grid(row=row, column=1, sticky="w", pady=3)

            # 업체 목록 초기화
            self._refresh_vendor_list()

        def _refresh_vendor_list(self):
            """업체 목록 새로고침"""
            self.vendor_listbox.delete(0, tk.END)
            for name in sorted(VENDOR_CONFIGS.keys()):
                self.vendor_listbox.insert(tk.END, name)
            
            # 실행 탭의 콤보박스도 새로고침
            vendor_names = sorted(VENDOR_CONFIGS.keys())
            self.vendor_combo['values'] = vendor_names
            if vendor_names and not self.vendor_var.get():
                self.vendor_combo.current(0)

        def _on_vendor_select(self, event):
            """업체 선택 시 폼에 데이터 로드"""
            selection = self.vendor_listbox.curselection()
            if not selection:
                return
            
            name = self.vendor_listbox.get(selection[0])
            if name not in VENDOR_CONFIGS:
                return
            
            config = VENDOR_CONFIGS[name]
            self.ve_name.set(config.name)
            self.ve_company.set(config.company_value or "")
            self.ve_group_col.set(config.group_col or "B")
            self.ve_group_value.set(config.group_value or "")
            self.ve_group_exclude.set(",".join(config.group_exclude) if config.group_exclude else "")
            self.ve_invoice_sheet.set(config.invoice_sheet or "상세내역")
            self.ve_invoice_sheets.set(",".join(config.invoice_sheets) if config.invoice_sheets else "")
            self.ve_store_col.set(config.store_col_letter or "C")
            self.ve_header_text.set(config.table_header_text or "매장명")
            self.ve_date_cell.set(config.date_cell or "")

        def _new_vendor(self):
            """새 업체 폼 초기화"""
            self.vendor_listbox.selection_clear(0, tk.END)
            self.ve_name.set("")
            self.ve_company.set("")
            self.ve_group_col.set("B")
            self.ve_group_value.set("")
            self.ve_group_exclude.set("")
            self.ve_invoice_sheet.set("상세내역")
            self.ve_invoice_sheets.set("")
            self.ve_store_col.set("C")
            self.ve_header_text.set("매장명")
            self.ve_date_cell.set("")

        def _save_vendor(self):
            """업체 저장"""
            name = self.ve_name.get().strip()
            if not name:
                messagebox.showerror("오류", "업체명을 입력하세요.")
                return
            
            # 그룹 제외 목록 파싱
            group_exclude = None
            if self.ve_group_exclude.get().strip():
                group_exclude = [x.strip() for x in self.ve_group_exclude.get().split(",") if x.strip()]
            
            # 복수 시트 파싱
            invoice_sheets = None
            if self.ve_invoice_sheets.get().strip():
                invoice_sheets = [x.strip() for x in self.ve_invoice_sheets.get().split(",") if x.strip()]
            
            config = VendorConfig(
                name=name,
                list_sheet=None,
                header_row=3,
                list_id_col="D",
                invoice_sheet=self.ve_invoice_sheet.get().strip() or "상세내역",
                invoice_sheets=invoice_sheets,
                store_col_letter=self.ve_store_col.get().strip() or "C",
                table_header_text=self.ve_header_text.get().strip() or "매장명",
                id_sheet="ID",
                id_store_col="B",
                id_login_col="C",
                id_start_row=2,
                protected_table_headers=["공급가액", "부가세"],
                company_value=self.ve_company.get().strip() or None,
                group_col=self.ve_group_col.get().strip() or None,
                group_value=self.ve_group_value.get().strip() or None,
                group_exclude=group_exclude,
                date_cell=self.ve_date_cell.get().strip() or None,
            )
            
            add_vendor_config(config)
            self._refresh_vendor_list()
            messagebox.showinfo("완료", f"'{name}' 업체가 저장되었습니다.")

        def _delete_vendor(self):
            """업체 삭제"""
            selection = self.vendor_listbox.curselection()
            if not selection:
                messagebox.showwarning("경고", "삭제할 업체를 선택하세요.")
                return
            
            name = self.vendor_listbox.get(selection[0])
            if messagebox.askyesno("확인", f"'{name}' 업체를 삭제하시겠습니까?"):
                delete_vendor_config(name)
                self._refresh_vendor_list()
                self._new_vendor()  # 폼 초기화

        def _select_list_file(self):
            path = filedialog.askopenfilename(
                title="전체리스트 파일 선택",
                filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")]
            )
            if path:
                self.list_path_var.set(path)

        def _select_invoice_file(self):
            path = filedialog.askopenfilename(
                title="거래명세서 파일 선택",
                filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")]
            )
            if path:
                self.invoice_path_var.set(path)

        def _combo_prev(self, event):
            """방향키 위로 이전 항목 선택"""
            current = self.vendor_combo.current()
            if current > 0:
                self.vendor_combo.current(current - 1)
            return "break"  # 기본 동작 방지

        def _combo_next(self, event):
            """방향키 아래로 다음 항목 선택"""
            current = self.vendor_combo.current()
            values = self.vendor_combo['values']
            if current < len(values) - 1:
                self.vendor_combo.current(current + 1)
            return "break"  # 기본 동작 방지

        def _update_progress(self, pct: int, total: int, msg: str):
            """진행률 업데이트 (메인 스레드에서 실행)"""
            self.progress_var.set(pct)
            self.status_var.set(f"{pct}% - {msg}")
            self.root.update_idletasks()

        def _run_task(self, vendor_key, list_path, invoice_path, output_path):
            """백그라운드 스레드에서 실행되는 작업"""
            try:
                def progress_callback(pct, total, msg):
                    # 람다에서 값을 캡처하기 위해 기본 인자 사용
                    self.root.after(0, lambda p=pct, t=total, m=msg: self._update_progress(p, t, m))
                
                missing, actual_output_path, existing_count, excluded = run_build(
                    list_path, invoice_path, vendor_key, output_path, progress_callback
                )
                
                # 결과 표시
                excluded_msg = f", 제외: {len(excluded)}개" if excluded else ""
                if missing:
                    result_msg = f"100% - 완료! 저장: {os.path.basename(actual_output_path)} | 기존: {existing_count}개, 추가: {len(missing)}개{excluded_msg}"
                else:
                    result_msg = f"100% - 완료! 저장: {os.path.basename(actual_output_path)} | 기존: {existing_count}개, 추가할 매장 없음{excluded_msg}"
                
                self.root.after(0, lambda r=result_msg: self._update_progress(100, 100, r))
                
            except Exception as e:
                error_msg = f"오류: {str(e)}"
                self.root.after(0, lambda m=error_msg: self._update_progress(0, 100, m))
            finally:
                self.root.after(0, lambda: self.run_button.config(state="normal"))

        def _run(self):
            vendor_key = self.vendor_var.get()
            list_path = self.list_path_var.get()
            invoice_path = self.invoice_path_var.get()

            # 유효성 검사
            if not vendor_key:
                self.status_var.set("오류: 업체명을 선택해주세요.")
                return
            if not list_path:
                self.status_var.set("오류: 전체리스트 파일을 선택해주세요.")
                return
            if not invoice_path:
                self.status_var.set("오류: 거래명세서 파일을 선택해주세요.")
                return

            # 저장 경로 자동 생성 (거래명세서와 같은 폴더에 _완성 붙여서 저장, 항상 .xlsx)
            folder = os.path.dirname(invoice_path)
            filename = os.path.basename(invoice_path)
            name, ext = os.path.splitext(filename)
            # 항상 .xlsx로 저장
            output_path = os.path.join(folder, f"{name}_완성.xlsx")

            # 버튼 비활성화
            self.run_button.config(state="disabled")
            self.progress_var.set(0)
            self.status_var.set("시작 중...")

            # 백그라운드 스레드에서 실행
            thread = threading.Thread(
                target=self._run_task,
                args=(vendor_key, list_path, invoice_path, output_path),
                daemon=True
            )
            thread.start()

    root = tk.Tk()
    app = InvoiceBuilderApp(root)
    root.mainloop()
