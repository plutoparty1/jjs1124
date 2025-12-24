"""
Outlook 메일 자동 발송 스크립트

사용법:
    python send_mail.py                          # 기본 경로의 엑셀 파일 사용
    python send_mail.py "C:\\메일목록.xlsx"      # 엑셀 경로 지정
    python send_mail.py "메일목록.xlsx" "로고.png"  # 엑셀 + 로고 경로 지정

엑셀 파일 필수 컬럼:
    - 발송 여부    : Y면 발송, 그 외는 건너뜀
    - 이메일       : 받는 사람 이메일 주소
    - 이름         : 받는 사람 이름
    - 제목         : 메일 제목
    - 본문 내용    : 메일 본문 (HTML 가능)
    - 참조자 이메일 : (선택) 참조자, 세미콜론(;)으로 여러명 구분
    - 첨부파일 경로 : (선택) 첨부파일, 세미콜론(;)으로 여러개 구분
"""

from __future__ import annotations

import os
import re
import sys
import time
import logging
import traceback
from datetime import datetime
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    import pandas as pd

# ============================================================
# 환경 체크
# ============================================================

if sys.platform != "win32":
    print("[오류] Windows에서만 실행 가능합니다")
    sys.exit(1)

try:
    import pandas as pd
except ImportError:
    print("[오류] pandas 없음 → pip install pandas openpyxl")
    sys.exit(1)

try:
    import win32com.client as win32
except ImportError:
    print("[오류] pywin32 없음 → pip install pywin32")
    sys.exit(1)

try:
    import tkinter as tk
    from tkinter import ttk, messagebox
except ImportError:
    tk = None

# ============================================================
# 설정
# ============================================================

EXCEL_PATH = r"C:\Users\Code\mail_list.xlsx"
LOGO_PATH = r"C:\Users\정산-PC\Desktop\jjs\04_기타\plantym_logo_wordmark.png"

REQUIRED_COLUMNS = ("발송 여부", "이메일", "이름", "제목", "본문 내용")
SEND_DELAY_SEC = 1.0      # 메일 간 대기 (초)
MAX_ATTACHMENT_MB = 25    # 첨부파일 최대 크기 (MB)

# Outbox 체크 옵션 (원하는대로 조절 가능)
OUTBOX_WAIT_MAX_SEC = 240       # 최대 대기 시간
OUTBOX_WAIT_INTERVAL = 2.0      # 체크 간격
OUTBOX_ZERO_STREAK = 5          # "0건" 연속 몇 번이면 비워짐으로 인정?

SIGNATURE = """
<div>
<div style="font-family:Aptos,Aptos_EmbeddedFont,Aptos_MSFontService,Calibri,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
<br>
</div>
<div style="font-family:Aptos,Aptos_EmbeddedFont,Aptos_MSFontService,Calibri,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
<br>
</div>
<table style="text-align:left; border-top:1pt solid gray; background-color:white; color:rgb(36,36,36); box-sizing:border-box; border-collapse:collapse; border-spacing:0px">
<tbody>
<tr>
<td style="text-align:left; padding-top:7.5pt; width:262.5pt; height:107.5px">
<div style="margin:7.5pt 0px 0px">
<p style="text-align:left; margin-top:0px; margin-bottom:0px"><span style="font-family:&quot;Malgun Gothic&quot;,Gulim,&quot;Segoe UI&quot;,-apple-system,BlinkMacSystemFont,Roboto,&quot;Helvetica Neue&quot;,sans-serif; font-size:13pt; color:rgb(104,137,61)"><b>정 지 수</b>&nbsp;Jung</span><span style="font-family:&quot;Malgun Gothic&quot;,Gulim,&quot;Segoe UI&quot;,-apple-system,BlinkMacSystemFont,Roboto,&quot;Helvetica Neue&quot;,sans-serif; font-size:9pt; color:rgb(104,137,61)">&nbsp;Ji Soo</span></p>
</div>
<p style="text-align:left; margin-top:0px; margin-bottom:0px"><span style="font-family:&quot;Malgun Gothic&quot;,Gulim,&quot;Segoe UI&quot;,-apple-system,BlinkMacSystemFont,Roboto,&quot;Helvetica Neue&quot;,sans-serif; font-size:9pt; color:rgb(36,36,36)">&nbsp;</span></p>
<p style="text-align:left; text-indent:0px; background-color:white; margin:0px"><span style="font-family:&quot;Malgun Gothic&quot;,Gulim,&quot;Segoe UI&quot;,-apple-system,BlinkMacSystemFont,Roboto,&quot;Helvetica Neue&quot;,sans-serif; font-size:9pt; color:rgb(36,36,36)">경영지원팀</span><span style="font-family:Calibri,sans-serif; font-size:9pt; color:rgb(36,36,36)">&nbsp;</span></p>
<p style="text-align:left; text-indent:0px; background-color:white; margin:0px"><span style="font-family:Calibri,sans-serif; font-size:9pt; color:black">Professional&nbsp;/ Management Support Team</span></p>
</td>
<td style="text-align:left; padding-top:7.5pt; width:78.75pt; height:107.5px">
<p style="text-align:left; margin-top:0px; margin-bottom:0px"><span style="font-family:&quot;Malgun Gothic&quot;,Gulim,&quot;Segoe UI&quot;,-apple-system,BlinkMacSystemFont,Roboto,&quot;Helvetica Neue&quot;,sans-serif; font-size:9pt; color:blue"><a href="http://www.plantym.com/" title="http://www.plantym.com/" style="color:blue; margin:0px"><img src="cid:plantym_logo" width="105" height="47" style="width:105px; height:47px; min-width:auto; min-height:auto; margin:0px"></a></span></p>
</td>
</tr>
<tr>
<td colspan="2" style="text-align:left; padding-top:15pt; width:459px; height:84px">
<p style="text-align:left; margin-top:0px; margin-bottom:0px"><span style="font-family:&quot;Malgun Gothic&quot;,Gulim,&quot;Segoe UI&quot;,-apple-system,BlinkMacSystemFont,Roboto,&quot;Helvetica Neue&quot;,sans-serif; font-size:9pt; color:rgb(12,100,192)"><u>jjs1124<a href="mailto:yyyoon@plantym.com" title="mailto:yyyoon@plantym.com" style="color:rgb(12,100,192); margin:0px">@plantym.com</a></u></span></p>
<p style="text-align:left; margin-top:0px; margin-bottom:0px"><span style="font-family:&quot;Malgun Gothic&quot;,Gulim,&quot;Segoe UI&quot;,-apple-system,BlinkMacSystemFont,Roboto,&quot;Helvetica Neue&quot;,sans-serif; font-size:9pt; color:rgb(104,137,61)"><b>PHONE</b></span><span style="font-family:&quot;Malgun Gothic&quot;,Gulim,&quot;Segoe UI&quot;,-apple-system,BlinkMacSystemFont,Roboto,&quot;Helvetica Neue&quot;,sans-serif; font-size:9pt; color:rgb(36,36,36)">&nbsp;+82.70.4489.7181 </span><span style="font-family:&quot;Malgun Gothic&quot;,Gulim,&quot;Segoe UI&quot;,-apple-system,BlinkMacSystemFont,Roboto,&quot;Helvetica Neue&quot;,sans-serif; font-size:9pt; color:rgb(104,137,61)"><b>FAX</b></span><span style="font-family:&quot;Malgun Gothic&quot;,Gulim,&quot;Segoe UI&quot;,-apple-system,BlinkMacSystemFont,Roboto,&quot;Helvetica Neue&quot;,sans-serif; font-size:9pt; color:rgb(36,36,36)">&nbsp;+82.31.709.5100</span></p>
<p style="text-align:left; margin-top:0px; margin-bottom:0px"><span style="font-family:&quot;Malgun Gothic&quot;,Gulim,&quot;Segoe UI&quot;,-apple-system,BlinkMacSystemFont,Roboto,&quot;Helvetica Neue&quot;,sans-serif; font-size:9pt; color:rgb(36,36,36)">13494 경기도 성남시 분당구 대왕판교로&nbsp;670 유스페이스2 A동&nbsp;4F (주)플랜티엠</span></p>
</td>
</tr>
</tbody>
</table>
<div style="font-family:Aptos,Aptos_EmbeddedFont,Aptos_MSFontService,Calibri,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
<br>
</div>
</div>
"""

# 관대한 이메일 정규식: 점(.)이 포함된 이메일도 허용
# 로컬 부분: 문자, 숫자, 점, 언더스코어, 하이픈, 플러스 등 허용 (시작/끝 제한 완화)
# 도메인 부분: 문자, 숫자, 점, 하이픈 허용, 최소 2자 이상의 TLD
_EMAIL_RE = re.compile(r"^[a-zA-Z0-9][a-zA-Z0-9._%+-]*@[a-zA-Z0-9][a-zA-Z0-9.-]*\.[a-zA-Z]{2,}$")

# ============================================================
# 유틸리티
# ============================================================

def 에러(코드: str, 메시지: str, 원인: str = "", 해결: str = "") -> None:
    print(f"\n[오류:{코드}] {메시지}")
    if 원인:
        print(f"  원인: {원인}")
    if 해결:
        print(f"  해결: {해결}")


def 셀값(값: Any) -> str:
    if 값 is None or pd.isna(값):
        return ""
    return str(값).strip()


def 이메일_검증(이메일: str) -> bool:
    return bool(이메일 and _EMAIL_RE.match(이메일))


# ============================================================
# 엑셀 처리
# ============================================================

def 엑셀_불러오기(파일경로: str, 로그: logging.Logger | None = None) -> pd.DataFrame | None:
    if not 파일경로.lower().endswith((".xlsx", ".xls")):
        에러("E002", "엑셀 파일이 아닙니다", 파일경로, ".xlsx 파일을 사용하세요")
        return None

    try:
        # keep_default_na=False: 빈 셀을 NaN이 아닌 빈 문자열로 처리
        # header=0: 첫 번째 행을 헤더로 사용
        df = pd.read_excel(파일경로, keep_default_na=False, header=0)
        if 로그:
            로그.debug(f"엑셀 파일 읽기 완료: 총 {len(df)}행, 컬럼: {list(df.columns)}")
    except FileNotFoundError:
        에러("E001", "파일을 찾을 수 없습니다", 파일경로)
        return None
    except PermissionError:
        에러("E003", "파일이 열려있습니다", "다른 프로그램에서 사용 중")
        return None
    except Exception as e:
        에러("E004", "파일 읽기 실패", str(e))
        return None

    누락 = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if 누락:
        에러("E005", "필수 컬럼 누락", ", ".join(누락))
        return None

    return df


def 발송목록_만들기(df: pd.DataFrame, 로그: logging.Logger) -> list[dict[str, Any]]:
    목록: list[dict[str, Any]] = []

    # 인덱스를 리셋하여 연속된 인덱스 보장
    df = df.reset_index(drop=True)
    
    로그.info(f"발송목록 생성 시작: DataFrame 총 {len(df)}행")
    
    # 11행과 12행 주변 디버깅을 위해 상세 로그 추가
    for idx, (i, 행) in enumerate(df.iterrows()):
        # 엑셀 행 번호 = 헤더(1행) + 인덱스(0부터 시작) + 1
        엑셀행 = idx + 2
        
        # 10~13행 주변은 상세 로그 출력
        if 9 <= 엑셀행 <= 13:
            발송여부값 = 셀값(행.get("발송 여부"))
            이메일값 = 셀값(행.get("이메일"))
            이름값 = 셀값(행.get("이름"))
            로그.info(f"[디버그] 행 {엑셀행} (인덱스 {idx}): 발송여부='{발송여부값}', 이메일='{이메일값}', 이름='{이름값}'")
            로그.info(f"[디버그] 행 {엑셀행} 전체 데이터: {dict(행)}")
        
        발송여부값 = 셀값(행.get("발송 여부")).upper()
        로그.debug(f"행 {엑셀행} (인덱스 {idx}): 발송 여부 = '{발송여부값}'")
        
        if 발송여부값 != "Y":
            if 9 <= 엑셀행 <= 13:
                로그.info(f"[디버그] 행 {엑셀행}: 발송 여부가 'Y'가 아님 ('{발송여부값}') → 건너뜀")
            continue

        이메일_원본 = 셀값(행.get("이메일"))
        if not 이메일_원본:
            로그.warning(f"행 {엑셀행}: 이메일 없음 → 건너뜀")
            continue
        
        # 이메일 컬럼도 세미콜론으로 구분된 여러 이메일 처리
        이메일_목록: list[str] = []
        for email in 이메일_원본.replace("; ", ";").split(";"):
            email = email.strip()
            if email:
                검증결과 = 이메일_검증(email)
                if 검증결과:
                    이메일_목록.append(email)
                else:
                    로그.warning(f"행 {엑셀행}: 이메일 형식 오류 '{email}' (길이: {len(email)}) → 제외")
                    # 디버깅: 정규식이 왜 실패하는지 확인
                    로그.debug(f"  이메일 검증 실패 상세: '{email}'")
        
        if not 이메일_목록:
            로그.warning(f"행 {엑셀행}: 유효한 이메일 없음 → 건너뜀")
            continue
        
        # 첫 번째 이메일을 받는 사람으로, 나머지는 참조자에 추가
        받는사람_이메일 = 이메일_목록[0]
        추가_참조 = 이메일_목록[1:] if len(이메일_목록) > 1 else []

        참조_원본 = 셀값(행.get("참조자 이메일"))
        참조_검증됨: list[str] = []
        if 참조_원본:
            # 세미콜론 또는 세미콜론+공백으로 구분
            for cc in 참조_원본.replace("; ", ";").split(";"):
                cc = cc.strip()
                if cc and 이메일_검증(cc):
                    참조_검증됨.append(cc)
                elif cc:
                    로그.warning(f"행 {엑셀행}: 참조자 '{cc}' 형식 오류 → 제외")
        
        # 이메일 컬럼의 추가 이메일도 참조자에 포함
        참조_검증됨.extend(추가_참조)

        목록.append({
            "행": 엑셀행,
            "이름": 셀값(행.get("이름")) or "(이름없음)",
            "받는사람": 받는사람_이메일,
            "참조": "; ".join(참조_검증됨) if 참조_검증됨 else "",  # Outlook은 세미콜론+공백 형식 선호
            "제목": 셀값(행.get("제목")) or "(제목없음)",
            "본문": 셀값(행.get("본문 내용")),
            "첨부": 셀값(행.get("첨부파일 경로")),
        })

    return 목록


# ============================================================
# Outlook 발송
# ============================================================

def Outlook_연결() -> Any | None:
    try:
        outlook = win32.Dispatch("Outlook.Application")
        outlook.GetNamespace("MAPI")
        return outlook
    except Exception as e:
        err = str(e).lower()
        if "class not registered" in err:
            에러("O001", "Outlook이 설치되지 않았습니다")
        else:
            에러("O002", "Outlook 연결 실패", str(e))
        return None


def 메일_발송(outlook: Any, 항목: dict[str, Any], 로고경로: str, 로그: logging.Logger) -> tuple[bool, str]:
    메일 = None
    try:
        메일 = outlook.CreateItem(0)
        메일.To = 항목["받는사람"]
        메일.Subject = 항목["제목"]

        if 항목["참조"]:
            try:
                # Outlook CC 필드에 세미콜론+공백으로 구분된 문자열 할당
                메일.CC = 항목["참조"]
                로그.debug(f"  참조자 설정: {항목['참조']}")
            except Exception as e:
                로그.warning(f"  참조자 설정 실패: {e}")
                # 참조자 설정 실패해도 메일 발송은 계속 진행

        # 로고 첨부 (cid)
        if 로고경로 and os.path.exists(로고경로):
            try:
                첨부 = 메일.Attachments.Add(로고경로)
                첨부.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "plantym_logo"
                )
            except Exception as e:
                로그.warning(f"  로고 첨부 실패: {e}")

        메일.HTMLBody = (항목["본문"] or "") + SIGNATURE

        # 첨부파일
        if 항목["첨부"]:
            for 경로 in 항목["첨부"].split(";"):
                경로 = 경로.strip()
                if not 경로:
                    continue
                if not os.path.exists(경로):
                    로그.warning(f"  첨부파일 없음: {경로}")
                    continue
                크기 = os.path.getsize(경로) / (1024 * 1024)
                if 크기 > MAX_ATTACHMENT_MB:
                    로그.warning(f"  첨부 너무 큼: {경로} ({크기:.1f}MB)")
                    continue
                try:
                    메일.Attachments.Add(경로)
                except Exception as e:
                    로그.warning(f"  첨부 실패 '{경로}': {e}")


        메일.Send()
        return True, ""

    except Exception as e:
        에러상세 = f"{type(e).__name__}: {e}\n{traceback.format_exc()}"
        로그.debug(에러상세)
        return False, f"{type(e).__name__}: {e}"
    finally:
        if 메일:
            try:
                del 메일
            except Exception:
                pass


# ============================================================
# GUI
# ============================================================

def 확인창_띄우기(목록: list[dict[str, Any]]) -> bool:
    if not 목록:
        print("\n발송할 메일이 없습니다")
        return False

    if tk is None:
        return 콘솔_확인(목록)

    try:
        창 = tk.Tk()
        창.title("발송 확인")
        창.geometry("1200x520")

        tk.Label(창, text=f"총 {len(목록)}건 발송 예정",
                font=("맑은 고딕", 14, "bold")).pack(pady=10)

        프레임 = tk.Frame(창)
        프레임.pack(fill=tk.BOTH, expand=True, padx=10)

        컬럼 = ("번호", "행", "이름", "받는사람", "제목", "첨부")
        표 = ttk.Treeview(프레임, columns=컬럼, show="headings", height=15)

        for c in 컬럼:
            표.heading(c, text=c)
        표.column("번호", width=50)
        표.column("행", width=50)
        표.column("이름", width=100)
        표.column("받는사람", width=250)
        표.column("제목", width=320)
        표.column("첨부", width=350)

        스크롤 = ttk.Scrollbar(프레임, command=표.yview)
        표.configure(yscrollcommand=스크롤.set)
        표.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        스크롤.pack(side=tk.RIGHT, fill=tk.Y)

        for i, m in enumerate(목록, 1):
            첨부원본 = m["첨부"]
            첨부목록 = [
                os.path.basename(p.strip()) or p.strip()
                for p in (첨부원본.split(";") if 첨부원본 else [])
                if p.strip()
            ]
            if len(첨부목록) > 3:
                첨부표시 = ", ".join(첨부목록[:3]) + f" 외 {len(첨부목록) - 3}개"
            else:
                첨부표시 = ", ".join(첨부목록) if 첨부목록 else "-"
            표.insert("", "end", values=(i, m["행"], m["이름"], m["받는사람"], m["제목"], 첨부표시))

        결과 = [False]

        def 확인():
            결과[0] = True
            창.destroy()

        버튼프레임 = tk.Frame(창)
        버튼프레임.pack(pady=10)

        tk.Button(버튼프레임, text="발송", command=확인,
                 bg="#4CAF50", fg="white", width=10, height=2,
                 font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(버튼프레임, text="취소", command=창.destroy,
                 bg="#f44336", fg="white", width=10, height=2,
                 font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT, padx=5)

        창.update()
        x = (창.winfo_screenwidth() - 창.winfo_width()) // 2
        y = (창.winfo_screenheight() - 창.winfo_height()) // 2
        창.geometry(f"+{x}+{y}")

        창.mainloop()
        return 결과[0]

    except Exception:
        return 콘솔_확인(목록)


def 콘솔_확인(목록: list[dict[str, Any]]) -> bool:
    print(f"\n{'='*50}")
    print(f"총 {len(목록)}건 발송 예정")
    print(f"{'='*50}")
    for i, m in enumerate(목록[:5], 1):
        print(f"  {i}. {m['이름']} <{m['받는사람']}>")
    if len(목록) > 5:
        print(f"  ... 외 {len(목록)-5}건")
    print(f"{'='*50}")

    while True:
        응답 = input("발송할까요? (y/n): ").strip().lower()
        if 응답 in ("y", "yes"):
            return True
        if 응답 in ("n", "no"):
            return False


# ============================================================
# Outbox 모니터링 (수정본: '연속 0건' 확인까지 대기)
# ============================================================

def Outbox_비움_대기(
    outlook: Any,
    로그: logging.Logger,
    최대초: int = OUTBOX_WAIT_MAX_SEC,
    인터벌: float = OUTBOX_WAIT_INTERVAL,
    연속0회: int = OUTBOX_ZERO_STREAK,
) -> bool:
    """Outbox가 연속 N번 0건이 될 때까지 대기."""
    try:
        ns = outlook.GetNamespace("MAPI")
        outbox = ns.GetDefaultFolder(4)  # 4 = olFolderOutbox
    except Exception as e:
        로그.warning(f"Outbox 비움 대기 실패(폴더): {e}")
        return False

    시작 = time.time()
    zero_streak = 0
    직전 = None

    while True:
        try:
            # stale 방지: 매번 Items를 다시 잡아준다
            items = outbox.Items
            개수 = items.Count
        except Exception as e:
            로그.warning(f"Outbox 비움 대기 실패(카운트): {e}")
            return False

        if 개수 != 직전:
            로그.info(f"Outbox 대기 {개수}건")
            직전 = 개수

        if 개수 == 0:
            zero_streak += 1
            로그.info(f"Outbox 0건 확인 ({zero_streak}/{연속0회})")
            if zero_streak >= 연속0회:
                print("\nOutbox가 비워졌습니다.")
                로그.info("Outbox 비워짐(연속 0 확인) → 완료")
                return True
        else:
            zero_streak = 0

        if time.time() - 시작 >= 최대초:
            print(f"\nOutbox 비움 대기 시간 초과 (현재 {개수}건)")
            로그.info(f"Outbox 비움 대기 시간 초과: {개수}건 남음")
            return False

        time.sleep(인터벌)


def 미처리_강제발송(outlook: Any, 로그: logging.Logger) -> tuple[int, int, bool]:
    """Outbox 대기 메일을 즉시 Send() 호출로 발송 시도."""
    try:
        ns = outlook.GetNamespace("MAPI")
        outbox = ns.GetDefaultFolder(4)  # 4 = olFolderOutbox
        items = outbox.Items
        개수 = items.Count
    except Exception as e:
        로그.warning(f"미처리 강제발송 실패(폴더): {e}")
        return 0, 0, False

    if 개수 == 0:
        로그.info("Outbox 대기 없음 → 강제발송 스킵")
        return 0, 0, True

    대기목록 = []
    try:
        for i in range(개수, 0, -1):
            try:
                대기목록.append(items.Item(i))
            except Exception:
                pass
    except Exception as e:
        로그.warning(f"미처리 강제발송 실패(목록화): {e}")
        return 0, 0, False

    성공 = 실패 = 0
    for m in 대기목록:
        try:
            subj = getattr(m, "Subject", "(제목 없음)")
            받는이 = getattr(m, "To", "")
            로그.info(f"미처리 강제발송: {subj} → {받는이}")
            m.Send()
            성공 += 1
        except Exception as e:
            실패 += 1
            로그.error(f"미처리 강제발송 실패: {e}")

    로그.info(f"미처리 강제발송 결과: 성공 {성공}건 / 실패 {실패}건")
    return 성공, 실패, True


def 미처리_확인(outlook: Any, 로그: logging.Logger) -> bool:
    """Outbox에 남은 미처리 메일을 확인 후 계속 여부 반환."""
    try:
        ns = outlook.GetNamespace("MAPI")
        outbox = ns.GetDefaultFolder(4)
        items = outbox.Items
        개수 = items.Count
    except Exception as e:
        로그.warning(f"미처리 확인 실패: {e}")
        return True

    if 개수 == 0:
        로그.info("미처리 메일 없음 (Outbox 비어 있음)")
        return True

    미리보기 = []
    try:
        for i in range(1, min(개수, 5) + 1):
            m = items.Item(i)
            subj = getattr(m, "Subject", "(제목 없음)")
            받는이 = getattr(m, "To", "")
            미리보기.append(f"{i}. {subj} → {받는이}")
    except Exception:
        pass

    로그.info(f"미처리 메일 {개수}건 발견")
    if 미리보기:
        for line in 미리보기:
            로그.info(f"  - {line}")
    if 개수 > 5:
        로그.info(f"  - ... 외 {개수 - 5}건")

    # 0. 강제 발송 시도 여부
    강제 = False
    if tk is not None:
        try:
            root = tk.Tk()
            root.withdraw()
            강제 = messagebox.askyesno("미처리 발송", "Outbox 대기 메일을 즉시 발송 시도할까?", parent=root)
            root.destroy()
        except Exception:
            강제 = False
    else:
        while True:
            응답 = input("Outbox 대기 메일을 즉시 발송 시도할까? (y/n): ").strip().lower()
            if 응답 in ("y", "yes"):
                강제 = True
                break
            if 응답 in ("n", "no"):
                강제 = False
                break

    if 강제:
        성공, 실패, 시도 = 미처리_강제발송(outlook, 로그)
        if not 시도:
            로그.warning("강제발송 시도 실패 → 기존 흐름 계속")
        # ✅ 여기 핵심: 강제발송 후 Outbox가 비워질 때까지 대기
        Outbox_비움_대기(outlook, 로그)

    # 이후 새 발송 계속 여부
    안내 = "Outbox(보낼 편지함)에 미처리 메일이 남아있어.\n그래도 새 메일 발송을 계속할까?"
    if tk is not None:
        try:
            root = tk.Tk()
            root.withdraw()
            ok = messagebox.askyesno("미처리 메일 확인", 안내, parent=root)
            root.destroy()
            return ok
        except Exception:
            pass

    print(f"\n{'='*50}")
    print("미처리 메일 알림")
    print(f"총 {개수}건 대기 중")
    for line in 미리보기:
        print(f"  {line}")
    print(f"{'='*50}")

    while True:
        응답 = input("계속 발송할까? (y/n): ").strip().lower()
        if 응답 in ("y", "yes"):
            return True
        if 응답 in ("n", "no"):
            return False


def 발송후_미처리_처리(outlook: Any, 로그: logging.Logger) -> None:
    """발송 완료 후 Outbox 잔여 건 안내 및 '비워질 때까지' 체크."""
    try:
        ns = outlook.GetNamespace("MAPI")
        outbox = ns.GetDefaultFolder(4)
        items = outbox.Items
        개수 = items.Count
    except Exception as e:
        로그.warning(f"발송 후 Outbox 확인 실패: {e}")
        return

    if 개수 == 0:
        로그.info("발송 후 Outbox 비어 있음")
        return

    미리보기 = []
    try:
        for i in range(1, min(개수, 5) + 1):
            m = items.Item(i)
            subj = getattr(m, "Subject", "(제목 없음)")
            받는이 = getattr(m, "To", "")
            미리보기.append(f"{i}. {subj} → {받는이}")
    except Exception:
        pass

    안내 = f"발송 작업 후 Outbox(보낼 편지함)에 {개수}건이 남아있어.\n\n"
    if 미리보기:
        안내 += "예시:\n" + "\n".join(미리보기) + "\n\n"
    안내 += "남은 메일을 바로 발송(Send) 시도할까?\n(그리고 Outbox가 비워질 때까지 계속 확인할게)"

    진행 = False
    if tk is not None:
        try:
            root = tk.Tk()
            root.withdraw()
            진행 = messagebox.askyesno("Outbox 잔여 메일", 안내, parent=root)
            root.destroy()
        except Exception:
            진행 = False
    else:
        while True:
            응답 = input("Outbox 잔여 메일을 바로 발송 시도할까? (y/n): ").strip().lower()
            if 응답 in ("y", "yes"):
                진행 = True
                break
            if 응답 in ("n", "no"):
                진행 = False
                break

    if not 진행:
        로그.info("사용자 선택으로 Outbox 잔여 발송 건너뜀")
        return

    성공, 실패, 시도 = 미처리_강제발송(outlook, 로그)
    if not 시도:
        로그.warning("Outbox 잔여 발송 시도를 수행하지 못함")
        return

    # ✅ 여기 핵심: 1번 확인이 아니라 '비워질 때까지' 확인
    비움 = Outbox_비움_대기(outlook, 로그)

    try:
        잔여 = outbox.Items.Count
    except Exception:
        잔여 = None

    결과메시지 = f"잔여 메일 발송 시도 완료\n성공 {성공}건 / 실패 {실패}건"
    if 잔여 is not None:
        결과메시지 += f"\n현재 Outbox 잔여: {잔여}건"
    if not 비움:
        결과메시지 += "\n(참고) 시간 내 Outbox 비움 확인이 안됐어도 실제 발송은 진행 중일 수 있어"

    if tk is not None:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("Outbox 잔여 발송 결과", 결과메시지, parent=root)
            root.destroy()
        except Exception:
            pass
    else:
        print(f"\n{결과메시지}")

    로그.info(결과메시지.replace("\n", " | "))


# ============================================================
# 메인
# ============================================================

def main() -> int:
    로그 = logging.getLogger()
    로그.setLevel(logging.DEBUG)

    콘솔 = logging.StreamHandler()
    콘솔.setLevel(logging.INFO)
    콘솔.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

    로그.addHandler(콘솔)
    # 로그 파일 생성을 비활성화 (폴더에 쌓이지 않도록)

    try:
        로그.info("시작")

        엑셀 = sys.argv[1] if len(sys.argv) > 1 else EXCEL_PATH
        로고 = sys.argv[2] if len(sys.argv) > 2 else LOGO_PATH
        로그.info(f"엑셀: {엑셀}")

        df = 엑셀_불러오기(엑셀, 로그)
        if df is None:
            return 1
        로그.info(f"엑셀 로드: {len(df)}행")
        
        # 디버깅: 11행과 12행 주변 데이터 확인
        if len(df) >= 10:
            행9 = df.iloc[9]
            로그.info(f"[디버그] DataFrame 인덱스 9 (엑셀 11행) 데이터:")
            로그.info(f"  발송여부: '{셀값(행9.get('발송 여부'))}', 이메일: '{셀값(행9.get('이메일'))}'")
            if len(df) >= 11:
                행10 = df.iloc[10]
                로그.info(f"[디버그] DataFrame 인덱스 10 (엑셀 12행) 데이터:")
                로그.info(f"  발송여부: '{셀값(행10.get('발송 여부'))}', 이메일: '{셀값(행10.get('이메일'))}'")
            else:
                로그.warning(f"[디버그] DataFrame에 12행(인덱스 10)이 없습니다! 총 {len(df)}행만 읽혔습니다.")

        목록 = 발송목록_만들기(df, 로그)
        if not 목록:
            안내 = "발송할 메일이 없습니다.\n'발송 여부'가 Y인지 확인하세요."
            if tk is not None:
                try:
                    root = tk.Tk()
                    root.withdraw()
                    messagebox.showinfo("알림", 안내, parent=root)
                    root.destroy()
                except Exception:
                    print("\n" + 안내)
            else:
                print("\n" + 안내)
            로그.info("발송 대상 없음 → 종료")
            return 0
        로그.info(f"발송 대상: {len(목록)}건")

        if not 확인창_띄우기(목록):
            로그.info("사용자 취소")
            return 0

        outlook = Outlook_연결()
        if outlook is None:
            return 1
        로그.info("Outlook 연결됨")

        # ✅ 새 발송 전에 Outbox 대기 확인 (원하면 끌 수도 있음)
        if not 미처리_확인(outlook, 로그):
            로그.info("사용자 선택으로 중단 (미처리 메일 존재)")
            return 0

        성공, 실패목록 = 0, []
        로그.info("=" * 50)
        로그.info("발송 시작")

        총개수 = len(목록)
        for i, 항목 in enumerate(목록, 1):
            로그.info(f"[{i}/{총개수}] {항목['이름']} <{항목['받는사람']}>")

            ok, err = 메일_발송(outlook, 항목, 로고, 로그)

            if ok:
                성공 += 1
                로그.info("  → 완료")
            else:
                실패목록.append({"항목": 항목, "사유": err})
                로그.error(f"  → 실패: {err}")

            if i < 총개수 and SEND_DELAY_SEC > 0:
                time.sleep(SEND_DELAY_SEC)

        print(f"\n{'='*50}")
        print(f"결과: 성공 {성공}건 / 실패 {len(실패목록)}건")
        print(f"{'='*50}")

        if 실패목록:
            print("\n실패 목록:")
            for f in 실패목록:
                print(f"  - 행 {f['항목']['행']}: {f['항목']['이름']} → {f['사유']}")

        로그.info(f"완료: 성공 {성공}, 실패 {len(실패목록)}")

        # ✅ 발송 후 Outbox 잔여 처리: 이제 '비워질 때까지' 체크함
        발송후_미처리_처리(outlook, 로그)

        if 실패목록:
            return 1 if 성공 == 0 else 2
        return 0

    finally:
        logging.shutdown()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\n중단됨")
        sys.exit(130)
    except Exception as e:
        print(f"\n[치명적 오류] {e}")
        sys.exit(1)
