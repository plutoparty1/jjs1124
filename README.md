# Python Scripts Collection

이 저장소는 다양한 업무 자동화를 위한 Python 스크립트 모음입니다.

## 주요 스크립트

### 1. excel_copy.py
- **기능**: 엑셀 파일의 파일명에서 연/월을 추출하여 다음 달로 변경한 사본 생성
- **지원 형식**: 
  - `YY년MM월`, `YYYY년MM월`
  - `YY.MM`, `YYYY.MM`
  - `YYMM`, `YYYYMM`
- **특징**: 
  - 드래그 앤 드롭으로 간편하게 사용
  - .xls 파일을 .xlsx로 자동 변환
  - 원본 파일의 공백 형식 유지

### 2. send_mail.py
- **기능**: Outlook을 사용한 자동 메일 발송

### 3. apply_white_fill.py
- **기능**: 엑셀 파일에 흰색 배경 적용

### 4. excel_unmerge.py
- **기능**: 엑셀 파일의 병합된 셀 해제

### 5. diumsong_filter_final.py
- **기능**: 엑셀 데이터 필터링 및 처리

### 6. Performance_Royalties.py
- **기능**: 공연료 관련 처리

### 7. ERP.py
- **기능**: ERP 시스템 연동

### 8. apply_btn.py
- **기능**: 버튼 적용 관련 처리

### 9. always_on_top.py
- **기능**: 항상 위에 표시되는 창 관리

### 10. JOOS#_List.py
- **기능**: JOOS 리스트 처리

## 설치 방법

```bash
# 가상환경 생성 (선택사항)
python -m venv venv

# 가상환경 활성화
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# 필요한 패키지 설치
pip install -r requirements.txt
```

## 필수 패키지

- `pywin32`: Windows COM 객체 사용
- `tkinterdnd2`: 드래그 앤 드롭 지원
- `openpyxl`: 엑셀 파일 처리
- `pymupdf`: PDF 처리

## 사용 방법

각 스크립트는 독립적으로 실행 가능합니다:

```bash
python excel_copy.py
```

## 주의사항

- 일부 스크립트는 Windows 환경에서만 동작합니다
- Excel COM 객체를 사용하는 스크립트는 Microsoft Excel이 설치되어 있어야 합니다
- 개인 데이터가 포함된 파일(.xlsx, .xls 등)은 .gitignore에 포함되어 있습니다

## 라이선스

개인 사용 목적의 스크립트입니다.

