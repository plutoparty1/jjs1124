@echo off
chcp 65001 >nul
echo ========================================
echo Python 파일 GitHub 업로드 스크립트
echo ========================================
echo.

REM Git 설치 확인
where git >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [오류] Git이 설치되어 있지 않습니다.
    echo.
    echo Git 설치 방법:
    echo 1. https://git-scm.com/download/win 에서 Git for Windows 다운로드
    echo 2. 설치 후 이 스크립트를 다시 실행하세요
    echo.
    pause
    exit /b 1
)

echo [1/6] Git 버전 확인...
git --version
echo.

REM 이미 Git 저장소인지 확인
if exist .git (
    echo [2/6] 기존 Git 저장소 발견
) else (
    echo [2/6] Git 저장소 초기화...
    git init
    if %ERRORLEVEL% NEQ 0 (
        echo [오류] Git 초기화 실패
        pause
        exit /b 1
    )
)

echo.
echo [3/6] Python 파일 추가...
git add always_on_top.py
git add apply_btn.py
git add apply_white_fill.py
git add diumsong_filter_final.py
git add ERP.py
git add excel_copy.py
git add excel_unmerge.py
git add "JOOS#_List.py"
git add Performance_Royalties.py
git add send_mail.py

echo.
echo [4/6] 설정 파일 추가...
git add .gitignore README.md requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo [경고] 일부 설정 파일이 없을 수 있습니다
)

echo.
echo [5/6] 커밋 생성...
git commit -m "Initial commit: Python scripts collection (10 files)"
if %ERRORLEVEL% NEQ 0 (
    echo [경고] 커밋 실패 또는 변경사항 없음
)

echo.
echo [6/6] 완료!
echo.
echo ========================================
echo 다음 단계: GitHub에 업로드
echo ========================================
echo.
echo 1. GitHub에서 새 저장소를 생성하세요: https://github.com/new
echo.
echo 2. 아래 명령어를 실행하세요 (저장소 URL을 본인의 것으로 변경):
echo.
echo    git remote add origin https://github.com/사용자명/저장소명.git
echo    git branch -M main
echo    git push -u origin main
echo.
echo 또는 GitHub Desktop을 사용하여 업로드할 수 있습니다.
echo.
pause


