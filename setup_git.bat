@echo off
echo Git 저장소 초기화 및 GitHub 업로드 스크립트
echo.

REM Git 설치 확인
where git >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [오류] Git이 설치되어 있지 않습니다.
    echo Git을 설치해주세요: https://git-scm.com/download/win
    pause
    exit /b 1
)

echo [1/5] Git 저장소 초기화...
git init
if %ERRORLEVEL% NEQ 0 (
    echo [오류] Git 초기화 실패
    pause
    exit /b 1
)

echo [2/5] Python 파일 추가...
git add *.py
if %ERRORLEVEL% NEQ 0 (
    echo [경고] 일부 파일 추가 실패
)

echo [3/5] 설정 파일 추가...
git add .gitignore README.md requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo [경고] 설정 파일 추가 실패
)

echo [4/5] 첫 커밋 생성...
git commit -m "Initial commit: Python scripts collection"
if %ERRORLEVEL% NEQ 0 (
    echo [오류] 커밋 실패
    pause
    exit /b 1
)

echo.
echo [5/5] GitHub 저장소 연결
echo.
echo 다음 단계를 수행하세요:
echo 1. GitHub에서 새 저장소를 생성하세요
echo 2. 아래 명령어를 실행하세요 (저장소 URL을 본인의 것으로 변경):
echo.
echo    git remote add origin https://github.com/사용자명/저장소명.git
echo    git branch -M main
echo    git push -u origin main
echo.
echo 또는 GitHub Desktop을 사용하여 업로드할 수 있습니다.
echo.
pause


