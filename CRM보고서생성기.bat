@echo off
chcp 65001 > nul
echo ==============================
echo  CRM 매출분석 보고서 생성기
echo ==============================
echo.

cd /d "%~dp0"

where streamlit >nul 2>&1
if %errorlevel% neq 0 (
    echo [설치] 필요한 패키지를 설치합니다...
    pip install -r requirements.txt
    echo.
)

echo [실행] 브라우저에서 자동으로 열립니다...
echo        열리지 않으면 http://localhost:8501 접속
echo.
streamlit run app.py --server.headless false

pause
