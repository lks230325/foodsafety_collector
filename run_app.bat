@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo ====================================================================
echo   [충청·전북 양념육·세절육 가공장 수집기] Streamlit 웹 대시보드 실행
echo   브라우저 접속 주소: http://localhost:8501
echo ====================================================================
python -m streamlit run app.py --server.port 8501 --server.headless false
pause
