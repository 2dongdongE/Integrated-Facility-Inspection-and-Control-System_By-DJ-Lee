@echo off
title 시설관리 시스템 통합 실행기
cd /d "%~dp0"

echo ==============================================
echo [1/2] 시설관리 웹서버(Streamlit) 실행 중...
echo ==============================================
if exist ".venv\Scripts\activate.bat" (
    start "Streamlit Server" cmd /k ".venv\Scripts\activate.bat && python -m streamlit run app.py --server.port 8501"
) else (
    start "Streamlit Server" cmd /k "python -m streamlit run app.py --server.port 8501"
)

timeout /t 3 /nobreak > nul

echo ==============================================
echo [2/2] 인터넷 연결 터널(Cloudflare) 실행 중...
echo ==============================================
start "Cloudflare Tunnel" cmd /k ".\cloudflared.exe tunnel --url http://localhost:8501"