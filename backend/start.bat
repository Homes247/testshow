@echo off
echo Starting Office Suite Backend...

if not exist .venv (
    echo Creating virtual environment...
    python -m venv .venv
)

call .venv\Scripts\activate.bat

echo Installing dependencies...
pip install -r requirements.txt -q

echo.
echo Backend running at http://192.168.0.187:8000
echo WebSocket live sync at ws://192.168.0.187:8000/ws/
echo Press Ctrl+C to stop.
echo.

python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
