@echo off
echo Setting up ILS Kickoff Tool...
python -m venv .venv
.venv\Scripts\pip install --upgrade pip
.venv\Scripts\pip install -r requirements.txt
echo.
echo Setup complete! Run "run.bat" to start the analysis.
pause
