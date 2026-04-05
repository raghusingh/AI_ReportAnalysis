@echo off
echo Starting AI Dashboard Studio...
echo.

:: Activate virtual environment if it exists
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
    echo Virtual environment activated.
) else (
    echo No venv found, using system Python.
)

:: Run with file watcher disabled to prevent VS Code crashes
streamlit run app.py ^
    --server.fileWatcherType none ^
    --server.headless false ^
    --browser.gatherUsageStats false

pause
