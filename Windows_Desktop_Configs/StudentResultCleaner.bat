@echo off
:: Activate the WSL virtual environment and run the launcher
wsl bash -c "source /home/ernest/student_result_cleaner/venv/bin/activate && python3 /home/ernest/student_result_cleaner/cli_launcher.py"
pause
