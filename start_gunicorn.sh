#!/bin/bash
cd /home/ernest/student_result_cleaner/launcher
source ../venv/bin/activate
exec gunicorn --workers 2 --bind 0.0.0.0:5000 --timeout 300 app:app