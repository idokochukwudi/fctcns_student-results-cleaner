#!/bin/bash
cd /home/ernest/student_result_cleaner
source venv/bin/activate
exec gunicorn --workers 2 --bind 127.0.0.1:5000 --timeout 60 --log-level debug launcher.app:app