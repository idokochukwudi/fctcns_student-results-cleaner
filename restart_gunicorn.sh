#!/bin/bash

# Navigate to the project directory
cd /home/ernest/student_result_cleaner/launcher

# Source the virtual environment
source ../venv/bin/activate

# Find and kill the running Gunicorn process
# This looks for the Gunicorn process bound to port 5000
echo "Stopping Gunicorn..."
pkill -f "gunicorn.*0.0.0.0:5000"

# Wait briefly to ensure the process is terminated
sleep 2

# Start Gunicorn again
echo "Starting Gunicorn..."
exec gunicorn --workers 2 --bind 0.0.0.0:5000 --timeout 300 app:app