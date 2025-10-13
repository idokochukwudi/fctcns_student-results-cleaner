#!/bin/bash
echo "=== Creating PythonAnywhere Deployment Package ==="

# Create deployment directory
DEPLOY_DIR="pythonanywhere_deploy"
mkdir -p $DEPLOY_DIR
cd $DEPLOY_DIR

echo "Copying application files..."
# Copy your main app file
cp ../app.py ./app_pa.py
# Copy requirements
cp ../requirements.txt .
# Copy environment file (remove if contains sensitive local paths)
cp ../.env .
# Copy essential directories
cp -r ../scripts .
cp -r ../templates .
cp -r ../img .

echo "Creating PythonAnywhere directory structure..."
# Create data directories
mkdir -p data/PUTME_RESULT/RAW_PUTME_RESULT
mkdir -p data/CAOSCE_RESULT/RAW_CAOSCE_RESULT  
mkdir -p data/INTERNAL_RESULT/RAW_INTERNAL_RESULT
mkdir -p data/JAMB_DB/RAW_JAMB_DB
mkdir -p data/EXAMS_INTERNAL

echo "Creating WSGI file..."
cat > wsgi.py << 'WSGI_EOF'
import sys
import os

# Add your app directory to Python path
path = os.path.dirname(os.path.abspath(__file__))
if path not in sys.path:
    sys.path.insert(0, path)

from app_pa import app as application
WSGI_EOF

echo "Creating PythonAnywhere optimized app file..."
cat > app_pa.py << 'APP_EOF'
import os
import subprocess
import re
from flask import Flask, request, redirect, url_for, render_template, flash, session
from functools import wraps
from dotenv import load_dotenv

app = Flask(__name__)

# PythonAnywhere specific configuration
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
print(f"PythonAnywhere: Base directory is {BASE_DIR}")

# Use relative paths for PythonAnywhere
PATHS = {
    "utme": os.path.join(BASE_DIR, "data", "PUTME_RESULT", "RAW_PUTME_RESULT"),
    "caosce": os.path.join(BASE_DIR, "data", "CAOSCE_RESULT", "RAW_CAOSCE_RESULT"), 
    "clean": os.path.join(BASE_DIR, "data", "INTERNAL_RESULT", "RAW_INTERNAL_RESULT"),
    "split": os.path.join(BASE_DIR, "data", "JAMB_DB", "RAW_JAMB_DB"),
    "exam_processor": os.path.join(BASE_DIR, "data", "EXAMS_INTERNAL")
}

# Ensure directories exist
def setup_directories():
    for path_name, path in PATHS.items():
        os.makedirs(path, exist_ok=True)
        print(f"PythonAnywhere: Directory ready - {path_name}: {path}")

# Load environment and setup
load_dotenv()
app.secret_key = os.getenv("FLASK_SECRET", "pythonanywhere_secret_123")
setup_directories()

# Your existing configuration would go here...
# [PASTE YOUR ENTIRE APP.PY CONTENT AFTER THIS LINE]
APP_EOF

# Append your actual app.py content to app_pa.py
echo "Appending your application code..."
cat ../app.py >> app_pa.py

echo "=== Deployment package created in: $DEPLOY_DIR ==="
echo ""
echo "=== NEXT STEPS ==="
echo "1. ZIP the deployment folder:"
echo "   zip -r pythonanywhere_deploy.zip pythonanywhere_deploy/"
echo "2. Upload to PythonAnywhere"
echo "3. Follow setup instructions in PythonAnywhere"
