#!/bin/bash
echo "Fixing script permissions..."
cd ~/student_result_cleaner

# Make all scripts executable
chmod +x scripts/*.py

# Verify permissions
echo "Current permissions:"
ls -la scripts/

# Verify each script is executable
echo "Checking executability:"
for script in scripts/*.py; do
    if [ -x "$script" ]; then
        echo "✅ $script - Executable"
    else
        echo "❌ $script - NOT Executable"
    fi
done

echo "Permissions fixed!"