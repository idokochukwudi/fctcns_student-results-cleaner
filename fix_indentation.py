import re

# Read the file
with open('/home/ernest/student_result_cleaner/launcher/app.py', 'r') as f:
    lines = f.readlines()

# Find and fix the problematic section
new_lines = []
i = 0
while i < len(lines):
    line = lines[i]
    
    # Look for the pattern: return statement followed by indented code
    if 'return redirect(url_for("dashboard"))' in line and i + 1 < len(lines):
        next_line = lines[i + 1]
        # If next line has significant indentation after a return, it's wrong
        if next_line.strip() and next_line.startswith('        ') and not next_line.strip().startswith('#'):
            print(f"Found problematic code after return statement at line {i+1}")
            # Skip the improperly indented block until we find a properly indented line
            i += 1
            while i < len(lines) and lines[i].startswith('        ') and lines[i].strip():
                print(f"Removing line: {lines[i].strip()}")
                i += 1
            continue
    
    new_lines.append(line)
    i += 1

# Write the fixed file
with open('/home/ernest/student_result_cleaner/launcher/app.py', 'w') as f:
    f.writelines(new_lines)

print("Fixed indentation issues!")
