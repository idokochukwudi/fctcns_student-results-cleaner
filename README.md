# Running the `clean_results.py` Script in WSL

This documentation provides a **step-by-step guide** for preparing folders, running the script, and managing cleaned results.

---

### 1. Prerequisites

Before running the script, ensure the following:

- **Windows Subsystem for Linux (WSL)** installed on your machine.
- **Python 3** installed in WSL.
- Required Python libraries installed:

```bash
pip3 install pandas openpyxl
```

### 2. Folder Setup in Windows

1. Open File Explorer and navigate to your Documents folder.
2. Create two folders:

- **RAW_RESULTS**→ Folder where you will copy all raw CSV/Excel files.
    ```makefile
    C:\Users\MTECH COMPUTERS\Documents\RAW_RESULTS
    ```
- **CLEANED_RESULTS** → The script will automatically create timestamped subfolders here for cleaned files.
    ```makefile
    C:\Users\MTECH COMPUTERS\Documents\CLEANED_RESULTS
    ```
### 3. Copy Raw Files
1. Copy all raw CSV or Excel files into:

    ```makefile
    C:\Users\MTECH COMPUTERS\Documents\RAW_RESULTS
    ```
2. Ensure each file has the required columns:
- Surname
- First name
- Grade/20.00
> Files missing these columns will be skipped by the script.

### 4. Verify Script Location

1. Place the script clean_results.py in your WSL home directory or any preferred directory:
    ```bash
    /home/ernest/clean_results.py
    ```
2. Ensure the script has execute permissions:

    ```bash
    chmod +x ~/clean_results.py
    ```
### 5. Script Paths Configuration

The script is pre-configured to:

- **Read raw files from RAW_RESULTS:**
    ```makefile
    /mnt/c/Users/MTECH COMPUTERS/Documents/RAW_RESULTS
    ```
- **Save cleaned files to timestamped folders in CLEANED_RESULTS:**

    ```makefile
    /mnt/c/Users/MTECH COMPUTERS/Documents/CLEANED_RESULTS/obj_result_YYYY-MM-DD_HH-MM-SS
    ```
> No additional changes are needed.

### 6. Run the Script in WSL

1. Open your WSL terminal.
2. Run the script:

```bash
python3 ~/clean_results.py
```

3. The script will:

- Loop through all CSV and Excel files in RAW_RESULTS.
- Clean data by keeping only the required columns.
- Sort rows by MAT NO. (A-Z).
- Add Serial Number (SN) for all rows except Overall Average (SN left blank).
- Save individual cleaned files and a master CSV in a timestamped folder in both WSL and Windows.

### 7. Output Structure

After running the script:
```php-template
CLEANED_RESULTS/
└── obj_result_YYYY-MM-DD_HH-MM-SS/
    ├── cleaned_<file1>.csv
    ├── cleaned_<file2>.csv
    └── master_cleaned_results.csv
```
- Individual cleaned CSVs for each raw file.
- Master CSV combining all cleaned files.
- Overall Average row: SN column is blank, other columns unchanged.

### 8. Verify Windows Output

1. Open File Explorer and navigate to:

    ```makefile
    C:\Users\MTECH COMPUTERS\Documents\CLEANED_RESULTS
    ```
2. Open the timestamped folder created by the script.
3. Check that:
   - All cleaned files are present.
   - Data is sorted by MAT NO..
   - Serial Numbers are correct, except for Overall Average rows.
### 9. Optional: Re-run with New Data
1. Copy new raw CSV/Excel files to RAW_RESULTS.
2. Run the script again:

    ```bash
    python3 ~/clean_results.py
    ```
3. The script will create a new timestamped folder in CLEANED_RESULTS automatically.
    >This ensures previous cleaned files are not overwritten.
### 10. Troubleshooting

| Problem                                   | Solution                                                                                  |
|-------------------------------------------|------------------------------------------------------------------------------------------|
| Script says No CSV or Excel files found    | Make sure files are in RAW_RESULTS and not temporary files starting with `~$`            |
| Permission denied saving to Windows folder| Make sure CLEANED_RESULTS exists and you have write access                                |
| Script skips a file                        | File does not contain required columns: Surname, First name, Grade/20.00                 |
| Python error                               | Make sure pandas and openpyxl are installed: `pip3 install pandas openpyxl`              |

