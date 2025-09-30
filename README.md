## Student Data Cleaning Scripts

### Introduction
This repository contains a set of **Python scripts** I developed to solve a real-world problem faced at **FCT College of Nursing Sciences, Gwagwalada**.  
When working with student records, the **raw exports** from various sources (JAMB, Post-UTME, CAOSCE, and Internal Exams) often come in formats that are **inconsistent, cluttered with unnecessary columns, and not directly usable**.  

To address this, I built a pipeline of scripts that **clean, restructure, and standardize the data** into formats that are ready for either:  
- **Portal Uploads** (for JAMB biodata only).  
- **Internal Use** (Post-UTME, CAOSCE, Internal Exams).  

This project not only automated what was once a tedious manual process, but also reduced errors, improved consistency, and made it possible for anyone in the team to run the process efficiently.

---

### Tech Stack
- **Python 3** → Core programming language.  
- **Pandas** → Data manipulation and transformation.  
- **OpenPyXL** → Excel file handling.  
- **python-dotenv** → For managing environment variables.  
- **WSL (Windows Subsystem for Linux)** → Provides a Linux environment on Windows.  
- **VS Code** → IDE used for development and testing.  

---

### Scripts Overview

### 1. JAMB Candidate Biodata Cleaner
- **Problem**:  
  - Names appear as **one long field** instead of being split into **Surname, Firstname, Othername**.  
  - Raw file contains **extra columns** not needed for processing.  
- **Solution**:  
  - Splits full names into the required fields.  
  - Drops irrelevant columns and **renames the needed ones** to match the portal format.  
- **Purpose**:  
  - Generates a cleaned biodata file that is **ready for upload into the school’s Post-UTME portal**.  

---

### 2. Post-UTME Results Cleaner
- **Problem**:  
  - Raw exports include unnecessary columns and inconsistent headers.  
- **Solution**:  
  - Removes irrelevant columns.  
  - Keeps only the needed ones and **renames them** appropriately.  
- **Purpose**:  
  - Produces a standardized Post-UTME results file for **internal reporting and processing** (not for portal upload).  

---

### 3. CAOSCE Results Cleaner
- **Problem**:  
  - Raw CAOSCE exam results also contain irrelevant columns.  
- **Solution**:  
  - Cleans up the file by dropping extras and renaming required fields.  
- **Purpose**:  
  - Produces a cleaned CAOSCE results file for **internal academic use**.  

---

### 4. Internal Exam Results Cleaner
- **Problem**:  
  - Internal exam exports share the same issue of extra, inconsistent columns.  
- **Solution**:  
  - Standardizes internal results by removing unnecessary fields and renaming useful ones.  
- **Purpose**:  
  - Provides a cleaned results file for **internal record-keeping**.  

---

## Folder Structure

On **Windows**, the setup is organized as follows:


