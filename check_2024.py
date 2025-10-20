import pandas as pd
import os

def check_nd_files():
    base_dir = "EXAMS_INTERNAL/ND"
    
    for nd_set in ["ND-2024", "ND-2025"]:
        print(f"\n{'='*50}")
        print(f"CHECKING {nd_set}")
        print(f"{'='*50}")
        
        raw_dir = f"{base_dir}/{nd_set}/RAW_RESULTS"
        
        if os.path.exists(raw_dir):
            files = [f for f in os.listdir(raw_dir) if f.endswith('.xlsx')]
            
            for file in files:
                file_path = os.path.join(raw_dir, file)
                print(f"\n📁 File: {file}")
                
                try:
                    # Check sheets
                    xl = pd.ExcelFile(file_path)
                    print(f"   Sheets: {xl.sheet_names}")
                    
                    # Read first sheet
                    df = pd.read_excel(file_path)
                    print(f"   Shape: {df.shape}")
                    print(f"   Columns: {list(df.columns)}")
                    
                    # Check for score columns
                    score_columns = [col for col in df.columns if any(keyword in str(col).lower() 
                                    for keyword in ['score', 'mark', 'total', 'grade'])]
                    print(f"   Possible score columns: {score_columns}")
                    
                    # Show sample data
                    if not df.empty:
                        print(f"   First student scores:")
                        for col in score_columns[:3]:  # Show first 3 score columns
                            if col in df.columns:
                                print(f"     {col}: {df[col].iloc[0] if len(df) > 0 else 'N/A'}")
                    
                except Exception as e:
                    print(f"   ❌ ERROR: {e}")

if __name__ == "__main__":
    check_nd_files()