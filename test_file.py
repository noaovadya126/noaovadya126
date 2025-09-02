import pandas as pd

try:
    file_path = r"C:\Users\נועה\Downloads\2017년 국제 통용 한국어 표준 교육과정 적용 연구(4단계) 어휘, 문법 등급 목록_20180227_20201117 수정 (1).xlsx"
    df = pd.read_excel(file_path)
    
    print("File loaded successfully!")
    print(f"Total rows: {len(df)}")
    print(f"Columns: {df.columns.tolist()}")
    print("\nFirst 10 rows:")
    print(df.head(10))
    
    if len(df) > 0:
        print(f"\nFirst column name: {df.columns[0]}")
        print(f"First column sample values:")
        print(df.iloc[:10, 0].tolist())
        
except Exception as e:
    print(f"Error: {e}")
