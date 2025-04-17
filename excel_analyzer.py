import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

def analyze_excel(file_path):
    """
    Analyze an Excel file and extract important information.
    
    Args:
        file_path (str): Path to the Excel file
    """
    try:
        # Read the Excel file with header=None to see raw data
        df = pd.read_excel(file_path, header=None)
        
        # Display the first few rows of raw data
        print("\n=== First 10 rows of data ===")
        print(df.head(10))
        
        # Display basic information about the dataset
        print("\n=== Basic Information ===")
        print(f"Number of rows: {len(df)}")
        print(f"Number of columns: {len(df.columns)}")
        
        # Display non-null values in each column
        print("\n=== Sample of non-null values in each column ===")
        for col in df.columns:
            non_null_values = df[col].dropna().unique()
            if len(non_null_values) > 0:
                print(f"\nColumn {col}:")
                print(non_null_values[:5])  # Show first 5 unique non-null values
        
        return df
    
    except Exception as e:
        print(f"Error analyzing Excel file: {str(e)}")
        return None

if __name__ == "__main__":
    # Excel file in the current directory
    excel_file = "Demat Holding Query Stmt_2430_17-04-25 03.11.XLS"
    
    # Analyze the Excel file
    df = analyze_excel(excel_file)
    
    if df is not None:
        print("\nAnalysis completed successfully!") 