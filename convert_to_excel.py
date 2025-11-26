
import io
import pandas as pd

# Define input and output file paths for Windows
input_file_path = r"C:\temp\RawGraylogAPIResponse_20250728.txt"
output_file_path = r"C:\temp\RawGraylogAPIResponse_20250728.xlsx"

# Read the content of the input text file
try:
    with open(input_file_path, 'r', encoding='utf-8') as f:
        file_content = f.read()
except FileNotFoundError:
    print(f"Error: Input file not found at {input_file_path}")
    exit()
except Exception as e:
    print(f"Error reading input file: {e}")
    exit()

# Use io.StringIO to treat the file content as a string and read into a DataFrame
try:
    data = io.StringIO(file_content)
    df = pd.read_csv(data)
except Exception as e:
    print(f"Error processing data: {e}")
    exit()

# Save the DataFrame to an Excel file
try:
    df.to_excel(output_file_path, index=False)
    print(f"Successfully converted data to Excel file: {output_file_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")
