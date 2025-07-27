import pandas as pd
import os

def consolidate_excel(file_path):
    try:
        # Read all sheets into a dictionary
        all_sheets = pd.read_excel(file_path, sheet_name=None)

        # Combine them with a 'Sheet Name' column
        combined_data = []
        for sheet_name, df in all_sheets.items():
            df['Sheet Name'] = sheet_name
            combined_data.append(df)

        # Concatenate all dataframes
        final_df = pd.concat(combined_data, ignore_index=True)

        # Create output filename
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_file = f"{base_name}_consolidated.xlsx"
        final_df.to_excel(output_file, index=False)

        print(f"✅ Consolidated file saved as: {output_file}")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    file_path = input("📂 Enter the path to your Excel file: ")
    if os.path.exists(file_path):
        consolidate_excel(file_path)
    else:
        print("❌ File not found. Please check the path.")
