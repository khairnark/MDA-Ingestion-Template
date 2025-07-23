import pandas as pd
import subprocess
import sys
import xlrd

print("✅ Running latest ingestion_template.py")

DATA_DICT_PATH = r"C:\MDM\IngestionScriptingTools\002_Ingestion Script xls to XML Conversion Tool\ADE Data Dictionary\DataDictionary.xlsx"
SHEET_NAME = "Data_Dictionary"

def check_cloneable_by_field():
    try:
        data_dict_df = pd.read_excel(DATA_DICT_PATH, sheet_name=SHEET_NAME)
    except Exception as e:
        print(f"❌ Error loading Excel file: {e}")
        return

    required_columns = ['Field Name (with ID)', 'Cloneable']
    if not all(col in data_dict_df.columns for col in required_columns):
        print(f"❌ Missing required columns in Excel. Required: {required_columns}")
        return

    while True:
        user_input = input("\n🔹 Enter the Field ID (as in 'Field Name (with ID)'): ").strip()
        if not user_input:
            print("⚠️ Field ID cannot be empty. Try again.")
            continue

        match = data_dict_df[data_dict_df['Field Name (with ID)'] == user_input]
        if match.empty:
            print(f"❌ No match found for '{user_input}'. Please check the Field ID and try again.")
            continue

        cloneable_value = match.iloc[0]['Cloneable']
        cloneable_status = str(cloneable_value).strip().lower()

        if cloneable_status == 'yes':
            print(f"✅ The field '{user_input}' is marked **CLONABLE** in the Data_Dictionary.")
        elif cloneable_status == 'no':
            print(f"🚫 The field '{user_input}' is marked **NON-CLONABLE** in the Data_Dictionary.")
        else:
            print(f"⚠️ The field '{user_input}' has an unclear cloneable status: '{cloneable_value}'.")

        # Ask for user confirmation
        print("\n🔸 Confirm the status and choose an action:")
        print("1. Proceed as NON-CLONABLE")
        print("2. Proceed as CLONABLE")
        print("3. Cancel")

        choice = input("Enter 1, 2, or 3: ").strip()
        if choice == '1':
            subprocess.run(['python', 'demo1.py'])
        elif choice == '2':
            subprocess.run(['python', 'newdtest1.py'])
        elif choice == '3':
            print("🔁 Restarting...")
            continue
        else:
            print("❌ Invalid option. Please enter 1, 2, or 3.")
        break

# Run main function
if __name__ == "__main__":
    check_cloneable_by_field()
