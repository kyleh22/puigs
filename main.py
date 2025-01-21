# main
import pandas as pd
import os
import functions as funcs
def main():
    while True:
        # User input for file path and target value
        file_name = input("Enter the name of the Excel file (with extension, e.g., 'Order excel sheet.xlsm') in the same directory: ")
        target_value = int(input("Enter the target value for grouping (e.g., 160000): "))
        # Construct the file path
        file_path = os.path.join(os.getcwd(), file_name)
        # Check if the file exists
        if not os.path.isfile(file_path):
            print(f"Error: The file '{file_name}' does not exist in the directory '{os.getcwd()}'.")
            continue
        # Load the Excel file
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            print(f"Error: Unable to load the Excel file. Details: {e}")
            continue
        # Optimize combinations
        try:
            optimized_groups = funcs.find_combinations_export(df, target_value)
            print("Processing complete. The results have been saved to 'output_combinations.xlsx'.")
        except Exception as e:
            print(f"Error during processing. Details: {e}")
            continue
        # Ask user if they want to re-run the script
        rerun = input("Do you want to process another file? (y/n): ").strip().lower()
        if rerun != 'y':
            print("Exiting the program. Goodbye!")
            break
if __name__ == "__main__":
    main()