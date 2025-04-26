README for Excel File Grouping and Optimization Software

Overview

This software processes Excel files containing inventory or order data. It organizes items into groups based on a specified target quantity and exports the results to a new Excel file with clear formatting for easy review.

Prerequisites

Before using this software, ensure the following are installed:

Python: Version 3.7 or higher.
Mac: Most macOS versions come with Python pre-installed, but it may be outdated. Install the latest version from Python.org or via Homebrew:

brew install python

Required Libraries:
Install the necessary Python libraries:

pandas
openpyxl

Open the Terminal app (press Cmd + Space, type "Terminal", and hit Enter) and run:

pip3 install pandas openpyxl

Microsoft: 

1. Go to the python.org website to download it. When installing, ensure to select 'add PATH' upon installation
2. Check to make sure python is installed by opening terminal (search 'cmd' on windows and then select the 'Command Prompt' application)
3. Copy and paste 'python --version' in the window and press enter. If you 'Python 3.13.3' or another number, it has been sucessfully installed
4. Now copy and paste this in the terminal 'curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py'
5. Once that is installed, copy and paste 'python get-pip.py'
6. Now pip should be installed (the tool used to install other packages). You can check it's installed by typing 'pip --version'
7. Now copy and paste 'pip install pandas'
8. and also 'pip install openpyxl'
9. Now the application should work

Excel File: Ensure your Excel file meets the formatting requirements (see below) and is saved in the same folder as the script.

How to Use
Place Your Excel File
Save the Excel file you want to process in the same directory as the main.py script. For example, if the script is saved in the Documents folder, place your Excel file there too.

Run the Script
Open Terminal, navigate to the folder containing main.py using the cd command, and run the script:

cd ~/Documents  # Replace with the folder where your script is located
python3 main.py
Provide Inputs

Enter the file name of your Excel sheet (e.g., Order excel sheet.xlsm).
Enter the target value for grouping (e.g., 160000).
Review the Results
After processing, the script will generate an output file named output_combinations.xlsx in the same folder. Open it using Excel or Numbers to review the grouped data.

Excel File Formatting Requirements
To ensure smooth operation, the input Excel file must follow these rules:

File Assumptions:

The file must have at least three columns:
Column 1: Product details (e.g., Item Number, Description).
Column 2: Item heading (used for reference but not calculations).
Column 3: Quantity (Qty) for each product (must contain numeric values).
Column for Lead Time: Must be labeled 'Leadtime Weeks' and contain numeric or text values like "2 weeks".
The first row is assumed to contain the PO, like PO012025DD, then the second row contains the column headers like 'Item', 'LUCKY STAR - ORDER', 'Qty', etc

Formatting Details:

Ensure numeric columns like Qty contain only numbers. Non-numeric values (like letters or words) will cause errors.
The Leadtime Weeks column must contain either:
Numeric values (e.g., 2 or 5).
Text values with embedded numbers (e.g., 5 weeks or 10 weeks).
Avoid merged cells, empty rows, or blank cells in key columns.

Additional Notes:

Avoid extra rows or columns outside the main dataset.
Do not include multiple sheets in the Excel file; only the first sheet will be processed.

How It Works

1. User specifies Excel file name
2. User specifies Qty
3. Script Loads the Excel File
4. The script reads your Excel file into a Python DataFrame.
5. Process the Data
6. Items are sorted by Leadtime Weeks (ascending) and Qty (descending).
7. Items are grouped into containers to ensure each group’s total Qty matches the target value except for the last container.
8. Remaining quantities are split into new groups as necessary.
9. Export the Results
10. The processed data is saved to a new Excel file, output_combinations.xlsx, with:

Clear container titles and totals.
Groups organized sequentially in a single sheet.

Troubleshooting

If you encounter issues, refer to the common errors below:

1. "The file does not exist in the directory"
Cause: The file name entered is incorrect, or the file is not in the same directory as main.py.
Fix:
Double-check the file name, including the extension (e.g., .xlsm or .xlsx).
Make sure the file is in the same directory as the script. Use the ls command in Terminal to list the files in the current folder.
2. "Unable to load the Excel file"
Cause: The file is corrupted or not a valid Excel format.
Fix:
Open the file in Excel or Numbers to verify it is valid.
Save it again as .xlsx or .xlsm.
3. "Error during processing"
Cause: Common causes include:
Non-numeric values in the Qty column.
Missing or incorrectly labeled columns.
Fix:
Ensure the headers match the expected format.
Verify that the Qty column contains only numbers.
4. "Could not find 'Qty' column in the Excel file"
Cause: The program cannot locate the Qty column.
Fix:
Ensure the column is labeled exactly as Qty (case-sensitive).

Assumptions and Limitations
1. The script assumes that the column headers are correctly labeled and in English.
2. If the Leadtime Weeks column contains text, the program extracts numeric values (e.g., "5 weeks" becomes 5).
3. The program only processes the first sheet in the Excel file.
4. Non-standard formats, extra rows, or merged cells may cause unexpected results.

Expected Output

The output file (output_combinations.xlsx) contains:

1. Clearly labeled groups, such as "Container 1," "Container 2," etc.
2. A single worksheet with all grouped data.

Mac-Specific Tips:
If python3 is not installed, install it using Homebrew:

brew install python

If you face issues running the script:

Ensure you’re in the correct directory using cd.
Use ls to verify the script and Excel file are in the same folder.
Open the output file using Microsoft Excel or Numbers. If formatting appears off in Numbers, use Excel for the best experience.

Contact for Support
If you encounter issues not covered here, feel free to reach out to Kyle Harrison (kyleh8100@gmail.com or +66819122803 on whatsapp)
