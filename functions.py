import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment

import pandas as pd

def find_combinations_export(df, target, output_file="output_combinations_optimized.xlsx"):
    # Set the first row as column headers
    original_headings = df.iloc[0]  # Store the first row for later re-use
    df = df.rename(columns=dict(zip(df.columns, original_headings)))

    # Remove the row used for headings and reset the index
    df = df[1:].reset_index(drop=True)
    df = df[1:].reset_index(drop=True)  # Skip the second metadata row

    # Convert the "Qty" column to numeric
    df.iloc[:, 2] = pd.to_numeric(df.iloc[:, 2], errors="coerce")

    # Extract numeric values from the "Leadtime Weeks" column and add a new column for numeric lead times
    leadtime_column = "Leadtime Weeks"
    df[leadtime_column] = df[leadtime_column].str.extract(r"(\d+)").astype(float)

    # Sort by Leadtime Weeks (ascending) and Qty (descending)
    df.sort_values(by=[leadtime_column, df.columns[2]], ascending=[True, False], inplace=True)

    remaining_df = df.copy()
    group_count = 1
    group_dfs = []

    while not remaining_df.empty:
        # Sort by "Leadtime Weeks" (shortest first) and "Qty" (largest first)
        remaining_df.sort_values(by=[leadtime_column, remaining_df.columns[2]], ascending=[True, False], inplace=True)

        current_sum = 0
        current_group = []
        indexes_to_remove = []

        for index, row in remaining_df.iterrows():
            qty = row.iloc[2]
            if current_sum + qty <= target:
                current_sum += qty
                current_group.append(row)
                indexes_to_remove.append(index)

        # Handle remaining items that need splitting, but only if not the last container
        if current_sum < target and len(remaining_df) > len(indexes_to_remove):
            for index, row in remaining_df.iterrows():
                if index not in indexes_to_remove:
                    qty = row.iloc[2]
                    if current_sum + qty > target:
                        split_amount = target - current_sum

                        # Create a copy of the row and update the quantity column
                        split_row = row.copy()
                        split_row.iloc[2] = split_amount
                        current_group.append(split_row)

                        # Update the original row in remaining_df
                        remaining_df.at[index, remaining_df.columns[2]] -= split_amount
                        current_sum = target
                        break

        # Create a DataFrame for the current group
        group_df = pd.DataFrame(current_group, columns=remaining_df.columns)
        group_df["Group"] = f"Container {group_count}"

        # Add a total row for the group
        total_row = {col: None for col in group_df.columns}
        total_row[df.columns[2]] = group_df.iloc[:, 2].sum()
        total_row["Group"] = f"Total for Container {group_count}"
        group_df = pd.concat([group_df, pd.DataFrame([total_row])], ignore_index=True)

        group_dfs.append(group_df)

        # Remove used items from remaining_df
        remaining_df.drop(index=indexes_to_remove, inplace=True)
        remaining_df = remaining_df[remaining_df.iloc[:, 2] > 0]  # Remove rows with zero quantity

        # Break if this is the last container
        if remaining_df.empty:
            break

        group_count += 1

    # Handle the last container without worrying about the target
    if not remaining_df.empty:
        final_group = remaining_df.copy()
        final_group["Group"] = f"Container {group_count}"

        # Add a total row for the final group
        total_row = {col: None for col in final_group.columns}
        total_row[df.columns[2]] = final_group.iloc[:, 2].sum()
        total_row["Group"] = f"Total for Container {group_count}"
        final_group = pd.concat([final_group, pd.DataFrame([total_row])], ignore_index=True)

        group_dfs.append(final_group)

    # Combine all groups into a single DataFrame
    formatted_dfs = []
    for i, group_df in enumerate(group_dfs):
        title_row = pd.DataFrame([[None] * (len(group_df.columns) - 1) + [f"Container {i + 1}"]], columns=group_df.columns)
        formatted_dfs.extend([title_row, group_df, pd.DataFrame(columns=group_df.columns)])

    combined_df = pd.concat(formatted_dfs, ignore_index=True)

    return combined_df


# Assuming `df` is your dataframe
def clean_group_column(df):
    # Replace 'Group' column values with an empty string where 'Qty' is None
    df.loc[df['Qty'].isna(), 'Group'] = ''
    # print('\nAfter cleaning\n')
    # print(df)
    return df

def export_df_to_excel(df, output_file="output_combinations_optimized.xlsx"):
    # Export to Excel
    df.to_excel(output_file, index=False, sheet_name="All Groups", engine="openpyxl")

    # Style the Excel file
    wb = load_workbook(output_file)
    ws = wb["All Groups"]

    # Define styles
    highlight_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    border_style = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    center_alignment = Alignment(horizontal="center", vertical="center")

    # Get column indices for "Qty" and the second column
    qty_col_idx = None
    group_col_idx = 2  # Second column index (1-based)

    # Check if the second column contains "ORDER"
    second_col_title = ws.cell(row=1, column=group_col_idx).value
    if not second_col_title or "ORDER" not in second_col_title.upper():
        print(f"The second column does not contain 'ORDER'. Found: '{second_col_title}'")
        return

    # Find "Qty" column
    for idx, cell in enumerate(ws[1], start=1):  # Assuming the first row contains headers
        if cell.value == "Qty":
            qty_col_idx = idx
        if qty_col_idx:
            break

    if qty_col_idx is None:
        print("Could not find 'Qty' column in the Excel file.")
        return

    # Apply styles
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        is_total_row = row[group_col_idx - 1].value and "Total for Container" in str(row[group_col_idx - 1].value)
        is_title_row = row[group_col_idx - 1].value and "Container" in str(row[group_col_idx - 1].value) and "Total" not in str(row[group_col_idx - 1].value)

        for idx, cell in enumerate(row, start=1):
            if idx == qty_col_idx or idx == group_col_idx:  # Only highlight the Qty and second column
                if is_total_row:
                    cell.fill = highlight_fill
            if cell.value is not None:  # Apply borders to non-empty cells
                cell.border = border_style
            if is_title_row:  # Center-align title rows
                cell.alignment = center_alignment

    # Save the updated workbook
    wb.save(output_file)

    print(f"Results exported to {output_file}")
    return
