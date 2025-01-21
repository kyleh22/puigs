import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment

def find_combinations_export(df, target, output_file="output_combinations_optimized.xlsx"):
    print("Column headings in the dataset:")
    print(f"1st Column: {df.columns[0]}")
    print(f"2nd Column: {df.columns[1]}")
    print(f"3rd Column: {df.columns[2]}")

    remaining_df = df.copy()
    group_count = 1
    group_dfs = []

    while not remaining_df.empty:
        # Sort the dataframe by the 3rd column (Qty) descending for better packing
        remaining_df.sort_values(by=remaining_df.columns[2], ascending=False, inplace=True)

        current_sum = 0
        current_group = []
        indexes_to_remove = []

        for index, row in remaining_df.iterrows():
            qty = row.iloc[2]
            if current_sum + qty <= target:
                current_sum += qty
                current_group.append(row)
                indexes_to_remove.append(index)

        # If we cannot reach the target and have remaining items, split the largest item
        if current_sum < target:
            for index, row in remaining_df.iterrows():
                qty = row.iloc[2]
                if current_sum + qty > target:
                    # Split the item to fulfill the target
                    split_amount = target - current_sum
                    current_group.append(pd.Series([row.iloc[0], row.iloc[1], split_amount], index=remaining_df.columns))
                    remaining_df.at[index, remaining_df.columns[2]] -= split_amount
                    current_sum = target
                    break

        # Create a dataframe for this group
        group_df = pd.DataFrame(current_group, columns=remaining_df.columns)
        group_df["Group"] = f"Container {group_count}"  # Add a group label

        # Add a row for the sum total of the group
        total_row = {col: None for col in group_df.columns}
        total_row[df.columns[2]] = group_df.iloc[:, 2].sum()
        total_row["Group"] = f"Total for Container {group_count}"
        group_df = pd.concat([group_df, pd.DataFrame([total_row])], ignore_index=True)

        group_dfs.append(group_df)

        # Remove used items from the remaining dataframe
        remaining_df.drop(index=indexes_to_remove, inplace=True)
        remaining_df = remaining_df[remaining_df.iloc[:, 2] > 0]  # Remove any fully used rows
        group_count += 1

    # Add remaining items to the final group (if any)
    if not remaining_df.empty:
        final_group = remaining_df.copy()
        final_group["Group"] = f"Container {group_count}"

        # Add a total row for the final group
        total_row = {col: None for col in final_group.columns}
        total_row[df.columns[2]] = final_group.iloc[:, 2].sum()
        total_row["Group"] = f"Total for Container {group_count}"
        final_group = pd.concat([final_group, pd.DataFrame([total_row])], ignore_index=True)

        group_dfs.append(final_group)

    # Combine all groups into a single DataFrame with titles and blank rows
    formatted_dfs = []
    for i, group_df in enumerate(group_dfs):
        # Add a title row for the group
        title_row = pd.DataFrame([[None, None, None, f"Container {i + 1}"]], columns=group_df.columns)
        # Append title, group, and an empty row
        formatted_dfs.extend([title_row, group_df, pd.DataFrame(columns=group_df.columns)])

    combined_df = pd.concat(formatted_dfs, ignore_index=True)

    return combined_df

# Assuming `df` is your dataframe
def clean_group_column(df):
    # Replace 'Group' column values with an empty string where 'Qty' is None

    # print('\nBefore cleaning\n')
    # print(df)
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

    # Get column indices for "Qty" and "Group"
    qty_col_idx = None
    group_col_idx = None
    for idx, cell in enumerate(ws[1], start=1):  # Assuming the first row contains headers
        if cell.value == "Qty":
            qty_col_idx = idx
        elif cell.value == "Group":
            group_col_idx = idx
        if qty_col_idx and group_col_idx:
            break

    if qty_col_idx is None or group_col_idx is None:
        print("Could not find 'Qty' or 'Group' column in the Excel file.")
        return

    # Apply styles
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        is_total_row = row[group_col_idx - 1].value and "Total for Container" in str(row[group_col_idx - 1].value)
        is_title_row = row[group_col_idx - 1].value and "Container" in str(row[group_col_idx - 1].value) and "Total" not in str(row[group_col_idx - 1].value)

        for idx, cell in enumerate(row, start=1):
            if idx == qty_col_idx or idx == group_col_idx:  # Only highlight the Qty and Group columns
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