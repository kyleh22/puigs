import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

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
        group_df["Group"] = f"Group {group_count}"  # Add a group label
        
        # Add a row for the sum total of the group
        total_row = {col: None for col in group_df.columns}
        total_row[df.columns[2]] = group_df.iloc[:, 2].sum()
        total_row["Group"] = f"Total Group {group_count}"
        group_df = pd.concat([group_df, pd.DataFrame([total_row])], ignore_index=True)
        
        group_dfs.append(group_df)
        
        # Remove used items from the remaining dataframe
        remaining_df.drop(index=indexes_to_remove, inplace=True)
        remaining_df = remaining_df[remaining_df.iloc[:, 2] > 0]  # Remove any fully used rows
        group_count += 1
    
    # Add remaining items to the final group (if any)
    if not remaining_df.empty:
        final_group = remaining_df.copy()
        final_group["Group"] = f"Group {group_count}"
        
        # Add a total row for the final group
        total_row = {col: None for col in final_group.columns}
        total_row[df.columns[2]] = final_group.iloc[:, 2].sum()
        total_row["Group"] = f"Total Group {group_count}"
        final_group = pd.concat([final_group, pd.DataFrame([total_row])], ignore_index=True)
        
        group_dfs.append(final_group)
    
    # Combine all groups into a single DataFrame
    combined_df = pd.concat(group_dfs, ignore_index=True)
    
    # Export to Excel
    combined_df.to_excel(output_file, index=False, sheet_name="All Groups", engine="openpyxl")
    
    # Highlight the total rows in the Excel file
    wb = load_workbook(output_file)
    ws = wb["All Groups"]
    
    # Define the highlight fill
    highlight_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=False):
        if row[-1].value and "Total Group" in str(row[-1].value):
            for cell in row:
                cell.fill = highlight_fill  # Apply the highlight to all cells in the row
    
    # Save the updated workbook
    wb.save(output_file)
    
    print(f"Results exported to {output_file}")
    return combined_df
