import pandas as pd
from itertools import combinations
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
        # Greedy algorithm: pick items from the top until the target is reached
        current_sum = 0
        current_group = []
        indexes_to_remove = []
        for index, row in remaining_df.iterrows():
            qty = row.iloc[2]
            if current_sum + qty <= target:
                current_sum += qty
                current_group.append(row)
                indexes_to_remove.append(index)
        # If no valid group was formed, handle large items by splitting
        if not current_group:
            largest_item = remaining_df.iloc[0]
            new_qty = target if largest_item.iloc[2] > target else largest_item.iloc[2]
            split_df = pd.DataFrame(
                [
                    [largest_item.iloc[0], largest_item.iloc[1], new_qty],
                    [largest_item.iloc[0], largest_item.iloc[1], largest_item.iloc[2] - new_qty],
                ],
                columns=remaining_df.columns,
            )
            remaining_df = pd.concat(
                [remaining_df.iloc[1:], split_df[split_df.iloc[:, 2] > 0]], ignore_index=True
            )
            continue
        # Create a dataframe for this group
        group_df = pd.DataFrame(current_group, columns=remaining_df.columns)
        group_df["Group"] = f"Group {group_count}" # Add a group label
        group_dfs.append(group_df)
        # Remove used items from the remaining dataframe
        remaining_df.drop(index=indexes_to_remove, inplace=True)
        group_count += 1
    # Combine all DataFrames into one, adding a blank row between groups
    combined_df = pd.concat(group_dfs, ignore_index=True)
    # Export the combined DataFrame to a single Excel sheet
    combined_df.to_excel(output_file, index=False, sheet_name="All Groups", engine="openpyxl")
    print(f"Results exported to {output_file}")
    return combined_df