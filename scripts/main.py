import pandas as pd
import os


def extract_chemistry_data(input_file, output_file):
    if not os.path.exists(input_file):
        print(f"Error: Could not find the file at {input_file}")
        return

    # Load the 'Recipes' sheet
    df = pd.read_excel(input_file, sheet_name="Recipes", header=None)
    extracted_data = []

    # Iterate through columns (starting index 1 skips the labels)
    for col_idx in range(1, df.shape[1]):
        amine_name = str(df.iloc[1, col_idx]).strip()

        # Skip empty or control columns
        if pd.isna(df.iloc[1, col_idx]) or "MeCN" in amine_name or amine_name == "nan":
            continue

        # Amine level metadata
        amine_meta = {
            "Amine_Name": amine_name,
            "Amine_MW_g_mol": df.iloc[3, col_idx],
            "Actual_Mass_Amine_g": df.iloc[7, col_idx],
            "Actual_Amine_Conc_mM": df.iloc[8, col_idx],
            "Actual_Mass_HMDSO_mg": df.iloc[11, col_idx],
            "Actual_Conc_HMDSO_mM": df.iloc[12, col_idx],
        }

        # Each aldehyde block is 15 rows high, 20 blocks total, starting at row 13
        for i in range(20):
            row = 13 + (i * 15)
            aldehyde_name = df.iloc[row, col_idx]

            if pd.isna(aldehyde_name):
                continue

            reaction = amine_meta.copy()
            reaction.update(
                {
                    "Aldehyde_Name": aldehyde_name,
                    "Aldehyde_Actual_Mass_mg": df.iloc[row + 6, col_idx],
                    "Aldehyde_Vol_Required_uL": df.iloc[row + 7, col_idx],
                    "Actual_Aldehyde_Conc_mM": df.iloc[row + 8, col_idx],
                    "Vol_Amine_Sol_uL": df.iloc[row + 9, col_idx],
                    "Amount_Amine_mmol": df.iloc[row + 10, col_idx],
                    "Vol_Aldehyde_Sol_uL": df.iloc[row + 11, col_idx],
                    "Total_Volume_uL": df.iloc[row + 12, col_idx],
                    "Amount_Aldehyde_mmol": df.iloc[row + 13, col_idx],
                    "Amount_HMDSO_mmol": df.iloc[row + 14, col_idx],
                }
            )
            extracted_data.append(reaction)

    # Save to CSV
    final_df = pd.DataFrame(extracted_data)
    final_df.to_csv(output_file, index=False)
    print(f"Done! Extracted {len(final_df)} reactions to {output_file}")


# Execution
if __name__ == "__main__":
    input_path = "data/raw/reaction_planner_tracker.xlsx"
    output_path = "data/output/extracted_reaction_data.csv"
    extract_chemistry_data(input_path, output_path)
