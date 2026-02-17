import pandas as pd
from docx import Document
from docx.shared import Pt
import os
import math

def format_sig_figs(val, n):
    """Formats a value to exactly n significant figures, preserving trailing zeros."""
    if pd.isna(val) or val == 0:
        return "0"
    decimals = n - 1 - int(math.floor(math.log10(abs(val))))
    if decimals <= 0:
        return format(round(val, decimals), '.0f')
    return format(val, f'.{decimals}f')

def generate_si_stock_solutions(csv_input, output_docx):
    if not os.path.exists(csv_input):
        print(f"Error: Could not find {csv_input}")
        return

    df = pd.read_csv(csv_input)
    doc = Document()
    
    blocks = df.groupby('Experiment_Block')

    for block_id, block_data in blocks:
        # --- HEADER SECTION ---
        unique_amines_list = block_data['Amine_Name'].unique()
        amines_str = ", ".join(map(str, unique_amines_list))
        
        # 1. Main Header: Bold, Large (14pt)
        h1 = doc.add_paragraph()
        run1 = h1.add_run(f"Imine formation reactions block {block_id}: Amines {amines_str} with 20 aldehydes")
        run1.bold = True
        run1.font.size = Pt(14)
        h1.space_after = Pt(0) 

        # 2. Sub-Header: Stock Solution Recipes (12pt, Bold, Underlined)
        h2 = doc.add_paragraph()
        run2 = h2.add_run("Stock Solution Recipes")
        run2.bold = True
        run2.underline = True
        run2.font.size = Pt(12)
        h2.space_after = Pt(12) 

        # --- AMINE STOCK SOLUTIONS ---
        amines = block_data.drop_duplicates(subset=['Amine_Name'])
        for _, row in amines.iterrows():
            # Standardizing mass and mmol calculations
            amine_mass = row['Actual_Mass_Amine_g'] 
            amine_mmol = amine_mass / row['Amine_MW_g_mol']
            hmdso_mmol = row['Actual_Mass_HMDSO_mg'] / 162.38
            
            amine_conc = format_sig_figs(row['Actual_Amine_Conc_mM'], 3)
            hmdso_conc = format_sig_figs(row['Actual_Conc_HMDSO_mM'], 3)
            
            p = doc.add_paragraph()
            p.add_run(f"{row['Amine_Name']} stock solution: ").bold = True
            p.add_run(f"{row['Amine_Name']} ({amine_mass:.2f} mg, {amine_mmol:.3f} mmol) and ")
            p.add_run(f"hexamethyldisiloxane ({row['Actual_Mass_HMDSO_mg']:.2f} mg, {hmdso_mmol:.3f} mmol) ")
            p.add_run(f"was added to a 10 mL volumetric flask and filled up to the mark with 15% MeCN-d3 in MeCN ")
            p.add_run(f"dried over molecular sieves to give 10 mL of {amine_conc} mM ")
            p.add_run(f"{row['Amine_Name']} and {hmdso_conc} mM hexamethyldisiloxane stock solution.")

        # --- ALDEHYDE STOCK SOLUTIONS ---
        aldehydes = block_data.drop_duplicates(subset=['Aldehyde_Name'])
        for _, row in aldehydes.iterrows():
            ald_mmol = row['Aldehyde_Actual_Mass_mg'] / row['Aldehyde_MW_g_mol']
            ald_conc = format_sig_figs(row['Actual_Aldehyde_Conc_mM'], 4)
            
            p = doc.add_paragraph()
            p.add_run(f"{row['Aldehyde_Name']} stock solution: ").bold = True
            p.add_run(f"{row['Aldehyde_Name']} ({row['Aldehyde_Actual_Mass_mg']:.2f} mg, {ald_mmol:.3f} mmol) ")
            p.add_run(f"was added to a sample vial, {row['Aldehyde_Vol_Required_uL']:.1f} uL was dispensed into ")
            p.add_run(f"the sample vial by the Hamilton liquid handler using 1000 uL disposable tips to create a ")
            p.add_run(f"{ald_conc} mM {row['Aldehyde_Name']} stock solution.")

        doc.add_page_break()

    doc.save(output_docx)
    print(f"SI generated with bold/underlined sub-headers: {output_docx}")

if __name__ == "__main__":
    generate_si_stock_solutions("data/output/extracted_reaction_data.csv", "data/output/SI_Stock_Solutions.docx")