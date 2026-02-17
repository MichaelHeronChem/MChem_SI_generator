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

def generate_unified_si(csv_input, output_docx):
    if not os.path.exists(csv_input):
        print(f"Error: Could not find {csv_input}")
        return

    df = pd.read_csv(csv_input)
    doc = Document()
    
    blocks = df.groupby('Experiment_Block')

    for block_id, block_data in blocks:
        # --- SECTION 1: BLOCK HEADER ---
        unique_amines_list = block_data['Amine_Name'].unique()
        amines_str = ", ".join(map(str, unique_amines_list))
        
        h1 = doc.add_paragraph()
        run1 = h1.add_run(f"Imine formation reactions block {block_id}: Amines {amines_str} with 20 aldehydes")
        run1.bold = True
        run1.font.size = Pt(14)
        h1.space_after = Pt(0)

        # --- SECTION 2: STOCK SOLUTIONS ---
        h2 = doc.add_paragraph()
        run2 = h2.add_run("Stock Solution Recipes")
        run2.bold = True
        run2.underline = True
        run2.font.size = Pt(12)
        h2.space_after = Pt(12)

        # Amine Stocks
        amines = block_data.drop_duplicates(subset=['Amine_Name'])
        for _, row in amines.iterrows():
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

        # Aldehyde Stocks
        aldehydes = block_data.drop_duplicates(subset=['Aldehyde_Name'])
        for _, row in aldehydes.iterrows():
            ald_mmol = row['Aldehyde_Actual_Mass_mg'] / row['Aldehyde_MW_g_mol']
            ald_conc = format_sig_figs(row['Actual_Aldehyde_Conc_mM'], 4)
            
            p = doc.add_paragraph()
            p.add_run(f"{row['Aldehyde_Name']} stock solution: ").bold = True
            p.add_run(f"{row['Aldehyde_Name']} ({row['Aldehyde_Actual_Mass_mg']:.2f} mg, {ald_mmol:.3f} mmol) ")
            p.add_run(f"was added to a sample vial, {row['Aldehyde_Vol_Required_uL']:.1f} µL of 15% MeCN-d3 in MeCN")
            p.add_run(f" was dispensed by the Hamilton liquid handler using 1000 µL disposable tips to create a ")
            p.add_run(f"{ald_conc} mM {row['Aldehyde_Name']} stock solution.")

        # --- SECTION 3: INDIVIDUAL REACTION PROCEDURES ---
        doc.add_paragraph().add_run("\nIndividual Reaction Procedures").bold = True
        
        amine_groups = block_data.groupby('Amine_Name', sort=False)
        for amine_name, amine_group in amine_groups:
            doc.add_heading(f"Reactions with {amine_name}", level=2)

            for _, row in amine_group.iterrows():
                amine_conc = format_sig_figs(row['Actual_Amine_Conc_mM'], 3)
                ald_conc = format_sig_figs(row['Actual_Aldehyde_Conc_mM'], 4)
                v_amine = f"{row['Vol_Amine_Sol_uL']:.1f}"
                v_ald = f"{row['Vol_Aldehyde_Sol_uL']:.1f}"

                p = doc.add_paragraph()
                p.add_run("To a 1.5 mL sample vial, ")
                p.add_run(f"{v_amine} µL ").bold = True
                p.add_run(f"of {amine_conc} mM ")
                p.add_run(f"{amine_name} ").bold = True
                p.add_run("solution and ")
                p.add_run(f"{v_ald} µL ").bold = True
                p.add_run(f"of {ald_conc} mM ")
                p.add_run(f"{row['Aldehyde_Name']} ").bold = True
                p.add_run("solution was added using the Hamilton liquid handler with 300 µL disposable tips. ")
                p.add_run(" A teflon coated stirrer bar was added to the vial, the vial was sealed, and the mixture left to stir overnight. ")
                p.add_run("The reaction mixture was then transferred to an NMR tube and a 1H NMR was taken with acetonitrile solvent suppression.")

        doc.add_page_break()

    doc.save(output_docx)
    print(f"Unified SI generated: {output_docx}")

if __name__ == "__main__":
    generate_unified_si("data/output/extracted_reaction_data.csv", "data/output/Complete_Supporting_Information.docx")