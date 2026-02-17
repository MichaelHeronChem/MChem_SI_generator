import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
import os
import math

def format_sig_figs(val, n):
    if pd.isna(val) or val == 0:
        return "0"
    decimals = n - 1 - int(math.floor(math.log10(abs(val))))
    if decimals <= 0:
        return format(round(val, decimals), '.0f')
    return format(val, f'.{decimals}f')

def generate_final_si(csv_input, nmr_image_folder, output_docx):
    if not os.path.exists(csv_input):
        print(f"Error: Could not find {csv_input}")
        return

    df = pd.read_csv(csv_input)
    doc = Document()
    
    # Global Style: Minimize spacing
    style = doc.styles['Normal']
    style.paragraph_format.space_after = Pt(2) # Minimal gap between text and image

    blocks = df.groupby('Experiment_Block')

    for block_id, block_data in blocks:
        # --- 1. BLOCK HEADER (14pt) ---
        unique_amines = block_data['Amine_Name'].unique()
        amines_str = ", ".join(map(str, unique_amines))
        h1 = doc.add_paragraph()
        run1 = h1.add_run(f"Imine formation reactions block {block_id}: Amines {amines_str} with 20 aldehydes")
        run1.bold = True
        run1.font.size = Pt(14)

        # --- 2. RESTORED: STOCK SOLUTION RECIPES (12pt, Bold, Underline) ---
        h2 = doc.add_paragraph()
        run2 = h2.add_run("Stock Solution Recipes")
        run2.bold = True
        run2.underline = True
        run2.font.size = Pt(12)

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
            p.add_run(f"{row['Amine_Name']} ({amine_mass:.2f} mg, {amine_mmol:.3f} mmol) and hexamethyldisiloxane ({row['Actual_Mass_HMDSO_mg']:.2f} mg, {hmdso_mmol:.3f} mmol) was added to a 10 mL volumetric flask and filled up to the mark with 15% MeCN-d3 in MeCN dried over molecular sieves to give 10 mL of {amine_conc} mM {row['Amine_Name']} and {hmdso_conc} mM hexamethyldisiloxane stock solution.")

        # Aldehyde Stocks
        aldehydes = block_data.drop_duplicates(subset=['Aldehyde_Name'])
        for _, row in aldehydes.iterrows():
            ald_mmol = row['Aldehyde_Actual_Mass_mg'] / row['Aldehyde_MW_g_mol']
            ald_conc = format_sig_figs(row['Actual_Aldehyde_Conc_mM'], 4)
            
            p = doc.add_paragraph()
            p.add_run(f"{row['Aldehyde_Name']} stock solution: ").bold = True
            p.add_run(f"{row['Aldehyde_Name']} ({row['Aldehyde_Actual_Mass_mg']:.2f} mg, {ald_mmol:.3f} mmol) was added to a sample vial, {row['Aldehyde_Vol_Required_uL']:.1f} µL was dispensed into the sample vial by the Hamilton liquid handler using 1000 µL disposable tips to create a {ald_conc} mM {row['Aldehyde_Name']} stock solution.")

        # --- 3. REACTION PROCEDURES & FIGURES ---
        doc.add_paragraph().add_run("\nIndividual Reaction Procedures and 1H NMR Spectra").bold = True
        
        amine_groups = block_data.groupby('Amine_Name', sort=False)
        for amine_name, amine_group in amine_groups:
            doc.add_heading(f"Reactions with {amine_name}", level=2)

            for _, row in amine_group.iterrows():
                # Text First
                amine_conc = format_sig_figs(row['Actual_Amine_Conc_mM'], 3)
                ald_conc = format_sig_figs(row['Actual_Aldehyde_Conc_mM'], 4)
                
                p = doc.add_paragraph()
                p.add_run("To a sample vial was added, ").bold = False
                p.add_run(f"{row['Vol_Amine_Sol_uL']:.1f} µL ").bold = True
                p.add_run(f"of {amine_conc} mM {amine_name} solution and ")
                p.add_run(f"{row['Vol_Aldehyde_Sol_uL']:.1f} µL ").bold = True
                p.add_run(f"of {ald_conc} mM {row['Aldehyde_Name']} solution using the Hamilton liquid handler with 300 µL disposable tips. Teflon coated stirrer bars was added to the vial and the vial sealed and the mixture left to stir overnight. The reaction mixture was then transferred to an NMR tube and 1H NMR was taken with acetonitrile solvent suppression.")

                # Image Second
                img_filename = f"{amine_name.replace(' ', '_')}_{row['Aldehyde_Name'].replace(' ', '_')}.png"
                img_path = os.path.join(nmr_image_folder, img_filename)

                # Caption/Title for Figure
                cap = doc.add_paragraph()
                cap.paragraph_format.space_before = Pt(6)
                cap.add_run(f"Figure: 1H NMR of {amine_name} and {row['Aldehyde_Name']}").bold = True

                if os.path.exists(img_path):
                    doc.add_picture(img_path, width=Inches(5.5)) # Reduced width slightly to help with whitespace/page fit
                else:
                    doc.add_paragraph(f"[MISSING NMR: {img_filename}]").style = 'Intense Quote'

                doc.add_page_break()

    doc.save(output_docx)
    print(f"Final SI generated successfully: {output_docx}")

if __name__ == "__main__":
    NMR_IMG_DIR = "data/raw/nmr_images"
    generate_final_si("data/output/extracted_reaction_data.csv", NMR_IMG_DIR, "data/output/Final_SI_Report.docx")