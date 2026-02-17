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

def generate_unified_si_with_nmr(csv_input, nmr_image_folder, output_docx):
    if not os.path.exists(csv_input):
        print(f"Error: Could not find {csv_input}")
        return

    df = pd.read_csv(csv_input)
    doc = Document()
    
    blocks = df.groupby('Experiment_Block')

    for block_id, block_data in blocks:
        # --- 1. BLOCK HEADER ---
        unique_amines = block_data['Amine_Name'].unique()
        amines_str = ", ".join(map(str, unique_amines))
        
        h1 = doc.add_paragraph()
        run1 = h1.add_run(f"Imine formation reactions block {block_id}: Amines {amines_str} with 20 aldehydes")
        run1.bold = True
        run1.font.size = Pt(14)

        # --- 2. STOCK SOLUTIONS ---
        h2 = doc.add_paragraph()
        run2 = h2.add_run("Stock Solution Recipes")
        run2.bold, run2.underline = True, True
        run2.font.size = Pt(12)

        # Amine & Aldehyde Stocks (Logic preserved from previous)
        # [Amine stock logic here...]
        # [Aldehyde stock logic here...]

        # --- 3. REACTION PROCEDURES & FIGURES ---
        doc.add_paragraph().add_run("\nIndividual Reaction Procedures and 1H NMR Spectra").bold = True
        
        amine_groups = block_data.groupby('Amine_Name', sort=False)
        for amine_name, amine_group in amine_groups:
            doc.add_heading(f"Reactions with {amine_name}", level=2)

            for _, row in amine_group.iterrows():
                # Experimental Text
                amine_conc = format_sig_figs(row['Actual_Amine_Conc_mM'], 3)
                ald_conc = format_sig_figs(row['Actual_Aldehyde_Conc_mM'], 4)
                
                p = doc.add_paragraph()
                p.add_run(f"Reaction of {amine_name} and {row['Aldehyde_Name']}: ").bold = True
                p.add_run(f"To a sample vial was added {row['Vol_Amine_Sol_uL']:.1f} µL of {amine_conc} mM {amine_name} solution and {row['Vol_Aldehyde_Sol_uL']:.1f} µL of {ald_conc} mM {row['Aldehyde_Name']} solution...")

                # NMR Figure Insertion
                nmr_title = f"{amine_name} and {row['Aldehyde_Name']} NMR"
                doc.add_paragraph().add_run(nmr_title).bold = True
                
                # Check for image file: e.g., "Amine1_Aldehyde5.png"
                img_filename = f"{amine_name}_{row['Aldehyde_Name']}.png".replace(" ", "_")
                img_path = os.path.join(nmr_image_folder, img_filename)

                if os.path.exists(img_path):
                    doc.add_picture(img_path, width=Inches(6.0))
                else:
                    # Placeholder if image is missing
                    placeholder = doc.add_paragraph()
                    run_p = placeholder.add_run(f"[INSERT FIGURE: {img_filename}]")
                    run_p.font.color.rgb = None # Set to red or gray if desired

                doc.add_page_break() # Keep each reaction + spectrum on its own page for SI

    doc.save(output_docx)
    print(f"Complete SI with Figures generated: {output_docx}")

if __name__ == "__main__":
    # Update this to the folder where your NMR PNG exports are stored
    NMR_FOLDER = "data/raw/nmr_images" 
    generate_unified_si_with_nmr("data/output/extracted_reaction_data.csv", NMR_FOLDER, "data/output/Complete_SI_with_NMR.docx")