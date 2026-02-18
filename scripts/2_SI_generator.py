import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from fpdf import FPDF
import os
import math

def format_sig_figs(val, n):
    if pd.isna(val) or val == 0: return "0"
    decimals = n - 1 - int(math.floor(math.log10(abs(val))))
    return format(val, f'.{decimals}f') if decimals > 0 else format(round(val, decimals), '.0f')

class SI_PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 10)
        self.cell(0, 10, 'Supporting Information', 0, 1, 'R')

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def generate_si_files(csv_input, nmr_image_folder, output_base):
    df = pd.read_csv(csv_input)
    doc = Document()
    pdf = SI_PDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # --- 1. GLOBAL EXPERIMENTAL METHODS (At the very top) ---
    doc.add_heading('General Experimental Procedures', level=1)
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, 'General Experimental Procedures', 0, 1)
    pdf.ln(2)

    methods = [
        ("Amine Stock Solutions", "Amine stock solutions were made as follows: Amine and HMDSO was added to 10 mL volumetric flask and filled up to the mark with 15% MeCN-d3 in MeCN dried over molecular sieves to give 10 mL of amine stock solution."),
        ("Aldehyde Stock Solutions", "Aldehyde stock solutions were made as follows: Aldehyde was added to a sample vial. Enough MeCN was dispensed into the sample vial by the Hamilton liquid handler using 1000 uL disposable tips to create a 100.00 mM aldehyde solution."),
        ("Reaction Method", "Reactions were made up using Hamilton liquid handler with 300 uL disposable tips. Teflon coated stirrer bars were added to the vials, which were then sealed and left to stir overnight before NMR analysis.")
    ]

    for title, text in methods:
        # Word
        p = doc.add_paragraph()
        p.add_run(f"{title}: ").bold = True
        p.add_run(text)
        # PDF
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(0, 8, title, 0, 1)
        pdf.set_font('Arial', '', 10)
        pdf.multi_cell(0, 5, text)
        pdf.ln(4)

    # --- 2. BLOCK-SPECIFIC DATA ---
    blocks = df.groupby('Experiment_Block')

    for block_id, block_data in blocks:
        doc.add_page_break()
        pdf.add_page()
        
        block_title = f"Imine formation reactions block {block_id}"
        doc.add_heading(block_title, level=1)
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, block_title, 0, 1)
        pdf.ln(5)

        # --- AMINE TABLE ---
        doc.add_heading(f"Amine Stocks - Block {block_id}", level=2)
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(0, 8, "Amine Stock Details", 0, 1)
        
        amines = block_data.drop_duplicates(subset=['Amine_Name'])
        headers = ['Amine Name', 'Mass (mg)', 'Amount (mmol)', 'Conc (mM)', 'HMDSO (mmol)']
        # PDF Column Widths: Name is much wider to prevent spill
        a_widths = [65, 30, 30, 30, 35] 
        
        t1 = doc.add_table(rows=1, cols=5)
        t1.style = 'Table Grid'
        for i, h in enumerate(headers): t1.rows[0].cells[i].text = h

        for _, row in amines.iterrows():
            m_mg = row['Actual_Mass_Amine_g'] * 1000
            vals = [str(row['Amine_Name']), f"{m_mg:.2f}", f"{m_mg/row['Amine_MW_g_mol']:.3f}", 
                    format_sig_figs(row['Actual_Amine_Conc_mM'], 3), f"{row['Actual_Mass_HMDSO_mg']/162.38:.3f}"]
            
            cells = t1.add_row().cells
            pdf.set_font('Arial', '', 9)
            for i, v in enumerate(vals):
                cells[i].text = v
                pdf.cell(a_widths[i], 7, v, 1)
            pdf.ln()

        # --- ALDEHYDE TABLE ---
        doc.add_heading(f"Aldehyde Stocks - Block {block_id}", level=2)
        pdf.ln(10)
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(0, 8, "Aldehyde Stock Details", 0, 1)

        aldehydes = block_data.drop_duplicates(subset=['Aldehyde_Name'])
        headers = ['Aldehyde Name', 'Mass (mg)', 'Amount (mmol)', 'Vol (uL)', 'Conc (mM)']
        
        t2 = doc.add_table(rows=1, cols=5)
        t2.style = 'Table Grid'
        for i, h in enumerate(headers): t2.rows[0].cells[i].text = h

        for _, row in aldehydes.iterrows():
            m_mg = row['Aldehyde_Actual_Mass_mg']
            vals = [str(row['Aldehyde_Name']), f"{m_mg:.2f}", f"{m_mg/row['Aldehyde_MW_g_mol']:.3f}", 
                    f"{row['Aldehyde_Vol_Required_uL']:.1f}", format_sig_figs(row['Actual_Aldehyde_Conc_mM'], 4)]
            
            cells = t2.add_row().cells
            pdf.set_font('Arial', '', 9)
            for i, v in enumerate(vals):
                cells[i].text = v
                pdf.cell(a_widths[i], 7, v, 1)
            pdf.ln()

        # --- REACTION TABLE ---
        doc.add_heading(f"Reaction Volumes - Block {block_id}", level=2)
        pdf.ln(10)
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(0, 8, "Reaction Volumes", 0, 1)

        headers = ['Amine', 'Aldehyde', 'Vol Amine (uL)', 'Vol Ald (uL)']
        r_widths = [50, 50, 45, 45]
        
        t3 = doc.add_table(rows=1, cols=4)
        t3.style = 'Table Grid'
        for i, h in enumerate(headers): t3.rows[0].cells[i].text = h

        for _, row in block_data.iterrows():
            vals = [str(row['Amine_Name']), str(row['Aldehyde_Name']), f"{row['Vol_Amine_Sol_uL']:.1f}", f"{row['Vol_Aldehyde_Sol_uL']:.1f}"]
            cells = t3.add_row().cells
            pdf.set_font('Arial', '', 8) # Smaller font for the big reaction table
            for i, v in enumerate(vals):
                cells[i].text = v
                pdf.cell(r_widths[i], 6, v, 1)
            pdf.ln()

    # --- 3. NMR APPENDIX ---
    doc.add_page_break()
    pdf.add_page()
    doc.add_heading('Appendix: 1H NMR Spectra', level=1)
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, "Appendix: 1H NMR Spectra", 0, 1)

    for _, row in df.iterrows():
        title = f"1H NMR: {row['Amine_Name']} + {row['Aldehyde_Name']}"
        img_name = f"{str(row['Amine_Name']).replace(' ', '_')}_{str(row['Aldehyde_Name']).replace(' ', '_')}.png"
        img_path = os.path.join(nmr_image_folder, img_name)
        
        doc.add_paragraph(title).bold = True
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(0, 8, title, 0, 1)

        if os.path.exists(img_path):
            doc.add_picture(img_path, width=Inches(6.0))
            if pdf.get_y() > 180: pdf.add_page()
            pdf.image(img_path, w=180)
            pdf.ln(5)

    doc.save(f"{output_base}.docx")
    pdf.output(f"{output_base}.pdf")
    print(f"Files successfully generated: {output_base}.docx and .pdf")

if __name__ == "__main__":
    generate_si_files("data/output/extracted_reaction_data.csv", "data/raw/nmr_images", "data/output/Final_SI_Report")