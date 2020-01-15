import docx  # To read docx and extract data
from docx.oxml import OxmlElement  # For defining and targeting xml elements to change
from docx.oxml.ns import qn  # For defining and targeting xml elements to change
from docx.enum.section import WD_SECTION  # To get word sections
from docx.text.paragraph import Paragraph
from docxcompose.composer import Composer  # Append files together, preserving everything except sections

# Helper functions
def get_para_data(dest, src_p):
    """
    Write the run to the new file and then set its font, bold, alignment, color etc. data.
    
    More text attributes: https://python-docx.readthedocs.io/en/latest/api/text.html
    """
    dest_para = dest.add_paragraph(style=src_p.style.name)

    for run in src_p.runs:
        
        dest_run = dest_para.add_run(run.text)
        
        # Apply text styles
        dest_run.bold = run.bold
        dest_run.italic = run.italic
        dest_run.underline = run.underline
        dest_run.style.name = run.style.name
        dest_run.font.color.rgb = run.font.color.rgb
        dest_run.font.name = run.font.name
        dest_run.font.subscript = run.font.subscript
        dest_run.font.superscript = run.font.superscript
        dest_run.font.size = run.font.size
        
        # Add run for footnote
        if run.footnote:
            dest_para.add_footnote(run.footnote)
        
    # Align paragraph
    dest_para.alignment = src_p.alignment
        

def new_section_cols(dest, num_cols):
    new_section = dest.add_section(WD_SECTION.CONTINUOUS)
    sectPr = new_section._sectPr
    cols = OxmlElement('w:cols')
    cols.set(qn('w:num'), str(num_cols))
    sectPr.append(cols)