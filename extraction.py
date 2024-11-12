import fitz
from pathlib import Path

# Define path to document
path = Path(Path.home().joinpath("Desktop", "BMC3.2", "BMC 4.1", "SoftWareEngineer", 
                                 "Software Engineering A Practitioners Approach by Roger Pressman, Bruce Maxim (z-lib.org).pdf"))
out_path = Path(Path.home().joinpath("Desktop", "BMC3.2", "BMC 4.1", "SoftWareEngineer", "chapter_7.pdf"))
out_path2 = Path(Path.home().joinpath("Desktop", "BMC3.2", "BMC 4.1", "SoftWareEngineer", "chapter_8.pdf"))
out_path3 = Path(Path.home().joinpath("Desktop", "BMC3.2", "BMC 4.1", "SoftWareEngineer", "chapter_9.pdf"))

def extract_text(path, start_page, end_page, out_path):
    # Open the original PDF
    doc = fitz.open(path)
    # Create a new PDF for the output
    output_pdf = fitz.open()
    
    # Loop through the specified range of pages
    for num_page in range(start_page - 1, end_page):
        page = doc.load_page(num_page)
        output_pdf.insert_pdf(doc, from_page=num_page, to_page=num_page)
    
    # Save the output PDF
    output_pdf.save(out_path)
    output_pdf.close()
    doc.close()
    print(f"\nPages {start_page} to {end_page} have been saved to {out_path}")

# Run the extraction function
extract_text(path, 133, 160, out_path)
extract_text(path, 161, 195, out_path2)
extract_text(path, 196, 213, out_path3)






