import fitz  # PyMuPDF
import pandas as pd
import os
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename, askdirectory, asksaveasfilename

def update_pdf_with_comments(pdf_path, excel_path, output_pdf_path):
    # Read the Excel file
    df = pd.read_excel(excel_path)
    df['tag'] = df['tag'].astype(str)
    df['comment'] = df['comment'].astype(str)

    doc = fitz.open(pdf_path)

    for page_num in range(len(doc)):
        page = doc[page_num]
        content = page.get_text("text")

        for index, row in df.iterrows():
            tag = row['tag']
            comment = row['comment']

            if tag in content:
                text_instances = page.search_for(tag)

                for inst in text_instances:
                    # Estimate width based on character count
                    width = len(comment) * 5.5 + 20
                    height = 20

                    # Position comment box to the right of the tag #inst.x1 + 0 for Elevation
                    rect = fitz.Rect(inst.x1 + 0, inst.y0 , inst.x1 + 10 + width, inst.y0 + height )

                    # Add horizontal free text annotation
                    annot = page.add_freetext_annot(
                        rect, comment,
                        fontsize=12,
                        text_color=(0, 0, 0),
                        fill_color=(1, 1, 0),  # Optional: yellow background
                        rotate=0,
                        align=fitz.TEXT_ALIGN_LEFT
                    )
                    annot.set_border(width=0.5, dashes=[2])
                    annot.set_info({"subject": "tag2"})
                    annot.update()

    doc.save(output_pdf_path)
    doc.close()

def process_folder(folder_path, excel_path, output_folder):
    # Create output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Get all PDF files in the folder
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        messagebox.showwarning("Warning", "No PDF files found in the selected folder.")
        return
    
    for pdf_file in pdf_files:
        try:
            pdf_path = os.path.join(folder_path, pdf_file)
            output_pdf_path = os.path.join(output_folder, f"updated_{pdf_file}")
            
            update_pdf_with_comments(pdf_path, excel_path, output_pdf_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process {pdf_file}:\n{e}")
            continue
    
    messagebox.showinfo("Success", f"Processed {len(pdf_files)} PDF files.\nSaved in: {output_folder}")

def main():
    root = Tk()
    root.withdraw()

    try:
        # Ask for Excel file
        excel_path = askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
        if not excel_path:
            return

        # Ask for input folder containing PDFs
        folder_path = askdirectory(title="Select Folder with PDF Files")
        if not folder_path:
            return

        # Ask for output folder
        output_folder = askdirectory(title="Select Output Folder")
        if not output_folder:
            return

        process_folder(folder_path, excel_path, output_folder)
        
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

if __name__ == "__main__":
    main()
