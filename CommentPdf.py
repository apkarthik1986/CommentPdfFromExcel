import fitz  # PyMuPDF
import pandas as pd
import os
from tkinter import Tk, messagebox, simpledialog
from tkinter.filedialog import askopenfilename, askdirectory

def update_pdf_with_comments(pdf_path, excel_path, output_pdf_path, subject="Comment", distance=10):
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

                    # Position comment box to the right of the tag using dynamic distance
                    rect = fitz.Rect(inst.x1 + 0, inst.y0 , inst.x1 + distance + width, inst.y0 + height )

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
                    annot.set_info({"subject": subject})
                    annot.update()

    doc.save(output_pdf_path)
    doc.close()

def process_folder(folder_path, excel_path, subject="Comment", distance=10):
    # Get all PDF files in the folder
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        messagebox.showwarning("Warning", "No PDF files found in the selected folder.")
        return
    
    for pdf_file in pdf_files:
        try:
            pdf_path = os.path.join(folder_path, pdf_file)
            # Save in the same folder with "updated_" prefix
            output_pdf_path = os.path.join(folder_path, f"updated_{pdf_file}")
            
            update_pdf_with_comments(pdf_path, excel_path, output_pdf_path, subject, distance)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process {pdf_file}:\n{e}")
            continue
    
    messagebox.showinfo("Success", f"Processed {len(pdf_files)} PDF files.\nSaved in the same folder as input PDFs.")

def get_annotation_distance():
    """Show a dialog to get annotation distance from user."""
    distance_str = simpledialog.askstring(
        "Annotation Distance",
        "Enter the distance between tag and annotation (in points):\n(Default: 10)"
    )
    # If user cancels or provides empty string, use default
    if not distance_str or distance_str.strip() == "":
        return 10
    try:
        distance = int(distance_str)
        if distance < 0:
            messagebox.showwarning("Warning", "Distance cannot be negative. Using default value of 10.")
            return 10
        return distance
    except ValueError:
        messagebox.showwarning("Warning", "Invalid input. Using default value of 10.")
        return 10

def get_comment_subject():
    """Show a dialog to get comment subject from user."""
    subject = simpledialog.askstring(
        "Comment Subject",
        "Enter the comment subject:\n(Leave empty for default 'Comment')",
        initialvalue="Comment"
    )
    # If user cancels or provides empty string, use default
    if not subject or subject.strip() == "":
        subject = "Comment"
    return subject

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

        # Ask for comment subject
        subject = get_comment_subject()

        # Ask for annotation distance
        distance = get_annotation_distance()

        process_folder(folder_path, excel_path, subject, distance)
        
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

if __name__ == "__main__":
    main()
