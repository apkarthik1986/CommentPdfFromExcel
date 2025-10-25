import fitz  # PyMuPDF
import pandas as pd
import os
from tkinter import (
    Tk,
    StringVar,
    IntVar,
    Label,
    Entry,
    Button,
    Radiobutton,
    Text,
    END,
    W,
    E,
    N,
    S,
    DISABLED,
    NORMAL,
    filedialog,
    messagebox,
    Toplevel,
)
from tkinter.ttk import Frame

# Pillow is optional but required for preview mode
try:
    from PIL import Image, ImageDraw, ImageTk, ImageFont
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False


def update_pdf_with_comments(
    pdf_path, df, output_pdf_path, subject="Comment", distance=10, log_func=None
):
    """Update a single PDF using tags/comments from dataframe df.

    - df must contain 'tag' and 'comment' columns (strings).
    - distance is in points: how far from the tag the annotation should be placed.
    - Uses a fallback to place the comment to the left of the tag if there's no room to the right.
    """
    if log_func:
        log_func(f"Processing: {os.path.basename(pdf_path)}")

    doc = fitz.open(pdf_path)

    for page_num in range(len(doc)):
        page = doc[page_num]
        content = page.get_text("text")

        for index, row in df.iterrows():
            tag = str(row["tag"])
            comment = str(row["comment"])

            if not tag or tag.strip() == "":
                continue

            if tag in content:
                # find all text instances on this page
                text_instances = page.search_for(tag)

                for inst in text_instances:
                    # Determine comment box size
                    # Basic heuristic: 6 points per character + padding
                    width = max(60, len(comment) * 6 + 20)
                    # Height: a single line at fontsize ~12 -> ~16 points; allow two lines by increasing if long
                    height = 18

                    page_rect = page.rect

                    # Preferred position: to the right of the tag, offset by distance
                    pref_x0 = inst.x1 + distance
                    pref_x1 = pref_x0 + width

                    # If preferred position fits on the page, use it; otherwise place to left of tag
                    if pref_x1 <= page_rect.x1 - 5:
                        x0 = pref_x0
                        x1 = pref_x1
                    else:
                        # Place left of tag
                        x1 = inst.x0 - distance
                        x0 = x1 - width
                        # If still negative, clamp to page with minimal margins
                        if x0 < page_rect.x0 + 5:
                            # adjust width to fit
                            x0 = page_rect.x0 + 5
                            x1 = min(page_rect.x1 - 5, x0 + width)

                    # Vertical position: align top with inst.y0
                    # If the box would go off bottom edge, adjust up
                    y0 = inst.y0
                    y1 = y0 + height
                    if y1 > page_rect.y1 - 5:
                        y1 = page_rect.y1 - 5
                        y0 = y1 - height
                        if y0 < page_rect.y0 + 5:
                            y0 = page_rect.y0 + 5

                    rect = fitz.Rect(x0, y0, x1, y1)

                    # Add horizontal free text annotation
                    try:
                        annot = page.add_freetext_annot(
                            rect,
                            comment,
                            fontsize=12,
                            text_color=(0, 0, 0),
                            fill_color=(1, 1, 0),  # yellow background
                            rotate=0,
                            align=fitz.TEXT_ALIGN_LEFT,
                        )
                        annot.set_border(width=0.5, dashes=[2])
                        annot.set_info({"subject": subject})
                        annot.update()
                    except Exception as e:
                        if log_func:
                            log_func(f"  Warning: could not add annotation at {rect}: {e}")

    doc.save(output_pdf_path)
    doc.close()
    if log_func:
        log_func(f"Saved: {os.path.basename(output_pdf_path)}")


def process_files(
    pdf_paths,
    excel_path,
    output_folder,
    subject="Comment",
    distance=10,
    log_func=None,
):
    # Read Excel
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        raise RuntimeError(f"Failed to read Excel file: {e}")

    # Validate columns
    if "tag" not in df.columns or "comment" not in df.columns:
        raise RuntimeError("Excel must contain 'tag' and 'comment' columns.")

    # Ensure strings
    df["tag"] = df["tag"].astype(str)
    df["comment"] = df["comment"].astype(str)

    os.makedirs(output_folder, exist_ok=True)

    for pdf_path in pdf_paths:
        base = os.path.basename(pdf_path)
        output_pdf_path = os.path.join(output_folder, f"updated_{base}")
        try:
            update_pdf_with_comments(
                pdf_path, df, output_pdf_path, subject=subject, distance=distance, log_func=log_func
            )
        except Exception as e:
            if log_func:
                log_func(f"Error processing {base}: {e}")
            continue


# ---------- Preview utilities ----------
def render_page_preview_image(page, annotations, zoom=2.0):
    """Render a fitz page to a PIL image and draw annotation rectangles and text.

    - page: fitz.Page
    - annotations: list of dicts with keys 'rect' (fitz.Rect) and 'text'
    - zoom: scale factor (points -> pixels)
    """
    # Render page to pixmap (no alpha)
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    mode = "RGB"
    img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)

    # Convert to RGBA so we can draw semi-transparent highlights
    img = img.convert("RGBA")
    overlay = Image.new("RGBA", img.size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay)

    # Font fallback
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", size=int(12 * zoom))
    except Exception:
        font = ImageFont.load_default()

    for ann in annotations:
        rect = ann["rect"]
        text = ann["text"]

        # Convert points to pixels
        x0 = int(rect.x0 * zoom)
        y0 = int(rect.y0 * zoom)
        x1 = int(rect.x1 * zoom)
        y1 = int(rect.y1 * zoom)

        # Draw semi-transparent yellow rectangle
        draw.rectangle([x0, y0, x1, y1], fill=(255, 255, 0, 160), outline=(0, 0, 0, 200))

        # Draw text (clip if necessary)
        # Compute max width and do naive trimming with ellipsis if needed
        max_w = x1 - x0 - 4
        txt = text
        w, h = draw.textsize(txt, font=font)
        if w > max_w and max_w > 10:
            # trim characters until fit
            while txt and draw.textsize(txt + "...", font=font)[0] > max_w:
                txt = txt[:-1]
            txt = txt + "..."
        draw.text((x0 + 2, y0 + 1), txt, fill=(0, 0, 0, 255), font=font)

    # Composite overlay on the rendered page
    composite = Image.alpha_composite(img, overlay)
    return composite.convert("RGB")


def build_annotations_for_preview(page, df, distance):
    """Given a page and the df, compute annotation rectangles and texts (same logic as main flow)."""
    annotations = []
    content = page.get_text("text")

    for index, row in df.iterrows():
        tag = str(row["tag"])
        comment = str(row["comment"])
        if not tag or tag.strip() == "":
            continue

        if tag in content:
            text_instances = page.search_for(tag)
            for inst in text_instances:
                width = max(60, len(comment) * 6 + 20)
                height = 18
                page_rect = page.rect

                pref_x0 = inst.x1 + distance
                pref_x1 = pref_x0 + width

                if pref_x1 <= page_rect.x1 - 5:
                    x0 = pref_x0
                    x1 = pref_x1
                else:
                    x1 = inst.x0 - distance
                    x0 = x1 - width
                    if x0 < page_rect.x0 + 5:
                        x0 = page_rect.x0 + 5
                        x1 = min(page_rect.x1 - 5, x0 + width)

                y0 = inst.y0
                y1 = y0 + height
                if y1 > page_rect.y1 - 5:
                    y1 = page_rect.y1 - 5
                    y0 = y1 - height
                    if y0 < page_rect.y0 + 5:
                        y0 = page_rect.y0 + 5

                rect = fitz.Rect(x0, y0, x1, y1)
                annotations.append({"rect": rect, "text": comment})
    return annotations


def show_preview_window(parent, pdf_path, df, subject, distance):
    """Open a Toplevel that displays a preview of one PDF page with annotations.
    Allows navigation through pages (Prev/Next).
    """
    if not PIL_AVAILABLE:
        messagebox.showerror("Preview unavailable", "Pillow is required for preview mode. Install it with: pip install pillow")
        return

    doc = fitz.open(pdf_path)
    if len(doc) == 0:
        messagebox.showwarning("Preview", "PDF has no pages.")
        doc.close()
        return

    # Build a small UI window
    win = Toplevel(parent)
    win.title(f"Preview - {os.path.basename(pdf_path)}")
    win.geometry("900x700")

    img_label = Label(win)
    img_label.pack(expand=True, fill="both")

    info_label = Label(win, text="", anchor="w")
    info_label.pack(fill="x")

    # state
    state = {"page_index": 0, "photo": None}

    def render_current_page():
        page_i = state["page_index"]
        page = doc[page_i]
        annotations = build_annotations_for_preview(page, df, distance)
        if not annotations:
            # if no annotations found, still show the page
            annotations = []
        pil_img = render_page_preview_image(page, annotations, zoom=2.0)
        # resize to window while keeping aspect ratio
        w, h = pil_img.size
        win_w = max(400, win.winfo_width() - 20)
        win_h = max(300, win.winfo_height() - 80)
        ratio = min(win_w / w, win_h / h, 1.0)
        if ratio < 1.0:
            new_size = (int(w * ratio), int(h * ratio))
            pil_img = pil_img.resize(new_size, Image.LANCZOS)
        photo = ImageTk.PhotoImage(pil_img)
        state["photo"] = photo  # keep reference
        img_label.config(image=photo)
        info_label.config(text=f"Page {page_i + 1} / {len(doc)} - Found {len(annotations)} annotation(s)")

    def on_prev():
        if state["page_index"] > 0:
            state["page_index"] -= 1
            render_current_page()

    def on_next():
        if state["page_index"] < len(doc) - 1:
            state["page_index"] += 1
            render_current_page()

    btn_frame = Frame(win)
    btn_frame.pack(fill="x", pady=4)
    prev_btn = Button(btn_frame, text="<< Prev", command=on_prev, width=12)
    prev_btn.pack(side="left", padx=6)
    next_btn = Button(btn_frame, text="Next >>", command=on_next, width=12)
    next_btn.pack(side="left", padx=6)

    # initial render
    render_current_page()

    def on_close():
        try:
            doc.close()
        except Exception:
            pass
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", on_close)


# ---------- GUI Application ----------
class App(Frame):
    def __init__(self, root):
        Frame.__init__(self, root)
        self.root = root
        self.root.title("CommentPdfFromExcel - GUI")
        self.grid(sticky=(N, S, E, W))
        self.create_widgets()

        # Data holders
        self.excel_path = ""
        self.pdf_paths = []
        self.folder_mode = IntVar(value=1)  # 1 = folder, 0 = files
        self.output_folder = ""
        self.subject = StringVar(value="Comment")
        self.distance = IntVar(value=10)

    def create_widgets(self):
        row = 0

        Label(self, text="Excel File (with 'tag' and 'comment' columns):").grid(column=0, row=row, sticky=W, padx=5, pady=5)
        self.excel_entry = Entry(self, width=60)
        self.excel_entry.grid(column=1, row=row, columnspan=2, sticky=(W, E), padx=5)
        Button(self, text="Browse...", command=self.browse_excel).grid(column=3, row=row, padx=5)
        row += 1

        Label(self, text="Input PDFs:").grid(column=0, row=row, sticky=W, padx=5, pady=5)
        Radiobutton(self, text="Process folder", variable=self.folder_mode, value=1, command=self.update_input_mode).grid(column=1, row=row, sticky=W)
        Radiobutton(self, text="Select individual PDF(s)", variable=self.folder_mode, value=0, command=self.update_input_mode).grid(column=2, row=row, sticky=W)
        row += 1

        self.input_entry = Entry(self, width=60)
        self.input_entry.grid(column=1, row=row, columnspan=2, sticky=(W, E), padx=5)
        self.input_button = Button(self, text="Browse...", command=self.browse_input)
        self.input_button.grid(column=3, row=row, padx=5)
        row += 1

        Label(self, text="Output folder:").grid(column=0, row=row, sticky=W, padx=5, pady=5)
        self.output_entry = Entry(self, width=60)
        self.output_entry.grid(column=1, row=row, columnspan=2, sticky=(W, E), padx=5)
        Button(self, text="Browse...", command=self.browse_output).grid(column=3, row=row, padx=5)
        row += 1

        Label(self, text="Comment subject:").grid(column=0, row=row, sticky=W, padx=5, pady=5)
        self.subject_entry = Entry(self, textvariable=self.subject, width=30)
        self.subject_entry.grid(column=1, row=row, sticky=W, padx=5)

        Label(self, text="Distance (points):").grid(column=2, row=row, sticky=W, padx=5, pady=5)
        self.distance_entry = Entry(self, textvariable=self.distance, width=10)
        self.distance_entry.grid(column=3, row=row, sticky=W, padx=5)
        row += 1

        Button(self, text="Preview", command=self.preview_sample, width=12).grid(column=1, row=row, padx=5, pady=10)
        Button(self, text="Start", command=self.start_processing, width=12).grid(column=2, row=row, padx=5, pady=10)
        Button(self, text="Quit", command=self.root.quit, width=12).grid(column=3, row=row, padx=5, pady=10)
        row += 1

        Label(self, text="Log:").grid(column=0, row=row, sticky=NW, padx=5)
        self.log = Text(self, width=90, height=15)
        self.log.grid(column=0, row=row+1, columnspan=4, padx=5, pady=(0,10))
        self.log.configure(state=NORMAL)

        # Configure grid resizing
        for c in range(4):
            self.grid_columnconfigure(c, weight=1)

        # initial mode
        self.update_input_mode()

    def update_input_mode(self):
        mode = self.folder_mode.get()
        if mode == 1:
            self.input_entry.delete(0, END)
            self.input_entry.insert(0, "")
            self.input_entry.config(state=NORMAL)
            self.input_button.config(text="Browse Folder...")
        else:
            self.input_entry.delete(0, END)
            self.input_button.config(text="Browse PDF(s)...")

    def browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            self.excel_path = path
            self.excel_entry.delete(0, END)
            self.excel_entry.insert(0, path)

    def browse_input(self):
        if self.folder_mode.get() == 1:
            folder = filedialog.askdirectory(title="Select Folder with PDF Files")
            if folder:
                self.input_entry.delete(0, END)
                self.input_entry.insert(0, folder)
                # set pdf_paths to all pdfs in folder
                self.pdf_paths = [
                    os.path.join(folder, f)
                    for f in os.listdir(folder)
                    if f.lower().endswith(".pdf")
                ]
        else:
            files = filedialog.askopenfilenames(
                title="Select PDF File(s)", filetypes=[("PDF files", "*.pdf")]
            )
            if files:
                self.pdf_paths = list(files)
                self.input_entry.delete(0, END)
                self.input_entry.insert(0, ", ".join([os.path.basename(p) for p in self.pdf_paths]))

    def browse_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder = folder
            self.output_entry.delete(0, END)
            self.output_entry.insert(0, folder)

    def append_log(self, msg):
        self.log.insert(END, msg + "\n")
        self.log.see(END)
        self.log.update_idletasks()

    def start_processing(self):
        # validate inputs
        excel = self.excel_entry.get().strip()
        if not excel or not os.path.isfile(excel):
            messagebox.showerror("Input error", "Please select a valid Excel file.")
            return

        if not self.pdf_paths:
            # if folder mode and input entry contains a folder path, try to load
            if self.folder_mode.get() == 1:
                folder = self.input_entry.get().strip()
                if folder and os.path.isdir(folder):
                    self.pdf_paths = [
                        os.path.join(folder, f)
                        for f in os.listdir(folder)
                        if f.lower().endswith(".pdf")
                    ]
            if not self.pdf_paths:
                messagebox.showerror("Input error", "Please select PDF files or a folder containing PDFs.")
                return

        out_folder = self.output_entry.get().strip()
        if not out_folder:
            # default to same folder as first PDF
            out_folder = os.path.dirname(self.pdf_paths[0])
            self.output_entry.insert(0, out_folder)

        try:
            dist = int(self.distance_entry.get())
            if dist < 0:
                raise ValueError("Distance must be >= 0")
        except Exception:
            messagebox.showerror("Input error", "Please enter a valid non-negative integer for distance.")
            return

        subj = self.subject_entry.get().strip() or "Comment"

        # disable UI while running
        self.disable_ui()
        self.append_log("Starting processing...")
        try:
            process_files(
                self.pdf_paths,
                excel,
                out_folder,
                subject=subj,
                distance=dist,
                log_func=self.append_log,
            )
            self.append_log("All done.")
            messagebox.showinfo("Success", f"Processed {len(self.pdf_paths)} PDF file(s).\nSaved to: {out_folder}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}")
            self.append_log(f"Error: {e}")
        finally:
            self.enable_ui()

    def preview_sample(self):
        # validate excel and PDF selection, then show preview for first PDF
        excel = self.excel_entry.get().strip()
        if not excel or not os.path.isfile(excel):
            messagebox.showerror("Input error", "Please select a valid Excel file before previewing.")
            return

        # load df
        try:
            df = pd.read_excel(excel)
        except Exception as e:
            messagebox.showerror("Input error", f"Failed to read Excel file: {e}")
            return

        if "tag" not in df.columns or "comment" not in df.columns:
            messagebox.showerror("Input error", "Excel must contain 'tag' and 'comment' columns.")
            return

        df["tag"] = df["tag"].astype(str)
        df["comment"] = df["comment"].astype(str)

        # ensure we have at least one PDF path
        if not self.pdf_paths:
            # try to infer from input entry if folder mode
            if self.folder_mode.get() == 1:
                folder = self.input_entry.get().strip()
                if folder and os.path.isdir(folder):
                    self.pdf_paths = [
                        os.path.join(folder, f)
                        for f in os.listdir(folder)
                        if f.lower().endswith(".pdf")
                    ]
        if not self.pdf_paths:
            messagebox.showerror("Input error", "Please select at least one PDF (or a folder with PDFs) to preview.")
            return

        # use the first PDF as the sample
        sample_pdf = self.pdf_paths[0]
        dist = 10
        try:
            dist = int(self.distance_entry.get())
            if dist < 0:
                dist = 10
        except Exception:
            dist = 10

        subj = self.subject_entry.get().strip() or "Comment"

        # show preview window
        show_preview_window(self.root, sample_pdf, df, subj, dist)

    def disable_ui(self):
        for child in self.winfo_children():
            try:
                child.configure(state=DISABLED)
            except Exception:
                pass

    def enable_ui(self):
        for child in self.winfo_children():
            try:
                child.configure(state=NORMAL)
            except Exception:
                pass


def main():
    root = Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
