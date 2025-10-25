import fitz  # PyMuPDF
import pandas as pd
import os
import io
from tkinter import (
    Tk,
    StringVar,
    IntVar,
    Label,
    Entry,
    Button,
    Radiobutton,
    Text,
    OptionMenu,
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
    NW,
)
from tkinter.ttk import Frame

# Pillow is optional but required for preview mode and accurate text metric measurements
try:
    from PIL import Image, ImageDraw, ImageTk, ImageFont
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False


# Map friendly font names to PDF "standard" font resource names and TTF candidates for preview/measurement
PDF_FONT_MAP = {
    "DejaVuSans": ("helv", ["DejaVuSans.ttf", "DejaVuSans.otf"]),
    "Arial": ("helv", ["arial.ttf", "Arial.ttf"]),
    "Times New Roman": ("times", ["Times New Roman.ttf", "Times_New_Roman.ttf", "Times.ttf"]),
    "Courier": ("cour", ["Courier.ttf", "cour.ttf"]),
}


def compute_text_size_points(text, fontsize, ttf_candidates=None, pdf_fontname="helv"):
    """
    Return (width_pts, height_pts, ascent_pts, descent_pts) for the given text and font size.

    Strategy:
    1) Try fitz.get_text_length(text, fontsize, pdf_fontname) to get a width in points that
       matches PyMuPDF annotation font metrics (best for annotation sizing).
    2) If that's not available or fails, fall back to Pillow/TTF measurement (pixel==point at 72dpi).
    3) Last resort: heuristic.

    Adds a small safety multiplier to the measured width to avoid clipping in viewers that may slightly
    vary metrics.
    """
    # 3) heuristic defaults (conservative estimates to avoid clipping)
    approx_w = max(10.0, len(text) * (fontsize * 0.6))
    approx_h = max(12.0, fontsize * 1.2)
    # Typography conventions: ascent ≈ 0.8 of font size, descent ≈ 0.2
    approx_ascent = fontsize * 0.8
    approx_descent = fontsize * 0.2

    # 1) try fitz.get_text_length with pdf_fontname (preferred)
    try:
        # fitz.get_text_length returns width in points for given fontsize and font name
        # Correct parameter order: get_text_length(text, fontname, fontsize, encoding)
        w = fitz.get_text_length(text, fontname=pdf_fontname, fontsize=fontsize)
        if w and w > 0:
            # add a safety margin (5%) to avoid clipping when viewer metrics differ slightly
            w = float(w) * 1.05
            h = float(max(12.0, fontsize * 1.2))
            # Typography conventions: ascent ≈ 0.8 of font size, descent ≈ 0.2
            ascent = fontsize * 0.8
            descent = fontsize * 0.2
            return float(w), h, ascent, descent
    except Exception:
        # not available or failed -> continue to Pillow fallback
        pass

    # 2) Pillow measurement (best-effort; assumes TTF available)
    if PIL_AVAILABLE:
        font = None
        if ttf_candidates:
            for fn in ttf_candidates:
                try:
                    font = ImageFont.truetype(fn, size=int(fontsize))
                    break
                except Exception:
                    font = None
        if font is None:
            try:
                font = ImageFont.truetype("DejaVuSans.ttf", size=int(fontsize))
            except Exception:
                font = ImageFont.load_default()

        try:
            # create a temp image large enough to measure
            img = Image.new("RGB", (4000, 800), (255, 255, 255))
            draw = ImageDraw.Draw(img)
            # use textbbox for accurate metrics
            bbox = draw.textbbox((0, 0), text, font=font)
            w_px = bbox[2] - bbox[0]
            h_px = bbox[3] - bbox[1]
            try:
                ascent, descent = font.getmetrics()
                ascent = float(ascent)
                descent = float(descent)
            except Exception:
                ascent = h_px * 0.75
                descent = h_px * 0.25
            # safety margin
            return float(w_px) * 1.05, float(h_px), ascent, descent
        except Exception:
            pass

    # fallback heuristic
    return approx_w, approx_h, approx_ascent, approx_descent


def update_pdf_with_comments(
    pdf_path,
    df,
    output_pdf_path,
    subject="Comment",
    distance=10,
    log_func=None,
    font_family="Arial",
    font_size=12,
):
    """
    Create freetext annotations (editable) and size them to the measured text metrics
    so the box fits the entire comment horizontally (no wrapping).
    """
    if log_func:
        log_func(f"Processing: {os.path.basename(pdf_path)}")

    # Map to PDF font resource and ttf candidates for measurement/preview
    pdf_fontname, ttf_candidates = PDF_FONT_MAP.get(font_family, ("helv", ["DejaVuSans.ttf"]))

    doc = fitz.open(pdf_path)

    annotation_count = 0
    for page_num in range(len(doc)):
        page = doc[page_num]
        content = page.get_text("text")

        for index, row in df.iterrows():
            tag = str(row["tag"])
            comment = str(row["comment"])

            if not tag or tag.strip() == "":
                continue

            if tag in content:
                text_instances = page.search_for(tag)
                for inst in text_instances:
                    # Measure comment text in points using fitz metrics first (pdf_fontname), then PIL fallback
                    text_w_pts, text_h_pts, ascent_pts, descent_pts = compute_text_size_points(
                        comment, font_size, ttf_candidates, pdf_fontname
                    )

                    # horizontal and vertical padding to ensure no clipping
                    padding_x = max(8.0, font_size * 0.5)
                    padding_y = max(4.0, font_size * 0.25)

                    width = text_w_pts + 2.0 * padding_x
                    measured_text_height = ascent_pts + descent_pts if (ascent_pts and descent_pts) else text_h_pts
                    height = max(12.0, measured_text_height + 2.0 * padding_y)

                    page_rect = page.rect

                    # Preferred position: to the right of the tag, offset by distance
                    pref_x0 = inst.x1 + distance
                    pref_x1 = pref_x0 + width

                    if pref_x1 <= page_rect.x1 - 5:
                        x0 = pref_x0
                        x1 = pref_x1
                    else:
                        # place left of tag
                        x1 = inst.x0 - distance
                        x0 = x1 - width
                        if x0 < page_rect.x0 + 5:
                            # clamp and reduce width if necessary
                            x0 = page_rect.x0 + 5
                            x1 = min(page_rect.x1 - 5, x0 + width)

                    # Vertical placement: center the annotation vertically relative to the tag instance
                    inst_mid = (inst.y0 + inst.y1) / 2.0
                    y0 = inst_mid - (height / 2.0)
                    y1 = y0 + height

                    # Clamp to page bounds
                    if y0 < page_rect.y0 + 5:
                        y0 = page_rect.y0 + 5
                        y1 = y0 + height
                    if y1 > page_rect.y1 - 5:
                        y1 = page_rect.y1 - 5
                        y0 = y1 - height
                        if y0 < page_rect.y0 + 5:
                            y0 = page_rect.y0 + 5

                    rect = fitz.Rect(x0, y0, x1, y1)

                    # Create freetext annotation and set font/size; use best-effort methods to influence appearance
                    try:
                        annot = page.add_freetext_annot(
                            rect,
                            comment,
                            fontsize=font_size,
                            text_color=(0, 0, 0),
                            fill_color=(1, 1, 0),
                            rotate=0,
                            align=fitz.TEXT_ALIGN_LEFT,
                        )

                        # Best-effort: set annotation font resource name (viewer/PyMuPDF dependent)
                        try:
                            annot.set_font(pdf_fontname)
                        except Exception:
                            try:
                                annot.set_font("helv")
                            except Exception:
                                pass

                        # set border and colors when available
                        try:
                            annot.set_border(width=0.5, dashes=[2])
                        except Exception:
                            pass
                        try:
                            annot.set_colors(stroke=(0, 0, 0), fill=(1, 1, 0))
                        except Exception:
                            pass
                        try:
                            annot.set_info({"subject": subject})
                        except Exception:
                            pass

                        annot.update()
                        annotation_count += 1
                        if log_func:
                            log_func(
                                f"  Added freetext annot on page {page_num+1} at {rect} (font={font_family}, size={font_size})"
                            )
                    except Exception as e:
                        if log_func:
                            log_func(f"  Error creating freetext annot at {rect}: {e}")

    # Save without flattening so annotations remain editable
    doc.save(output_pdf_path)
    doc.close()
    if log_func:
        log_func(f"Saved: {os.path.basename(output_pdf_path)} (Total annotations: {annotation_count})")


def process_files(
    pdf_paths,
    excel_path,
    output_folder,
    subject="Comment",
    distance=10,
    log_func=None,
    font_family="Arial",
    font_size=12,
):
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        raise RuntimeError(f"Failed to read Excel file: {e}")

    if "tag" not in df.columns or "comment" not in df.columns:
        raise RuntimeError("Excel must contain 'tag' and 'comment' columns.")

    df["tag"] = df["tag"].astype(str)
    df["comment"] = df["comment"].astype(str)

    os.makedirs(output_folder, exist_ok=True)

    for pdf_path in pdf_paths:
        base = os.path.basename(pdf_path)
        name, ext = os.path.splitext(base)
        output_pdf_path = os.path.join(output_folder, f"{name}_marked{ext}")
        try:
            update_pdf_with_comments(
                pdf_path,
                df,
                output_pdf_path,
                subject=subject,
                distance=distance,
                log_func=log_func,
                font_family=font_family,
                font_size=font_size,
            )
        except Exception as e:
            if log_func:
                log_func(f"Error processing {base}: {e}")
            continue


# ---------- Preview utilities ----------
def render_page_pil_from_pixmap(pix):
    png_bytes = pix.tobytes("png")
    return Image.open(io.BytesIO(png_bytes)).convert("RGBA")


def get_text_size(draw, text, font):
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
        return w, h
    except Exception:
        pass
    try:
        w, h = draw.textsize(text, font=font)
        return w, h
    except Exception:
        pass
    try:
        w, h = font.getsize(text)
        return w, h
    except Exception:
        return (len(text) * 7, int(getattr(font, "size", 12)))


def build_annotations_for_preview(page, df, distance, font_family="Arial", font_size=12):
    annotations = []
    content = page.get_text("text")

    _, ttf_candidates = PDF_FONT_MAP.get(font_family, ("helv", ["DeJaVuSans.ttf"]))

    for index, row in df.iterrows():
        tag = str(row["tag"])
        comment = str(row["comment"])
        if not tag or tag.strip() == "":
            continue

        if tag in content:
            text_instances = page.search_for(tag)
            for inst in text_instances:
                text_w_pts, text_h_pts, ascent_pts, descent_pts = compute_text_size_points(
                    comment, font_size, ttf_candidates, pdf_fontname="helv"
                )
                padding_x = max(8.0, font_size * 0.5)
                padding_y = max(4.0, font_size * 0.25)
                width = text_w_pts + 2.0 * padding_x
                measured_text_height = ascent_pts + descent_pts if (ascent_pts and descent_pts) else text_h_pts
                height = max(12.0, measured_text_height + 2.0 * padding_y)

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

                inst_mid = (inst.y0 + inst.y1) / 2.0
                y0 = inst_mid - (height / 2.0)
                y1 = y0 + height
                if y0 < page_rect.y0 + 5:
                    y0 = page_rect.y0 + 5
                    y1 = y0 + height
                if y1 > page_rect.y1 - 5:
                    y1 = page_rect.y1 - 5
                    y0 = y1 - height
                    if y0 < page_rect.y0 + 5:
                        y0 = page_rect.y0 + 5

                annot_rect = fitz.Rect(x0, y0, x1, y1)
                annotations.append(
                    {
                        "annot_rect": annot_rect,
                        "comment": comment,
                        "inst_rect": inst,
                        "tag": tag,
                    }
                )
    return annotations


def show_preview_snippet(parent, pdf_path, df, subject, distance, font_family, font_size):
    if not PIL_AVAILABLE:
        messagebox.showerror(
            "Preview unavailable",
            "Pillow is required for preview mode. Install it with: pip install pillow",
        )
        return

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        messagebox.showerror("Preview error", f"Failed to open PDF: {e}")
        return

    if len(doc) == 0:
        messagebox.showwarning("Preview", "PDF has no pages.")
        doc.close()
        return

    first_found = None
    first_page_index = None
    first_annotation = None

    for i in range(len(doc)):
        page = doc[i]
        anns = build_annotations_for_preview(page, df, distance, font_family=font_family, font_size=font_size)
        if anns:
            first_found = page
            first_page_index = i
            first_annotation = anns[0]
            break

    if first_found is None:
        messagebox.showinfo("Preview", "No tags found in the PDF to preview.")
        doc.close()
        return

    zoom = 2.0
    try:
        pix = first_found.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        pil_full = render_page_pil_from_pixmap(pix)
    except Exception as e:
        messagebox.showerror("Preview error", f"Failed to render page for preview: {e}")
        doc.close()
        return

    annot_rect = first_annotation["annot_rect"]
    inst_rect = first_annotation["inst_rect"]
    comment = first_annotation["comment"]
    tag = first_annotation["tag"]

    # conversion to pixels (pixmap scaled by zoom)
    x0 = int(annot_rect.x0 * zoom)
    y0 = int(annot_rect.y0 * zoom)
    x1 = int(annot_rect.x1 * zoom)
    y1 = int(annot_rect.y1 * zoom)
    margin = int(28 * zoom)

    ix0 = int(inst_rect.x0 * zoom)
    iy0 = int(inst_rect.y0 * zoom)
    ix1 = int(inst_rect.x1 * zoom)
    iy1 = int(inst_rect.y1 * zoom)

    cx0 = max(0, min(x0, ix0) - margin)
    cy0 = max(0, min(y0, iy0) - margin)
    cx1 = min(pil_full.width, max(x1, ix1) + margin)
    cy1 = min(pil_full.height, max(y1, iy1) + margin)

    snippet = pil_full.crop((cx0, cy0, cx1, cy1)).convert("RGBA")
    draw = ImageDraw.Draw(snippet)

    # Load font at scaled size so preview shows correct visual size
    _, ttf_candidates = PDF_FONT_MAP.get(font_family, ("helv", ["DejaVuSans.ttf"]))
    font_obj = None
    for fn in ttf_candidates:
        try:
            font_obj = ImageFont.truetype(fn, size=int(font_size * zoom))
            break
        except Exception:
            font_obj = None
    if font_obj is None:
        try:
            font_obj = ImageFont.truetype("DejaVuSans.ttf", size=int(font_size * zoom))
        except Exception:
            font_obj = ImageFont.load_default()

    r_ax0 = x0 - cx0
    r_ay0 = y0 - cy0
    r_ax1 = x1 - cx0
    r_ay1 = y1 - cy0

    r_ix0 = ix0 - cx0
    r_iy0 = iy0 - cy0
    r_ix1 = ix1 - cx0
    r_iy1 = iy1 - cy0

    # draw annotation area and tag area
    try:
        draw.rectangle([r_ax0, r_ay0, r_ax1, r_ay1], fill=(255, 255, 0, 200), outline=(0, 0, 0))
    except Exception:
        draw.rectangle([r_ax0, r_ay0, r_ax1, r_ay1], outline=(0, 0, 0))
    draw.rectangle([r_ix0, r_iy0, r_ix1, r_iy1], outline=(0, 120, 200), width=2)

    # Draw full comment text (no trimming, no wrapping). The annotation RECT is sized to fit.
    draw.text((r_ax0 + int(3 * zoom), r_ay0 + int(2 * zoom)), comment, fill=(0, 0, 0), font=font_obj)

    # Show in Toplevel
    win = Toplevel(parent)
    win.title(f"Preview - {os.path.basename(pdf_path)} (page {first_page_index+1})")
    win.geometry("700x420")
    win.minsize(320, 200)

    img_frame = Frame(win)
    img_frame.pack(expand=True, fill="both", padx=6, pady=6)

    img_label = Label(img_frame)
    img_label.pack(expand=True, fill="both")

    win.update_idletasks()
    avail_w = max(200, win.winfo_width() - 40)
    avail_h = max(120, win.winfo_height() - 120)
    ratio = min(avail_w / snippet.width, avail_h / snippet.height, 1.0)
    if ratio < 1.0:
        display_img = snippet.resize(
            (int(snippet.width * ratio), int(snippet.height * ratio)), Image.LANCZOS
        )
    else:
        display_img = snippet

    try:
        photo = ImageTk.PhotoImage(display_img.convert("RGB"))
    except Exception as e:
        messagebox.showerror("Preview error", f"Failed to build preview image: {e}")
        doc.close()
        win.destroy()
        return

    win._photo = photo
    img_label.config(image=photo)

    info = Label(win, text=f"First replacement on page {first_page_index+1}", anchor="w")
    info.pack(fill="x", padx=6, pady=(0, 6))

    def on_close():
        try:
            doc.close()
        except Exception:
            pass
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", on_close)


# ---------- GUI Application ----------
FONT_CHOICES = ["Arial", "DejaVuSans", "Times New Roman", "Courier"]


class App(Frame):
    def __init__(self, root):
        Frame.__init__(self, root)
        self.root = root
        self.root.title("CommentPdfFromExcel - GUI")
        self.grid(sticky=(N, S, E, W))

        # Data holders
        self.excel_path = ""
        self.pdf_paths = []
        self.folder_mode = IntVar(value=1)
        self.output_folder = ""
        self.subject = StringVar(value="Comment")
        self.distance = IntVar(value=10)
        self.font_family = StringVar(value="Arial")
        self.font_size = IntVar(value=12)

        self.preview_button = None
        self.start_button = None
        self.quit_button = None

        self.create_widgets()
        self.append_log("Ready")

    def create_widgets(self):
        row = 0

        Label(self, text="Excel File (with 'tag' and 'comment' columns):").grid(
            column=0, row=row, sticky=W, padx=5, pady=5
        )
        self.excel_entry = Entry(self, width=60)
        self.excel_entry.grid(column=1, row=row, columnspan=2, sticky=(W, E), padx=5)
        Button(self, text="Browse...", command=self.browse_excel).grid(column=3, row=row, padx=5)
        row += 1

        Label(self, text="Input PDFs:").grid(column=0, row=row, sticky=W, padx=5, pady=5)
        Radiobutton(
            self, text="Process folder", variable=self.folder_mode, value=1, command=self.update_input_mode
        ).grid(column=1, row=row, sticky=W)
        Radiobutton(
            self,
            text="Select individual PDF(s)",
            variable=self.folder_mode,
            value=0,
            command=self.update_input_mode,
        ).grid(column=2, row=row, sticky=W)
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

        # Font controls
        Label(self, text="Font:").grid(column=0, row=row, sticky=W, padx=5, pady=5)
        font_menu = OptionMenu(self, self.font_family, *FONT_CHOICES)
        font_menu.grid(column=1, row=row, sticky=W, padx=5)
        Label(self, text="Size:").grid(column=2, row=row, sticky=W, padx=5)
        self.font_size_entry = Entry(self, textvariable=self.font_size, width=6)
        self.font_size_entry.grid(column=3, row=row, sticky=W, padx=5)
        row += 1

        self.preview_button = Button(self, text="Preview", command=self.preview_sample, width=12)
        self.preview_button.grid(column=1, row=row, padx=5, pady=10)
        self.start_button = Button(self, text="Start", command=self.start_processing, width=12)
        self.start_button.grid(column=2, row=row, padx=5, pady=10)
        self.quit_button = Button(self, text="Quit", command=self.root.destroy, width=12)
        self.quit_button.grid(column=3, row=row, padx=5, pady=10)
        row += 1

        Label(self, text="Log:").grid(column=0, row=row, sticky=NW, padx=5)
        self.log = Text(self, width=90, height=15)
        self.log.grid(column=0, row=row + 1, columnspan=4, padx=5, pady=(0, 10))
        self.log.configure(state=DISABLED)

        for c in range(4):
            self.grid_columnconfigure(c, weight=1)

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
            self.append_log(f"Excel selected: {os.path.basename(path)}")

    def browse_input(self):
        if self.folder_mode.get() == 1:
            folder = filedialog.askdirectory(title="Select Folder with PDF Files")
            if folder:
                self.input_entry.delete(0, END)
                self.input_entry.insert(0, folder)
                self.pdf_paths = [
                    os.path.join(folder, f)
                    for f in os.listdir(folder)
                    if f.lower().endswith(".pdf")
                ]
                self.append_log(f"Loaded {len(self.pdf_paths)} PDF(s) from folder.")
        else:
            files = filedialog.askopenfilenames(
                title="Select PDF File(s)", filetypes=[("PDF files", "*.pdf")]
            )
            if files:
                self.pdf_paths = list(files)
                self.input_entry.delete(0, END)
                if len(self.pdf_paths) == 1:
                    self.input_entry.insert(0, self.pdf_paths[0])
                else:
                    self.input_entry.insert(0, "; ".join(self.pdf_paths))
                self.append_log(f"Selected {len(self.pdf_paths)} PDF(s).")

    def browse_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder = folder
            self.output_entry.delete(0, END)
            self.output_entry.insert(0, folder)
            self.append_log(f"Output folder: {folder}")

    def append_log(self, msg):
        try:
            self.log.configure(state=NORMAL)
            self.log.insert(END, msg + "\n")
            self.log.see(END)
            self.log.update_idletasks()
        finally:
            self.log.configure(state=DISABLED)

    def start_processing(self):
        excel = self.excel_entry.get().strip()
        if not excel or not os.path.isfile(excel):
            messagebox.showerror("Input error", "Please select a valid Excel file.")
            return

        if not self.pdf_paths:
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
            out_folder = os.path.dirname(self.pdf_paths[0])
            self.output_entry.insert(0, out_folder)

        try:
            dist = int(self.distance_entry.get())
            if dist < 0:
                raise ValueError("Distance must be >= 0")
        except Exception:
            messagebox.showerror("Input error", "Please enter a valid non-negative integer for distance.")
            return

        try:
            fsize = int(self.font_size_entry.get())
            if fsize <= 0:
                raise ValueError()
        except Exception:
            messagebox.showerror("Input error", "Please enter a valid positive integer for font size.")
            return

        subj = self.subject_entry.get().strip() or "Comment"
        ffamily = self.font_family.get() or "DejaVuSans"

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
                font_family=ffamily,
                font_size=fsize,
            )
            self.append_log("All done.")
            messagebox.showinfo("Success", f"Processed {len(self.pdf_paths)} PDF file(s).\nSaved to: {out_folder}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}")
            self.append_log(f"Error: {e}")
        finally:
            self.enable_ui()

    def preview_sample(self):
        excel = self.excel_entry.get().strip()
        if not excel or not os.path.isfile(excel):
            messagebox.showerror("Input error", "Please select a valid Excel file before previewing.")
            return

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

        if not self.pdf_paths:
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

        sample_pdf = self.pdf_paths[0]
        try:
            dist = int(self.distance_entry.get())
            if dist < 0:
                dist = 10
        except Exception:
            dist = 10

        try:
            fsize = int(self.font_size_entry.get())
            if fsize <= 0:
                fsize = 12
        except Exception:
            fsize = 12

        ffamily = self.font_family.get() or "DejaVuSans"
        subj = self.subject_entry.get().strip() or "Comment"

        self.append_log(f"Showing preview snippet for: {os.path.basename(sample_pdf)}")
        show_preview_snippet(self.root, sample_pdf, df, subj, dist, ffamily, fsize)

    def disable_ui(self):
        widgets_to_disable = [
            self.excel_entry,
            self.input_entry,
            self.input_button,
            self.output_entry,
            self.subject_entry,
            self.distance_entry,
            self.preview_button,
            self.start_button,
            self.font_size_entry,
        ]
        for w in widgets_to_disable:
            try:
                w.configure(state=DISABLED)
            except Exception:
                pass

    def enable_ui(self):
        widgets_to_enable = [
            self.excel_entry,
            self.input_entry,
            self.input_button,
            self.output_entry,
            self.subject_entry,
            self.distance_entry,
            self.preview_button,
            self.start_button,
            self.font_size_entry,
        ]
        for w in widgets_to_enable:
            try:
                w.configure(state=NORMAL)
            except Exception:
                pass


def main():
    root = Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
