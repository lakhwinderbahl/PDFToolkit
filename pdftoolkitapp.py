
import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
from typing import Optional, List

import pandas as pd

import fitz  # PyMuPDF, used for PDF compression
from pdf2docx import Converter
import img2pdf
from PIL import Image, ImageTk
from pdf2image import convert_from_path
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import pdfplumber

# If Tesseract is installed on the system, uncomment and adjust the path
# import pytesseract
# pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"


class PDFToolkitApp:
    """A Tkinter GUI application providing various PDF and document utilities."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PDF Toolkit")
        self.root.geometry("900x600")
        self.root.minsize(600, 400)

        # Track whether dark mode is active
        self.is_dark_mode = False

        # Define modern color palettes for light and dark themes.  These can be
        # tweaked to taste.  Light theme uses bright backgrounds with a blue
        # accent, while the dark theme uses deep greys with a lighter blue.
        self.light_theme = {
            "sidebar_bg": "#f5f7fa",
            "content_bg": "#ffffff",
            "button_bg": "#4f80ff",
            "button_fg": "#ffffff",
            "label_fg": "#444444",
            "canvas_bg": "#ffffff",
            "progress_bg": "#4f80ff",
            "header_fg": "#4f80ff",
        }
        self.dark_theme = {
            "sidebar_bg": "#1f1f1f",
            "content_bg": "#2e2e2e",
            "button_bg": "#3b6cb9",
            "button_fg": "#ffffff",
            "label_fg": "#dddddd",
            "canvas_bg": "#2e2e2e",
            "progress_bg": "#3b6cb9",
            "header_fg": "#81a1c1",
        }

        # Load icons for light/dark mode toggle
        # Note: these image files (sun.png and moon.png) need to exist in the
        # working directory.  Resize them for the button.
        self.sun_image = ImageTk.PhotoImage(Image.open("sun.png").resize((24, 24)))
        self.moon_image = ImageTk.PhotoImage(Image.open("moon.png").resize((24, 24)))

        # Create sidebar for actions
        self.sidebar = tk.Frame(root, bg=self.light_theme["sidebar_bg"], width=200)
        self.sidebar.pack(side="left", fill="y")

        # Create main content area for previews and status
        self.content = tk.Frame(root, bg=self.light_theme["content_bg"])
        self.content.pack(side="right", expand=True, fill="both")

        # Toggle button for switching themes
        self.toggle_btn = tk.Label(
            self.sidebar,
            image=self.sun_image,
            cursor="hand2",
            bg="#f0f0f0",
        )
        self.toggle_btn.pack(pady=10, anchor="ne")
        self.toggle_btn.bind("<Button-1>", self.toggle_theme)
        # Create a simple tooltip for the toggle button
        self.toggle_btn.tooltip = tk.Label(
            self.sidebar,
            text="Change Mode",
            bg="black",
            fg="white",
            font=("Segoe UI", 8),
            bd=1,
            relief="solid",
        )
        self.toggle_btn.bind(
            "<Enter>", lambda e: self.toggle_btn.tooltip.place(x=50, y=10)
        )
        self.toggle_btn.bind(
            "<Leave>", lambda e: self.toggle_btn.tooltip.place_forget()
        )

        # Buttons for available actions; each corresponds to a method below
        self.buttons = []
        actions = [
            ("Excel to PDF", self.excel_to_pdf),
            ("PDF to Excel", self.pdf_to_excel),
            ("PDF to Word", self.pdf_to_word),
            ("Compress PDF", self.compress_pdf),
            ("Compress Images", self.compress_images),
            ("Image to PDF", self.image_to_pdf),
            ("Merge PDFs", self.merge_pdfs),
            ("Split PDF", self.split_pdf),
            ("PDF to Images", self.pdf_to_images),
            ("Batch Compress PDFs", self.batch_compress),
            ("Merge Images to PDF", self.merge_images_to_pdf),
            # New functionality: extract all text from a PDF into a plain-text file
            ("PDF to Text", self.pdf_to_text),
        ]
        for text, command in actions:
            frame = tk.Frame(self.sidebar, bg=self.sidebar["bg"])
            btn = tk.Button(
                frame,
                text=text,
                command=command,
                font=("Segoe UI", 10, "bold"),
                bg=self.light_theme["button_bg"],
                fg=self.light_theme["button_fg"],
                bd=0,
                relief="flat",
                cursor="hand2",
                activebackground=self.light_theme["button_bg"],
                activeforeground=self.light_theme["button_fg"],
            )
            btn.pack(fill="x")
            frame.pack(pady=5, padx=10, fill="x")
            self.buttons.append(btn)

        # Status label to display currently selected file or operation status
        self.status_label = tk.Label(
            self.content,
            text="No file selected",
            fg=self.light_theme["label_fg"],
            bg=self.light_theme["content_bg"],
            anchor="w",
        )

        # Header label for application title
        self.header_label = tk.Label(
            self.content,
            text="PDF Toolkit",
            font=("Segoe UI", 16, "bold"),
            fg=self.light_theme["header_fg"],
            bg=self.light_theme["content_bg"],
        )
        self.header_label.pack(pady=(10, 0))
        self.status_label.pack(fill="x", padx=10, pady=5)

        # Canvas to display PDF page previews or images
        self.preview_canvas = tk.Canvas(self.content, bg=self.light_theme["canvas_bg"])
        self.preview_canvas.pack(pady=10, expand=True, fill="both")

        # Create a bottom frame to hold the progress bar and clear-preview button
        # This positions the clear button next to the progress bar.
        self.bottom_frame = tk.Frame(
            self.content,
            bg=self.light_theme["content_bg"],
        )
        # Progress bar to indicate long running tasks; default to indeterminate
        self.progress = ttk.Progressbar(
            self.bottom_frame,
            orient="horizontal",
            mode="indeterminate",
            length=200,
        )
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 10))
        # Clear preview button next to the progress bar
        self.clear_btn = tk.Button(
            self.bottom_frame,
            text="Clear Preview",
            command=self.clear_preview,
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            cursor="hand2",
        )
        self.clear_btn.pack(side="right")
        # Pack the bottom frame
        self.bottom_frame.pack(pady=5, padx=10, fill="x")

        # Apply the initial theme settings
        self.set_theme()

    # ------------------------------------------------------------------
    # Theme management
    # ------------------------------------------------------------------
    def toggle_theme(self, event: Optional[tk.Event] = None) -> None:
        """Toggle between light and dark themes."""
        self.is_dark_mode = not self.is_dark_mode
        self.set_theme()

    def set_theme(self) -> None:
        """
        Apply the current theme to all UI elements.

        This method selects the appropriate color palette based on
        ``self.is_dark_mode`` and updates the backgrounds, foregrounds,
        button colours and progress bar style accordingly.  A custom
        ttk.Style is configured for the progress bar to match the theme.
        """
        # Choose theme palette
        theme = self.dark_theme if self.is_dark_mode else self.light_theme

        # Sidebar and content backgrounds
        self.sidebar.config(bg=theme["sidebar_bg"])
        self.content.config(bg=theme["content_bg"])

        # Header and status labels
        self.header_label.config(bg=theme["content_bg"], fg=theme["header_fg"])
        self.status_label.config(bg=theme["content_bg"], fg=theme["label_fg"])

        # Preview canvas
        self.preview_canvas.config(bg=theme["canvas_bg"])

        # Buttons styling
        for btn in self.buttons:
            btn.config(
                bg=theme["button_bg"],
                fg=theme["button_fg"],
                activebackground=theme["button_bg"],
                activeforeground=theme["button_fg"],
            )

        # Update the clear preview button and its container if they exist
        if hasattr(self, "clear_btn"):
            self.clear_btn.config(
                bg=theme["button_bg"],
                fg=theme["button_fg"],
                activebackground=theme["button_bg"],
                activeforeground=theme["button_fg"],
            )
        if hasattr(self, "bottom_frame"):
            self.bottom_frame.config(bg=theme["content_bg"])  # match content

        # Update the toggle button icon and background
        self.toggle_btn.config(
            image=self.moon_image if self.is_dark_mode else self.sun_image,
            bg=theme["sidebar_bg"],
        )

        # Configure a custom style for the progress bar to reflect the theme
        style = ttk.Style()
        # Use the default theme as base to ensure our colors apply
        try:
            style.theme_use("clam")
        except Exception:
            # Fallback if clam is not available
            style.theme_use("default")
        bar_style = "Modern.Horizontal.TProgressbar"
        style.configure(
            bar_style,
            troughcolor=theme["content_bg"],
            bordercolor=theme["content_bg"],
            background=theme["progress_bg"],
            lightcolor=theme["progress_bg"],
            darkcolor=theme["progress_bg"],
        )
        self.progress.config(style=bar_style)

    # ------------------------------------------------------------------
    # Action handlers
    # ------------------------------------------------------------------
    def excel_to_pdf(self) -> None:
        """
        Convert an Excel file to PDF.

        The original implementation converted only the first part of the
        DataFrame and ignored column widths, causing tables to overflow
        horizontally.  This rewritten method handles an entire workbook,
        computes relative column widths based on the content, and switches to
        landscape orientation if the sheet is wide.  Headers are repeated on
        each page for readability.
        """
        excel_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not excel_path:
            return
        self.status_label.config(text=f"Selected: {os.path.basename(excel_path)}")
        self.progress.start()
        self.root.update()

        try:
            # Load all sheets in the workbook; sheet_name=None returns a dict
            sheets_dict = pd.read_excel(excel_path, sheet_name=None)
            # Determine maximum number of columns across all sheets
            max_cols = max(len(df.columns) for df in sheets_dict.values())
            # Use landscape if there are many columns (>5), else portrait
            from reportlab.lib.pagesizes import letter, landscape
            from reportlab.lib.units import inch
            pagesize = landscape(letter) if max_cols > 5 else letter

            # Prepare the PDF file path
            pdf_path = os.path.splitext(excel_path)[0] + ".pdf"

            from reportlab.platypus import (
                SimpleDocTemplate,
                Table,
                TableStyle,
                PageBreak,
                Paragraph,
                Spacer,
            )
            from reportlab.lib import colors
            from reportlab.lib.styles import getSampleStyleSheet

            # Build a story (list of flowables) containing a table per sheet
            story: list = []
            styles = getSampleStyleSheet()
            for sheet_name, df in sheets_dict.items():
                # Ensure all values are strings to avoid issues with floats/NaNs
                df_str = df.fillna("").astype(str)
                # Compute maximum length (in characters) for each column,
                # including the header.  Use a minimum of 1 to avoid zero width.
                max_lengths = []
                for col in df_str.columns:
                    max_len = max(
                        [len(str(col))] + [len(s) for s in df_str[col].tolist()]
                    )
                    max_lengths.append(max(max_len, 1))

                total_length = sum(max_lengths)
                # Available page width excluding margins (0.5 inch each side)
                page_width = pagesize[0] - 2 * 0.5 * inch
                # Compute individual column widths proportional to content length
                col_widths = [
                    (length / total_length) * page_width for length in max_lengths
                ]

                # Construct table data with header row followed by data rows
                table_data = [list(df_str.columns)] + df_str.values.tolist()
                table = Table(table_data, colWidths=col_widths, repeatRows=1)

                # Apply a simple style: alternating row backgrounds and grid lines
                style = TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("FONTSIZE", (0, 0), (-1, 0), 9),
                        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                        ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ]
                )
                # Alternate row background
                for row_num in range(1, len(table_data)):
                    if row_num % 2 == 0:
                        style.add("BACKGROUND", (0, row_num), (-1, row_num), colors.lightgrey)
                table.setStyle(style)

                # Add sheet name as a simple header above the table
                header = Paragraph(f"<b>{sheet_name}</b>", styles["Heading2"])
                story.append(header)
                story.append(Spacer(1, 0.2 * inch))
                story.append(table)
                story.append(PageBreak())

            # If story is empty, bail out
            if not story:
                raise ValueError("No sheets to convert.")

            # Build the PDF
            doc = SimpleDocTemplate(
                pdf_path,
                pagesize=pagesize,
                leftMargin=0.5 * inch,
                rightMargin=0.5 * inch,
                topMargin=0.5 * inch,
                bottomMargin=0.5 * inch,
            )
            # Remove the final PageBreak so there isn't a blank page at the end
            if isinstance(story[-1], PageBreak):
                story.pop()
            doc.build(story)

            # Update status and preview
            self.status_label.config(text=f"Saved as: {os.path.basename(pdf_path)}")
            self.preview_pdf_page(pdf_path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()

    def pdf_to_text(self) -> None:
        """
        Extract all text from a PDF and save it to a plain-text file.

        This method prompts the user to select a PDF file and then uses
        pdfplumber to read each page's text content.  The extracted text
        (with page breaks preserved) is written to a `.txt` file next to
        the original PDF, suffixed with `_extracted.txt`.  If the user
        cancels the file selection or the PDF has no extractable text,
        the operation is aborted gracefully.
        """
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not pdf_path:
            return
        self.status_label.config(text=f"Selected: {os.path.basename(pdf_path)}")
        self.progress.start()
        self.root.update()
        try:
            output_path = os.path.splitext(pdf_path)[0] + "_extracted.txt"
            extracted_text = []
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    try:
                        text = page.extract_text() or ""
                        extracted_text.append(text)
                    except Exception:
                        continue
            if not any(extracted_text):
                messagebox.showwarning(
                    "No Text",
                    "No extractable text found in this PDF.",
                )
            else:
                full_text = "\f".join(extracted_text)
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(full_text)
                self.status_label.config(
                    text=f"Saved as: {os.path.basename(output_path)}"
                )
                preview = full_text[:500]
                messagebox.showinfo(
                    "Success",
                    f"Extracted text saved to {output_path}.\n\nPreview:\n\n{preview}",
                )
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()

    def clear_preview(self) -> None:
        """
        Clear the preview canvas and reset the status label.

        This method removes any rendered images or PDF previews from the
        preview area and updates the status label to indicate that no file
        is currently selected. It can be triggered via the "Clear Preview"
        action button in the sidebar.
        """
        # Remove any image from the canvas
        self.preview_canvas.delete("all")
        # Reset the stored Tk image reference to allow garbage collection
        self.tk_img = None
        # Reset status label
        self.status_label.config(text="No file selected")

    def pdf_to_excel(self) -> None:
        """
        Extract tables from a PDF into an Excel workbook.

        Uses pdfplumber to detect tables on each page and writes each set of
        tables to a separate sheet in the output workbook.  Header rows are
        preserved if possible, but table structure may vary from PDF to PDF.
        """
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not pdf_path:
            return

        self.status_label.config(text=f"Selected: {os.path.basename(pdf_path)}")
        self.progress.start()
        self.root.update()

        try:
            output_path = os.path.splitext(pdf_path)[0] + "_tables.xlsx"
            all_tables: List[pd.DataFrame] = []

            # Determine page count to handle large PDFs
            try:
                doc = fitz.open(pdf_path)
                total_pages = doc.page_count
            except Exception:
                total_pages = None
            finally:
                try:
                    doc.close()
                except Exception:
                    pass

            start_page = 1
            end_page = None
            MAX_PAGES_FOR_FULL_EXTRACTION = 100
            if total_pages is not None and total_pages > MAX_PAGES_FOR_FULL_EXTRACTION:
                # Prompt user for page range when the document is large
                msg = (
                    f"This PDF has {total_pages} pages. Extracting tables from all pages "
                    "may take a long time.\n"
                    "Enter the starting and ending pages for extraction."
                )
                first_page_in = simpledialog.askinteger(
                    "Start Page",
                    msg + "\n\nStart page:",
                    minvalue=1,
                    maxvalue=total_pages,
                )
                if first_page_in is None:
                    return
                last_page_in = simpledialog.askinteger(
                    "End Page",
                    f"End page (between {first_page_in} and {total_pages}):",
                    minvalue=first_page_in,
                    maxvalue=total_pages,
                )
                if last_page_in is None:
                    return
                start_page = first_page_in
                end_page = last_page_in
            elif total_pages is not None:
                end_page = total_pages

            with pdfplumber.open(pdf_path) as pdf:
                # Build range of page indices to iterate (0-based)
                if end_page is None:
                    page_indices = range(len(pdf.pages))
                else:
                    page_indices = range(start_page - 1, end_page)
                for idx in page_indices:
                    try:
                        page = pdf.pages[idx]
                    except IndexError:
                        break
                    # Attempt to extract tables using line detection
                    tables = page.extract_tables(
                        {
                            "vertical_strategy": "lines",
                            "horizontal_strategy": "lines",
                            "intersection_tolerance": 5,
                            "snap_tolerance": 3,
                            "join_tolerance": 3,
                            "edge_min_length": 3,
                        }
                    )
                    for table in tables:
                        if table:
                            df = pd.DataFrame(table)
                            all_tables.append(df)

            if all_tables:
                with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                    for i, df in enumerate(all_tables):
                        # Write each table to a separate sheet; no header is assumed
                        df.to_excel(
                            writer,
                            sheet_name=f"Table_{i + 1}",
                            index=False,
                            header=False,
                        )
                self.status_label.config(
                    text=f"Saved as: {os.path.basename(output_path)}"
                )
                messagebox.showinfo(
                    "Success", f"Excel saved as {output_path}"
                )
            else:
                messagebox.showinfo(
                    "No Data",
                    "No tables could be extracted from this PDF."
                )
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()

    def pdf_to_word(self) -> None:
        """Convert a PDF file (or a range of pages) to a Word (.docx) document."""
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not pdf_path:
            return

        self.status_label.config(text=f"Selected: {os.path.basename(pdf_path)}")
        self.progress.start()
        self.root.update()
        try:
            docx_path = os.path.splitext(pdf_path)[0] + ".docx"
            # Determine page count to handle very large PDFs
            try:
                doc = fitz.open(pdf_path)
                total_pages = doc.page_count
            except Exception:
                total_pages = None
            finally:
                try:
                    doc.close()
                except Exception:
                    pass

            start_idx = 0  # zero-based index for start page
            end_idx = None  # zero-based index for end page (inclusive)
            MAX_PAGES_FOR_FULL_CONVERT = 100
            if total_pages is not None and total_pages > MAX_PAGES_FOR_FULL_CONVERT:
                # Ask the user for a page range to convert
                msg = (
                    f"This PDF has {total_pages} pages. Converting all pages to DOCX "
                    "may take a long time.\n"
                    "Enter the starting and ending pages to convert."
                )
                first_page_in = simpledialog.askinteger(
                    "Start Page",
                    msg + "\n\nStart page:",
                    minvalue=1,
                    maxvalue=total_pages,
                )
                if first_page_in is None:
                    return
                last_page_in = simpledialog.askinteger(
                    "End Page",
                    f"End page (between {first_page_in} and {total_pages}):",
                    minvalue=first_page_in,
                    maxvalue=total_pages,
                )
                if last_page_in is None:
                    return
                start_idx = first_page_in - 1
                end_idx = last_page_in - 1
            # Convert specified page range (or all pages if end_idx is None)
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=start_idx, end=end_idx)
            cv.close()
            self.status_label.config(text=f"Saved as: {os.path.basename(docx_path)}")
            messagebox.showinfo("Success", f"Word document saved as {docx_path}")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()

    def compress_pdf(self) -> None:
        """
        Compress a PDF file and save the result.

        This method first tries a lossless optimization using PyMuPDF's
        ``garbage=4`` and ``deflate=True`` options.  If the resulting file
        isn't smaller than the original, a secondary lossy fallback will
        render pages at a reduced DPI and compress them as JPEG images
        before inserting them back into the document.  The fallback is
        generally effective for scanned PDFs but may reduce vector
        fidelity.
        """
        # Prompt the user to select a PDF to compress
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not pdf_path:
            return

        # Ask the user for JPEG quality for lossy fallback (1–100)
        quality = simpledialog.askinteger(
            "PDF Compression Quality",
            "Enter JPEG quality for lossy compression (1–100, default 80):",
            minvalue=1,
            maxvalue=100,
        )
        if quality is None:
            quality = 80

        # Ask the user for DPI for lossy fallback (72–300). Lower DPI yields smaller files.
        dpi = simpledialog.askinteger(
            "PDF Compression DPI",
            "Enter DPI for rasterizing pages (72–300, default 150):",
            minvalue=72,
            maxvalue=300,
        )
        if dpi is None:
            dpi = 150

        self.status_label.config(text=f"Selected: {os.path.basename(pdf_path)}")
        self.progress.start()
        self.root.update()
        try:
            input_size = os.path.getsize(pdf_path)
            output_path = os.path.splitext(pdf_path)[0] + "_compressed.pdf"
            # First attempt: lossless compression using garbage collection and deflate
            doc = fitz.open(pdf_path)
            try:
                doc.save(output_path, garbage=4, deflate=True)
            finally:
                doc.close()
            out_size = os.path.getsize(output_path)

            # Determine page count to decide whether to attempt lossy fallback
            try:
                doc_info = fitz.open(pdf_path)
                page_count = doc_info.page_count
            except Exception:
                page_count = None
            finally:
                try:
                    doc_info.close()
                except Exception:
                    pass

            # Threshold to avoid processing extremely large PDFs in the lossy fallback
            MAX_PAGES_FOR_LOSSY = 500
            # If the file wasn't reduced or the user specifically requested stronger compression
            # and the PDF isn't too large, perform lossy downsampling; otherwise skip.
            if out_size >= input_size and (page_count is None or page_count <= MAX_PAGES_FOR_LOSSY):
                import io
                from PIL import Image
                doc = fitz.open(pdf_path)
                for page in doc:
                    # Compute zoom factor based on DPI
                    zoom = dpi / 72.0
                    mat = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=mat)
                    # Convert to PIL image and compress as JPEG
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=quality)
                    page.clean_contents()
                    page.insert_image(page.rect, stream=buf.getvalue())
                doc.save(output_path)
                doc.close()
            elif out_size >= input_size:
                # Skip lossy fallback for very large PDFs
                messagebox.showinfo(
                    "Skipped Fallback",
                    f"The PDF has {page_count} pages. Lossy compression was skipped to avoid memory issues."
                )
            self.status_label.config(
                text=f"Saved as: {os.path.basename(output_path)}"
            )
            self.preview_pdf_page(output_path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()

    def image_to_pdf(self) -> None:
        """
        Convert one or more images into a single PDF file.

        To keep the resulting PDF small, each input image is opened with
        Pillow, converted to RGB and recompressed as a JPEG with
        moderate quality (default 85) before being embedded into the PDF
        via img2pdf.  This approach reduces file size while maintaining
        reasonable visual fidelity.  Additional image formats such as
        TIFF, TIF, ICO and WEBP are supported.
        """
        img_paths = filedialog.askopenfilenames(
            filetypes=[(
                "Image Files",
                "*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.tif;*.gif;*.ico;*.webp",
            )]
        )
        if not img_paths:
            return

        self.status_label.config(text=f"Selected {len(img_paths)} image(s)")
        self.progress.start()
        self.root.update()
        try:
            pdf_path = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=[("PDF File", "*.pdf")]
            )
            if not pdf_path:
                return
            import io
            images_data: List[io.BytesIO] = []
            for path in img_paths:
                try:
                    img = Image.open(path)
                    # Convert to RGB if necessary
                    if img.mode not in ("RGB", "L"):
                        img = img.convert("RGB")
                    # Compress image to JPEG
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=85)
                    buf.seek(0)
                    images_data.append(buf)
                except Exception:
                    # Skip unreadable images
                    continue
            if not images_data:
                messagebox.showwarning("No Images", "No valid images selected.")
                return
            # Write the compressed images into a single PDF
            with open(pdf_path, "wb") as f:
                f.write(img2pdf.convert(images_data))
            self.status_label.config(
                text=f"Saved as: {os.path.basename(pdf_path)}"
            )
            self.preview_pdf_page(pdf_path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()

    def merge_pdfs(self) -> None:
        """Merge multiple PDF files into a single PDF."""
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if not files:
            return

        merger = PdfMerger()
        try:
            for pdf in files:
                merger.append(pdf)
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf")
            if save_path:
                merger.write(save_path)
                messagebox.showinfo("Merged", f"Merged PDF saved as {save_path}")
                self.preview_pdf_page(save_path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))
        finally:
            merger.close()

    def split_pdf(self) -> None:
        """Split a range of pages from a PDF into a new PDF file."""
        file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not file:
            return

        reader = PdfReader(file)
        start = simpledialog.askinteger("Start Page", "Enter start page (1-based):")
        end = simpledialog.askinteger("End Page", "Enter end page:")
        if start is None or end is None or start < 1 or end < start:
            messagebox.showwarning("Invalid input", "Please enter valid start and end pages.")
            return
        try:
            writer = PdfWriter()
            for i in range(start - 1, min(end, len(reader.pages))):
                writer.add_page(reader.pages[i])
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf")
            if save_path:
                with open(save_path, "wb") as f:
                    writer.write(f)
                messagebox.showinfo("Split", f"Pages saved as {save_path}")
                self.preview_pdf_page(save_path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))

    def pdf_to_images(self) -> None:
        """
        Convert pages of a PDF into individual PNG images, with optional page range.

        This method detects the number of pages in the selected PDF. If the
        document has more than a threshold number of pages (default 100), it
        prompts the user to specify a start and end page to avoid loading
        thousands of pages into memory at once. The resulting images are
        saved to a user-selected folder.
        """
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not pdf_path:
            return

        # Determine the number of pages to warn about large PDFs
        try:
            doc = fitz.open(pdf_path)
            total_pages = doc.page_count
        except Exception:
            total_pages = None
        finally:
            try:
                doc.close()
            except Exception:
                pass

        # Prompt for page range if the PDF is large (more than 100 pages)
        first_page = 1
        last_page = None
        MAX_PAGES_FOR_FULL_CONVERSION = 100
        if total_pages is not None and total_pages > MAX_PAGES_FOR_FULL_CONVERSION:
            # Ask user for a start page and end page
            msg = (
                f"This PDF has {total_pages} pages. Converting all pages "
                "may take a long time and use a lot of memory.\n"
                "Enter the starting page (1-based) and ending page for conversion."
            )
            first_page_in = simpledialog.askinteger(
                "Start Page",
                msg + "\n\nStart page:",
                minvalue=1,
                maxvalue=total_pages,
            )
            if first_page_in is None:
                return
            last_page_in = simpledialog.askinteger(
                "End Page",
                f"End page (between {first_page_in} and {total_pages}):",
                minvalue=first_page_in,
                maxvalue=total_pages,
            )
            if last_page_in is None:
                return
            first_page = first_page_in
            last_page = last_page_in
        elif total_pages is not None:
            last_page = total_pages

        self.status_label.config(text=f"Selected: {os.path.basename(pdf_path)}")
        try:
            # Convert the specified page range to images
            pages = convert_from_path(
                pdf_path,
                dpi=300,
                first_page=first_page,
                last_page=last_page,
            )
            # Ask user for output directory
            folder = filedialog.askdirectory()
            if not folder:
                return
            for i, page in enumerate(pages, start=first_page):
                page.save(os.path.join(folder, f"page_{i}.png"), "PNG")
            messagebox.showinfo(
                "Done", f"Saved {len(pages)} image(s)."
            )
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))

    def batch_compress(self) -> None:
        """Compress multiple PDFs in one operation."""
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if not files:
            return
        # Start the progress bar in indeterminate mode for batch processing
        self.progress.start()
        self.root.update()
        try:
            for pdf_path in files:
                try:
                    input_size = os.path.getsize(pdf_path)
                    output = os.path.splitext(pdf_path)[0] + "_batch_compressed.pdf"
                    # Lossless attempt
                    doc = fitz.open(pdf_path)
                    doc.save(output, garbage=4, deflate=True)
                    doc.close()
                    out_size = os.path.getsize(output)
                    # Fallback to lossy downsample if necessary
                    if out_size >= input_size:
                        import io
                        from PIL import Image
                        doc = fitz.open(pdf_path)
                        for page in doc:
                            zoom = 150 / 72.0
                            mat = fitz.Matrix(zoom, zoom)
                            pix = page.get_pixmap(matrix=mat)
                            img = Image.frombytes(
                                "RGB", (pix.width, pix.height), pix.samples
                            )
                            buf = io.BytesIO()
                            img.save(buf, format="JPEG", quality=80)
                            page.clean_contents()
                            page.insert_image(page.rect, stream=buf.getvalue())
                        doc.save(output)
                        doc.close()
                except Exception:
                    # Skip file on error
                    continue
            messagebox.showinfo("Done", "Batch compression complete.")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()

    def preview_pdf_page(self, pdf_path: str) -> None:
        """
        Display a preview of the first page of the given PDF.

        Uses pdf2image to render the first page at a moderate DPI, then
        resizes it for display in the preview canvas.  The resulting image
        remains in memory as an attribute (tk_img) to prevent garbage collection.
        """
        try:
            images = convert_from_path(pdf_path, dpi=150, first_page=1, last_page=1)
            if images:
                img = images[0]
                # Determine available canvas size for scaling
                self.root.update_idletasks()
                canvas_width = self.preview_canvas.winfo_width()
                canvas_height = self.preview_canvas.winfo_height()
                # Provide sensible defaults in case geometry is not yet updated
                if canvas_width <= 10 or canvas_height <= 10:
                    max_width, max_height = 400, 500
                else:
                    max_width, max_height = canvas_width, canvas_height
                # Compute scaling ratio
                ratio = min(max_width / img.width, max_height / img.height)
                # Avoid division by zero
                if ratio <= 0 or ratio is None:
                    ratio = 1.0
                new_width = int(img.width * ratio)
                new_height = int(img.height * ratio)
                # Resize the image to fit within the canvas
                resized_img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                self.tk_img = ImageTk.PhotoImage(resized_img)
                self.preview_canvas.delete("all")
                # Center the image within the canvas
                x = max(0, (max_width - new_width) // 2)
                y = max(0, (max_height - new_height) // 2)
                self.preview_canvas.create_image(x, y, anchor="nw", image=self.tk_img)
        except Exception as e:  # noqa: BLE001
            # If preview fails, silently ignore; optionally log to console
            print("Preview error:", e)

    def merge_images_to_pdf(self) -> None:
        """
        Merge multiple images into a single PDF.

        This function recompresses each selected image into JPEG format
        (quality 85) before embedding them into a single PDF using
        ``img2pdf``.  Supporting many common image formats ensures broad
        compatibility.  The resulting PDF is previewed once written.
        """
        img_paths = filedialog.askopenfilenames(
            filetypes=[(
                "Image Files",
                "*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.tif;*.gif;*.ico;*.webp",
            )]
        )
        if not img_paths:
            return

        self.status_label.config(text=f"Selected {len(img_paths)} image(s)")
        self.progress.start()
        self.root.update()
        try:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=[("PDF File", "*.pdf")]
            )
            if not save_path:
                return
            import io
            compressed_images: List[io.BytesIO] = []
            for img_path in img_paths:
                try:
                    img = Image.open(img_path)
                    if img.mode not in ("RGB", "L"):
                        img = img.convert("RGB")
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=85)
                    buf.seek(0)
                    compressed_images.append(buf)
                except Exception:
                    continue
            if not compressed_images:
                messagebox.showwarning(
                    "No Images", "No valid images to merge."
                )
                return
            with open(save_path, "wb") as f:
                f.write(img2pdf.convert(compressed_images))
            self.status_label.config(
                text=f"Saved as: {os.path.basename(save_path)}"
            )
            self.preview_pdf_page(save_path)
            messagebox.showinfo(
                "Success", f"Merged PDF saved as {save_path}"
            )
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()

    def compress_images(self) -> None:
        """
        Compress one or more image files using preset quality and scaling values.

        This method prompts the user to select multiple images (PNG, JPEG, BMP,
        TIFF, GIF, ICO, WEBP, etc.) and then compresses each one using a
        predefined JPEG quality (75%) and scaling factor (100%, i.e., no
        resizing). Images are converted to RGB if necessary, then saved as
        JPEGs alongside the originals with a ``_compressed`` suffix.  By using
        fixed defaults, the function avoids confusing prompts for layman users.

        Unsupported or unreadable files are silently skipped, and a preview of
        the first compressed image is displayed in the preview pane.
        """
        # Prompt the user for the image files to compress
        img_paths = filedialog.askopenfilenames(
            filetypes=[(
                "Image Files",
                "*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.tif;*.gif;*.ico;*.webp",
            )]
        )
        if not img_paths:
            return

        # Use default compression settings suitable for layman users
        quality = 75  # JPEG quality percentage
        scale = 100   # Scaling percentage (100% = no scaling)

        self.status_label.config(
            text=f"Selected {len(img_paths)} image(s) for compression"
        )
        self.progress.start()
        self.root.update()
        try:
            compressed_count = 0
            for path in img_paths:
                try:
                    img = Image.open(path)
                    # Convert images with alpha or palette to RGB for JPEG compression
                    if img.mode not in ("RGB", "L"):
                        img = img.convert("RGB")
                    # Apply scaling if necessary
                    if scale != 100:
                        new_size = (
                            max(1, int(img.width * scale / 100)),
                            max(1, int(img.height * scale / 100)),
                        )
                        img = img.resize(new_size, Image.Resampling.LANCZOS)
                    base, _ = os.path.splitext(path)
                    out_path = f"{base}_compressed.jpg"
                    # Save as JPEG with specified quality
                    img.save(out_path, format="JPEG", quality=quality, optimize=True)
                    compressed_count += 1
                except Exception:
                    # Ignore unreadable/unsupported files
                    continue
            if compressed_count > 0:
                messagebox.showinfo(
                    "Success", f"Compressed {compressed_count} image(s)."
                )
                # Preview the first compressed image using dynamic scaling
                first_base, _ = os.path.splitext(img_paths[0])
                first_out = f"{first_base}_compressed.jpg"
                if os.path.exists(first_out):
                    try:
                        img_prev = Image.open(first_out)
                        # Determine available canvas size for scaling
                        self.root.update_idletasks()
                        canvas_width = self.preview_canvas.winfo_width()
                        canvas_height = self.preview_canvas.winfo_height()
                        # Provide sensible defaults if geometry isn't ready
                        if canvas_width <= 10 or canvas_height <= 10:
                            max_width, max_height = 400, 500
                        else:
                            max_width, max_height = canvas_width, canvas_height
                        # Compute scaling ratio to fit image within the canvas
                        ratio = min(max_width / img_prev.width, max_height / img_prev.height)
                        # Avoid invalid ratio
                        if ratio <= 0 or ratio is None:
                            ratio = 1.0
                        new_width = int(img_prev.width * ratio)
                        new_height = int(img_prev.height * ratio)
                        resized_img = img_prev.resize(
                            (new_width, new_height), Image.Resampling.LANCZOS
                        )
                        self.tk_img = ImageTk.PhotoImage(resized_img)
                        self.preview_canvas.delete("all")
                        # Center the image within the canvas
                        x = max(0, (max_width - new_width) // 2)
                        y = max(0, (max_height - new_height) // 2)
                        self.preview_canvas.create_image(x, y, anchor="nw", image=self.tk_img)
                    except Exception:
                        pass
                self.status_label.config(text="Image compression complete")
            else:
                messagebox.showwarning("No Images", "No valid images were compressed.")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToolkitApp(root)
    root.mainloop()