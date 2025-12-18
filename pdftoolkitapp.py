import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
from typing import Optional, List, Tuple

import pandas as pd

import fitz  # PyMuPDF, used for PDF compression
from pdf2docx import Converter
import img2pdf
from PIL import Image, ImageTk
from pdf2image import convert_from_path
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import pdfplumber

# Additional imports for extended functionality
# Word→PDF conversion may use docx2pdf if available
try:
    from docx2pdf import convert as docx2pdf_convert  # type: ignore
except Exception:
    docx2pdf_convert = None  # type: ignore

# Fallback COM automation for Word→PDF
try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None  # type: ignore

# Use pikepdf for encryption/decryption if available
try:
    import pikepdf  # type: ignore
except Exception:
    pikepdf = None  # type: ignore

from reportlab.pdfgen import canvas  # type: ignore
from reportlab.lib.pagesizes import letter as reportlab_letter  # type: ignore


class PDFToolkitApp:
    """A Tkinter GUI application providing various PDF and document utilities."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PDF Toolkit")
        self.root.geometry("900x600")
        self.root.minsize(600, 400)

        # Track whether dark mode is active
        self.is_dark_mode = False
        # Keep track of the most recent files dropped onto the app window
        self.last_dropped_paths: List[str] = []

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

        # Create sidebar for actions with a scrollable area
        self.sidebar = tk.Frame(root, bg=self.light_theme["sidebar_bg"], width=200)
        self.sidebar.pack(side="left", fill="y")

        # Split sidebar into a top section (for the theme toggle) and a scrollable section (for buttons)
        self.top_sidebar_frame = tk.Frame(self.sidebar, bg=self.light_theme["sidebar_bg"])
        self.top_sidebar_frame.pack(side="top", fill="x")

        # Toggle button for switching themes; placed in top frame so it stays visible
        self.toggle_btn = tk.Label(
            self.top_sidebar_frame,
            image=self.sun_image,
            cursor="hand2",
            bg="#f0f0f0",
        )
        self.toggle_btn.pack(pady=10, anchor="ne")
        self.toggle_btn.bind("<Button-1>", self.toggle_theme)
        # Create a simple tooltip for the toggle button
        self.toggle_btn_tooltip = tk.Label(
            self.top_sidebar_frame,
            text="Change Mode",
            bg="black",
            fg="white",
            font=("Segoe UI", 8),
            bd=1,
            relief="solid",
        )
        self.toggle_btn.bind(
            "<Enter>", lambda e: self.toggle_btn_tooltip.place(x=50, y=10)
        )
        self.toggle_btn.bind(
            "<Leave>", lambda e: self.toggle_btn_tooltip.place_forget()
        )

        # Canvas and scrollbar to make the button list scrollable. This allows the sidebar to
        # accommodate many buttons even in a small window.
        self.sidebar_canvas = tk.Canvas(
            self.sidebar,
            bg=self.light_theme["sidebar_bg"],
            highlightthickness=0,
        )
        self.sidebar_vscrollbar = tk.Scrollbar(
            self.sidebar, orient="vertical", command=self.sidebar_canvas.yview
        )
        self.sidebar_canvas.configure(yscrollcommand=self.sidebar_vscrollbar.set)
        self.sidebar_canvas.pack(side="left", fill="both", expand=True)
        self.sidebar_vscrollbar.pack(side="right", fill="y")

        # Frame inside the canvas to hold all the action buttons
        self.buttons_frame = tk.Frame(self.sidebar_canvas, bg=self.light_theme["sidebar_bg"])
        # Add the frame into the canvas window; tag it for resizing
        self.sidebar_canvas.create_window(
            (0, 0), window=self.buttons_frame, anchor="nw", tags=("buttons_frame",)
        )
        # Ensure the scrollregion is updated whenever the buttons frame changes size
        def _update_sidebar_scrollregion(event=None):
            self.sidebar_canvas.configure(scrollregion=self.sidebar_canvas.bbox("all"))
        self.buttons_frame.bind("<Configure>", _update_sidebar_scrollregion)
        # Optionally, make the buttons frame width track the canvas width
        def _resize_sidebar_frame(event):
            self.sidebar_canvas.itemconfig("buttons_frame", width=event.width)
        self.sidebar_canvas.bind("<Configure>", _resize_sidebar_frame)

        # Enable mouse wheel scrolling within the sidebar.  Without this binding,
        # the canvas will not respond to the scroll wheel.  Bindings are
        # platform-specific: Windows uses <MouseWheel> with delta values, while
        # Linux often uses Button-4/Button-5 for wheel events.  macOS also
        # supports <MouseWheel> with smaller delta values.
        def _on_sidebar_mousewheel(event):
            try:
                delta = event.delta
                # On Windows, delta is a multiple of 120; invert to get direction
                if os.name == "nt":
                    self.sidebar_canvas.yview_scroll(int(-1 * (delta / 120)), "units")
                else:
                    # On other platforms, delta may already be small; just use it
                    if delta != 0:
                        self.sidebar_canvas.yview_scroll(int(-1 * delta), "units")
            except Exception:
                pass

        # Bind the wheel events globally on the root so that the scroll wheel
        # always controls the sidebar when it is present.  This ensures
        # compatibility with physical mouse wheels on all platforms.
        self.root.bind_all("<MouseWheel>", _on_sidebar_mousewheel)
        # Bind Linux-specific scroll events globally
        self.root.bind_all("<Button-4>", lambda e: self.sidebar_canvas.yview_scroll(-1, "units"))
        self.root.bind_all("<Button-5>", lambda e: self.sidebar_canvas.yview_scroll(1, "units"))

        # Create main content area for previews and status
        self.content = tk.Frame(root, bg=self.light_theme["content_bg"])
        self.content.pack(side="right", expand=True, fill="both")

        # Initialize list of buttons
        self.buttons = []
        # Populate the action buttons inside the buttons_frame
        actions: List[Tuple[str, callable]] = [
            ("Excel to PDF", self.excel_to_pdf),
            ("PDF to Excel", self.pdf_to_excel),
            ("PDF to Word", self.pdf_to_word),
            ("Word to PDF", self.word_to_pdf),
            ("Compress PDF", self.compress_pdf),
            ("Compress Images", self.compress_images),
            ("Image to PDF", self.image_to_pdf),
            ("Merge PDFs", self.merge_pdfs),
            ("Split PDF", self.split_pdf),
            ("Rotate Pages", self.rotate_pdf_pages),
            ("PDF to Images", self.pdf_to_images),
            ("Batch Compress PDFs", self.batch_compress),
            ("Encrypt PDF", self.encrypt_pdf),
            ("Decrypt PDF", self.decrypt_pdf),
            ("Merge Images to PDF", self.merge_images_to_pdf),
            ("Add Watermark", self.add_watermark),
            # New functionality: extract all text from a PDF into a plain-text file
            ("PDF to Text", self.pdf_to_text),
            ("About", self.show_about),
        ]
        for text, command in actions:
            btn = tk.Button(
                self.buttons_frame,
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
            btn.pack(fill="x", pady=5, padx=10)
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

        # Container frame and scrollable canvas for previews. When the window
        # is resized to a small size, scrollbars allow the user to pan around
        # a large preview. The canvas scrollregion is updated when images
        # or PDF pages are drawn.
        self.preview_container = tk.Frame(self.content, bg=self.light_theme["canvas_bg"])
        self.preview_container.pack(pady=10, expand=True, fill="both")
        # Canvas where images/PDF pages will be drawn
        self.preview_canvas = tk.Canvas(self.preview_container, bg=self.light_theme["canvas_bg"])
        # Vertical and horizontal scrollbars tied to the preview canvas
        self.preview_vscrollbar = tk.Scrollbar(self.preview_container, orient="vertical",
                                               command=self.preview_canvas.yview)
        self.preview_hscrollbar = tk.Scrollbar(self.preview_container, orient="horizontal",
                                               command=self.preview_canvas.xview)
        self.preview_canvas.configure(xscrollcommand=self.preview_hscrollbar.set,
                                      yscrollcommand=self.preview_vscrollbar.set)
        # Grid layout for canvas and scrollbars
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        self.preview_vscrollbar.grid(row=0, column=1, sticky="ns")
        self.preview_hscrollbar.grid(row=1, column=0, sticky="ew")
        self.preview_container.rowconfigure(0, weight=1)
        self.preview_container.columnconfigure(0, weight=1)

        # Create a bottom frame to hold the progress bar and clear-preview button
        # This positions the clear button next to the progress bar.
        self.bottom_frame = tk.Frame(
            self.content,
            bg=self.light_theme["content_bg"],
        )
        self.bottom_frame.columnconfigure(0, weight=1)
        # Progress bar to indicate long running tasks; default to indeterminate
        self.progress = ttk.Progressbar(
            self.bottom_frame,
            orient="horizontal",
            mode="indeterminate",
            length=200,
        )
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 10))
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
        self.clear_btn.grid(row=0, column=1, sticky="e")
        # Pack the bottom frame
        self.bottom_frame.pack(pady=5, padx=10, fill="x")

        # Enable drag-and-drop on Windows for quick file selection/preview
        self.setup_drag_and_drop()

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
        # Update nested sidebar frames and canvas backgrounds
        if hasattr(self, "top_sidebar_frame"):
            self.top_sidebar_frame.config(bg=theme["sidebar_bg"])
        if hasattr(self, "sidebar_canvas"):
            self.sidebar_canvas.config(bg=theme["sidebar_bg"])
        if hasattr(self, "buttons_frame"):
            self.buttons_frame.config(bg=theme["sidebar_bg"])
        self.content.config(bg=theme["content_bg"])

        # Header and status labels
        self.header_label.config(bg=theme["content_bg"], fg=theme["header_fg"])
        self.status_label.config(bg=theme["content_bg"], fg=theme["label_fg"])

        # Preview canvas and container scrollbars
        self.preview_canvas.config(bg=theme["canvas_bg"])
        # Ensure container matches canvas background
        if hasattr(self, "preview_container"):
            self.preview_container.config(bg=theme["canvas_bg"])

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
    # Drag-and-drop support (Windows)
    # ------------------------------------------------------------------
    def setup_drag_and_drop(self) -> None:
        """
        Register OS-level drag-and-drop for the main window (Windows only).

        This hooks the window procedure so dropped files arrive in
        ``handle_dropped_files``. It is intentionally no-op on non-Windows
        systems so the rest of the app continues to work unchanged.
        """
        if os.name != "nt":
            return
        try:
            import ctypes
            from ctypes import wintypes
        except Exception as exc:  # noqa: BLE001
            print("Drag-and-drop unavailable:", exc)
            return

        # Type aliases for Win32 pointer-sized values
        LONG_PTR = ctypes.c_ssize_t
        WPARAM = ctypes.c_size_t
        LPARAM = ctypes.c_ssize_t
        LRESULT = getattr(wintypes, "LRESULT", ctypes.c_ssize_t)

        user32 = ctypes.windll.user32
        shell32 = ctypes.windll.shell32

        WM_DROPFILES = 0x0233
        GWL_WNDPROC = -4
        DragAcceptFiles = shell32.DragAcceptFiles
        DragQueryFileW = shell32.DragQueryFileW
        DragFinish = shell32.DragFinish
        CallWindowProcW = user32.CallWindowProcW
        SetWindowLongPtrW = user32.SetWindowLongPtrW
        DefWindowProcW = user32.DefWindowProcW
        DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
        DragAcceptFiles.restype = None
        DragQueryFileW.argtypes = [
            wintypes.HANDLE,
            ctypes.c_uint,
            ctypes.c_wchar_p,
            ctypes.c_uint,
        ]
        DragQueryFileW.restype = ctypes.c_uint
        DragFinish.argtypes = [wintypes.HANDLE]
        DragFinish.restype = None
        # Ensure correct signatures to avoid pointer truncation
        SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, LONG_PTR]
        SetWindowLongPtrW.restype = LONG_PTR
        CallWindowProcW.argtypes = [
            LONG_PTR,
            wintypes.HWND,
            wintypes.UINT,
            WPARAM,
            LPARAM,
        ]
        CallWindowProcW.restype = LONG_PTR
        DefWindowProcW.argtypes = [
            wintypes.HWND,
            wintypes.UINT,
            WPARAM,
            LPARAM,
        ]
        DefWindowProcW.restype = LRESULT

        hwnd = self.root.winfo_id()

        WNDPROC = ctypes.WINFUNCTYPE(
            LRESULT,
            wintypes.HWND,
            wintypes.UINT,
            WPARAM,
            LPARAM,
        )

        def wnd_proc(hWnd, msg, wParam, lParam):
            if msg == WM_DROPFILES:
                try:
                    file_count = DragQueryFileW(wParam, 0xFFFFFFFF, None, 0)
                    paths: List[str] = []
                    for i in range(file_count):
                        length = DragQueryFileW(wParam, i, None, 0) + 1
                        buffer = ctypes.create_unicode_buffer(length)
                        DragQueryFileW(wParam, i, buffer, length)
                        paths.append(buffer.value)
                    DragFinish(wParam)
                    # Handle on the Tk event loop to avoid thread issues
                    self.root.after(0, lambda p=paths: self.handle_dropped_files(p))
                except Exception as exc_inner:  # noqa: BLE001
                    print("Drag/drop error:", exc_inner)
                return 0
            # If no previous window proc was stored, fall back to default
            if not self._old_wnd_proc:
                return DefWindowProcW(hWnd, msg, wParam, lParam)
            try:
                return CallWindowProcW(self._old_wnd_proc, hWnd, msg, wParam, lParam)
            except Exception as exc_call:  # noqa: BLE001
                print("CallWindowProcW error:", exc_call)
                return DefWindowProcW(hWnd, msg, wParam, lParam)

        try:
            self._new_wnd_proc = WNDPROC(wnd_proc)
            new_proc_ptr = ctypes.cast(self._new_wnd_proc, ctypes.c_void_p).value
            self._old_wnd_proc = int(SetWindowLongPtrW(
                hwnd,
                GWL_WNDPROC,
                LONG_PTR(new_proc_ptr),
            ))
            DragAcceptFiles(hwnd, True)
        except Exception as exc:  # noqa: BLE001
            print("Unable to enable drag-and-drop:", exc)

    def handle_dropped_files(self, paths: List[str]) -> None:
        """
        Handle files dropped onto the application window.

        The first dropped file is previewed (PDF or image). All dropped
        paths are retained in ``self.last_dropped_paths`` for quick reuse.
        """
        if not paths:
            return
        files = [p for p in paths if os.path.isfile(p)]
        if not files:
            return
        self.last_dropped_paths = files
        primary = files[0]
        name = os.path.basename(primary)
        self.status_label.config(text=f"Dropped: {name}")
        ext = os.path.splitext(primary)[1].lower()
        if ext == ".pdf":
            self.preview_pdf_page(primary)
        elif ext in (
            ".png",
            ".jpg",
            ".jpeg",
            ".bmp",
            ".tiff",
            ".tif",
            ".gif",
            ".ico",
            ".webp",
        ):
            self.preview_image_file(primary)
        else:
            messagebox.showinfo(
                "Drag & Drop",
                f"Received {name}. Choose an action from the sidebar to process it.",
            )

    def get_dropped_files(
        self, allowed_exts: Tuple[str, ...], multiple: bool = False
    ) -> Optional[List[str]]:
        """
        Return dropped file(s) if available and matching the expected extensions.

        When ``multiple`` is False, a single matching file (the first) is
        returned; otherwise, all matching dropped files are returned. If no
        suitable dropped files exist, None is returned and callers can fall
        back to a file dialog.
        """
        if not self.last_dropped_paths:
            return None
        filtered = [
            p
            for p in self.last_dropped_paths
            if os.path.splitext(p)[1].lower() in allowed_exts
            and os.path.isfile(p)
        ]
        if not filtered:
            return None
        return filtered if multiple else [filtered[0]]

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
        excel_path = None
        dropped = self.get_dropped_files((".xlsx", ".xls"), multiple=False)
        if dropped:
            excel_path = dropped[0]
        else:
            excel_path = filedialog.askopenfilename(
                filetypes=[("Excel Files", "*.xlsx *.xls")]
            )
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
            story: List = []
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
            from reportlab.platypus import PageBreak
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
        pdf_path = None
        dropped = self.get_dropped_files((".pdf",), multiple=False)
        if dropped:
            pdf_path = dropped[0]
        else:
            pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not pdf_path:
            return
        self.status_label.config(text=f"Selected: {os.path.basename(pdf_path)}")
        self.progress.start()
        self.root.update()
        try:
            output_path = os.path.splitext(pdf_path)[0] + "_extracted.txt"
            extracted_text: List[str] = []
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
        pdf_path = None
        dropped = self.get_dropped_files((".pdf",), multiple=False)
        if dropped:
            pdf_path = dropped[0]
        else:
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
        """
        Convert a PDF file (or a range of pages) to a Word (.docx) document.
        """
        pdf_path = None
        dropped = self.get_dropped_files((".pdf",), multiple=False)
        if dropped:
            pdf_path = dropped[0]
        else:
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
            if end_idx is None:
                cv.convert(docx_path, start=start_idx)
            else:
                cv.convert(docx_path, start=start_idx, end=end_idx)
            cv.close()
            self.status_label.config(text=f"Saved as: {os.path.basename(docx_path)}")
            messagebox.showinfo("Success", f"Word document saved as {docx_path}")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()

    def word_to_pdf(self) -> None:
        """
        Convert a Word (.docx/.doc) document to PDF. Requires MS Word on Windows.
        Attempts to use docx2pdf if available, otherwise falls back to COM automation.
        """
        doc_path = None
        dropped = self.get_dropped_files((".docx", ".doc"), multiple=False)
        if dropped:
            doc_path = dropped[0]
        else:
            doc_path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx *.doc")])
        if not doc_path:
            return

        self.status_label.config(text=f"Selected: {os.path.basename(doc_path)}")
        self.progress.start()
        self.root.update()
        try:
            pdf_path = os.path.splitext(doc_path)[0] + ".pdf"
            # Prefer docx2pdf if available
            if docx2pdf_convert is not None:
                try:
                    docx2pdf_convert(doc_path, pdf_path)
                except Exception:
                    # fallback to COM
                    if win32com:
                        word = win32com.client.Dispatch("Word.Application")  # type: ignore
                        word.Visible = False
                        doc = word.Documents.Open(doc_path)  # type: ignore
                        # 17 = wdFormatPDF
                        doc.SaveAs(pdf_path, FileFormat=17)  # type: ignore
                        doc.Close(False)  # type: ignore
                        word.Quit()
                    else:
                        raise
            else:
                # Use COM if docx2pdf is not available
                if win32com:
                    word = win32com.client.Dispatch("Word.Application")  # type: ignore
                    word.Visible = False
                    doc = word.Documents.Open(doc_path)  # type: ignore
                    doc.SaveAs(pdf_path, FileFormat=17)  # type: ignore
                    doc.Close(False)  # type: ignore
                    word.Quit()
                else:
                    raise RuntimeError("Neither docx2pdf nor win32com are available.")
            self.status_label.config(text=f"Saved as: {os.path.basename(pdf_path)}")
            self.preview_pdf_page(pdf_path)
            messagebox.showinfo("Success", f"PDF saved as:\n{pdf_path}")
        except Exception as e:
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
        # Use a dropped file if present; otherwise prompt the user
        pdf_path = None
        dropped = self.get_dropped_files((".pdf",), multiple=False)
        if dropped:
            pdf_path = dropped[0]
        else:
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
                    pix = page.get_pixmap(matrix=mat)  # type: ignore[attr-defined]
                    # Convert to PIL image and compress as JPEG
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=quality)
                    page.clean_contents()
                    page.insert_image(page.rect, stream=buf.getvalue())  # type: ignore[attr-defined]
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
        allowed_exts = (
            ".png",
            ".jpg",
            ".jpeg",
            ".bmp",
            ".tiff",
            ".tif",
            ".gif",
            ".ico",
            ".webp",
        )
        img_paths = self.get_dropped_files(allowed_exts, multiple=True)
        if not img_paths:
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
            merged_bytes = img2pdf.convert(images_data)
            with open(pdf_path, "wb") as f:
                f.write(merged_bytes)
            self.status_label.config(
                text=f"Saved as: {os.path.basename(pdf_path)}"
            )
            self.preview_pdf_page(pdf_path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()

    def preview_image_file(self, img_path: str) -> None:
        """
        Preview an image file in the main canvas, scaling to fit.

        The image is resized to the current canvas dimensions while
        preserving aspect ratio, and the Tk reference is kept alive
        on the instance to avoid garbage collection.
        """
        try:
            img_prev = Image.open(img_path)
            self.root.update_idletasks()
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            if canvas_width <= 10 or canvas_height <= 10:
                max_width, max_height = 400, 500
            else:
                max_width, max_height = canvas_width, canvas_height
            ratio = min(max_width / img_prev.width, max_height / img_prev.height)
            if ratio <= 0 or ratio is None:
                ratio = 1.0
            new_width = int(img_prev.width * ratio)
            new_height = int(img_prev.height * ratio)
            resized_img = img_prev.resize(
                (new_width, new_height), Image.Resampling.LANCZOS
            )
            self.tk_img = ImageTk.PhotoImage(resized_img)
            self.preview_canvas.delete("all")
            x = max(0, (max_width - new_width) // 2)
            y = max(0, (max_height - new_height) // 2)
            self.preview_canvas.create_image(x, y, anchor="nw", image=self.tk_img)
            # Update scrollregion and reset view for the preview canvas
            self.preview_canvas.config(scrollregion=self.preview_canvas.bbox("all"))
            self.preview_canvas.xview_moveto(0)
            self.preview_canvas.yview_moveto(0)
        except Exception as e:  # noqa: BLE001
            print("Preview error:", e)

    def merge_pdfs(self) -> None:
        """Merge multiple PDF files into a single PDF."""
        files = self.get_dropped_files((".pdf",), multiple=True)
        if not files:
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
        file = None
        dropped = self.get_dropped_files((".pdf",), multiple=False)
        if dropped:
            file = dropped[0]
        else:
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

    def rotate_pdf_pages(self) -> None:
        """
        Rotate pages of a PDF by a specified angle.

        Prompts the user to select a PDF, enter the rotation angle (90/180/270),
        and optionally specify a page range. Saves a new PDF with rotated pages.
        """
        pdf_path = None
        dropped = self.get_dropped_files((".pdf",), multiple=False)
        if dropped:
            pdf_path = dropped[0]
        else:
            pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not pdf_path:
            return

        try:
            angle = simpledialog.askinteger(
                "Rotation Angle",
                "Enter rotation angle (90, 180, 270):",
                minvalue=90,
                maxvalue=270,
            )
            if angle not in (90, 180, 270):
                messagebox.showwarning("Invalid Angle", "Please enter 90, 180, or 270.")
                return
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)
            # Ask for page range
            first_page_in = simpledialog.askinteger(
                "Start Page",
                f"Enter start page (1 to {total_pages}) (default 1):",
                minvalue=1,
                maxvalue=total_pages,
            )
            if first_page_in is None:
                first_page_in = 1
            last_page_in = simpledialog.askinteger(
                "End Page",
                f"Enter end page (between {first_page_in} and {total_pages}) (default {total_pages}):",
                minvalue=first_page_in,
                maxvalue=total_pages,
            )
            if last_page_in is None:
                last_page_in = total_pages
            writer = PdfWriter()
            for i, page in enumerate(reader.pages, start=1):
                if first_page_in <= i <= last_page_in:
                    rotated = page.rotate(angle)
                    writer.add_page(rotated)
                else:
                    writer.add_page(page)
            save_path = os.path.splitext(pdf_path)[0] + f"_rotated{angle}.pdf"
            save_path = filedialog.asksaveasfilename(
                initialfile=os.path.basename(save_path),
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
            )
            if save_path:
                with open(save_path, "wb") as f:
                    writer.write(f)
                messagebox.showinfo("Success", f"Rotated PDF saved as:\n{save_path}")
                self.preview_pdf_page(save_path)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def encrypt_pdf(self) -> None:
        """
        Encrypt a PDF with a user-supplied password.
        Uses PyPDF2 or pikepdf if available.
        """
        pdf_path = None
        dropped = self.get_dropped_files((".pdf",), multiple=False)
        if dropped:
            pdf_path = dropped[0]
        else:
            pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not pdf_path:
            return
        password = simpledialog.askstring(
            "Set Password", "Enter a password to encrypt the PDF:", show="*"
        )
        if not password:
            return
        self.progress.start()
        self.root.update()
        try:
            # Attempt with pikepdf if available (handles preservation of metadata)
            if pikepdf is not None:
                with pikepdf.Pdf.open(pdf_path) as pdf:
                    output_path = os.path.splitext(pdf_path)[0] + "_encrypted.pdf"
                    pdf.save(
                        output_path,
                        encryption=pikepdf.Encryption(owner=password, user=password),
                    )
            else:
                reader = PdfReader(pdf_path)
                writer = PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)
                writer.encrypt(user_pwd=password, owner_pwd=password)
                output_path = os.path.splitext(pdf_path)[0] + "_encrypted.pdf"
                with open(output_path, "wb") as f:
                    writer.write(f)
            self.status_label.config(text=f"Encrypted: {os.path.basename(output_path)}")
            messagebox.showinfo("Success", f"Encrypted PDF saved as:\n{output_path}")
            self.preview_pdf_page(output_path)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()

    def decrypt_pdf(self) -> None:
        """
        Remove password protection from a PDF. Prompts for the password.
        """
        pdf_path = None
        dropped = self.get_dropped_files((".pdf",), multiple=False)
        if dropped:
            pdf_path = dropped[0]
        else:
            pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not pdf_path:
            return
        password = simpledialog.askstring(
            "PDF Password", "Enter the current password:", show="*"
        )
        if password is None:
            return
        self.progress.start()
        self.root.update()
        try:
            # Try with pikepdf first
            if pikepdf is not None:
                try:
                    with pikepdf.Pdf.open(pdf_path, password=password) as pdf:
                        output_path = os.path.splitext(pdf_path)[0] + "_decrypted.pdf"
                        pdf.save(output_path)
                except pikepdf.PasswordError:
                    messagebox.showerror("Error", "Incorrect password or unable to decrypt.")
                    return
            else:
                try:
                    reader = PdfReader(pdf_path)
                    if reader.is_encrypted:
                        reader.decrypt(password)
                    writer = PdfWriter()
                    for page in reader.pages:
                        writer.add_page(page)
                    output_path = os.path.splitext(pdf_path)[0] + "_decrypted.pdf"
                    with open(output_path, "wb") as f:
                        writer.write(f)
                except Exception:
                    messagebox.showerror("Error", "Incorrect password or unable to decrypt.")
                    return
            self.status_label.config(text=f"Decrypted: {os.path.basename(output_path)}")
            messagebox.showinfo("Success", f"Decrypted PDF saved as:\n{output_path}")
            self.preview_pdf_page(output_path)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()

    def add_watermark(self) -> None:
        """
        Add a text watermark to each page of a PDF.

        Prompts the user for a PDF, watermark text, and outputs a new PDF with the
        watermark applied on each page.
        """
        pdf_path = None
        dropped = self.get_dropped_files((".pdf",), multiple=False)
        if dropped:
            pdf_path = dropped[0]
        else:
            pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not pdf_path:
            return
        watermark_text = simpledialog.askstring("Watermark Text", "Enter the watermark text:")
        if not watermark_text:
            return
        self.progress.start()
        self.root.update()
        try:
            # Create a temporary watermark PDF in memory
            import io
            packet = io.BytesIO()
            c = canvas.Canvas(packet, pagesize=reportlab_letter)
            width, height = reportlab_letter
            c.setFont("Helvetica", 40)
            c.setFillColorRGB(0.6, 0.6, 0.6, alpha=0.3)  # semi-transparent grey
            c.saveState()
            # rotate text at an angle and center
            c.translate(width / 2, height / 2)
            c.rotate(45)
            c.drawCentredString(0, 0, watermark_text)
            c.restoreState()
            c.showPage()
            c.save()
            packet.seek(0)
            watermark_reader = PdfReader(packet)
            watermark_page = watermark_reader.pages[0]
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            for page in reader.pages:
                page.merge_page(watermark_page)
                writer.add_page(page)
            save_path = os.path.splitext(pdf_path)[0] + "_watermarked.pdf"
            save_path = filedialog.asksaveasfilename(
                initialfile=os.path.basename(save_path),
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
            )
            if save_path:
                with open(save_path, "wb") as f:
                    writer.write(f)
                messagebox.showinfo("Success", f"Watermarked PDF saved as:\n{save_path}")
                self.preview_pdf_page(save_path)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()

    def show_about(self) -> None:
        """
        Display an About dialog showing version and credits.
        """
        about_text = (
            "PDF Toolkit\n"
            "Version 1.0.0\n\n"
            "This application provides a suite of tools for converting,"
            " compressing, securing and manipulating PDF and image files.\n\n"
            "Developed using Python, Tkinter and various open source libraries."
        )
        messagebox.showinfo("About PDF Toolkit", about_text)

    def pdf_to_images(self) -> None:
        """
        Convert pages of a PDF into individual PNG images, with optional page range.

        This method detects the number of pages in the selected PDF. If the
        document has more than a threshold number of pages (default 100), it
        prompts the user to specify a start and end page to avoid loading
        thousands of pages into memory at once. The resulting images are
        saved to a user-selected folder.
        """
        pdf_path = None
        dropped = self.get_dropped_files((".pdf",), multiple=False)
        if dropped:
            pdf_path = dropped[0]
        else:
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
            convert_kwargs = {
                "pdf_path": pdf_path,
                "dpi": 300,
                "first_page": first_page,
            }
            if last_page is not None:
                convert_kwargs["last_page"] = last_page
            pages = convert_from_path(**convert_kwargs)
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
        files = self.get_dropped_files((".pdf",), multiple=True)
        if not files:
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
                            pix = page.get_pixmap(matrix=mat)  # type: ignore[attr-defined]
                            img = Image.frombytes(
                                "RGB", (pix.width, pix.height), pix.samples
                            )
                            buf = io.BytesIO()
                            img.save(buf, format="JPEG", quality=80)
                            page.clean_contents()
                            page.insert_image(page.rect, stream=buf.getvalue())  # type: ignore[attr-defined]
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
                # Update scrollregion so that scrollbars reflect the drawn content
                self.preview_canvas.config(scrollregion=self.preview_canvas.bbox("all"))
                # Reset the scroll position to the top-left
                self.preview_canvas.xview_moveto(0)
                self.preview_canvas.yview_moveto(0)
        except Exception as e:  # noqa: BLE001
            self.status_label.config(text=f"Preview unavailable: {e}")

    def merge_images_to_pdf(self) -> None:
        """
        Merge multiple images into a single PDF.

        This function recompresses each selected image into JPEG format
        (quality 85) before embedding them into a single PDF using
        ``img2pdf``.  Supporting many common image formats ensures broad
        compatibility.  The resulting PDF is previewed once written.
        """
        allowed_exts = (
            ".png",
            ".jpg",
            ".jpeg",
            ".bmp",
            ".tiff",
            ".tif",
            ".gif",
            ".ico",
            ".webp",
        )
        img_paths = self.get_dropped_files(allowed_exts, multiple=True)
        if not img_paths:
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
        JPEGs alongside the originals with a ``_compressed`` suffix.

        Unsupported or unreadable files are silently skipped, and a preview of
        the first compressed image is displayed in the preview pane.
        """
        # Prompt the user for the image files to compress (or use dropped files)
        allowed_exts = (
            ".png",
            ".jpg",
            ".jpeg",
            ".bmp",
            ".tiff",
            ".tif",
            ".gif",
            ".ico",
            ".webp",
        )
        img_paths = self.get_dropped_files(allowed_exts, multiple=True)
        if not img_paths:
            img_paths = filedialog.askopenfilenames(
                filetypes=[(
                    "Image Files",
                    "*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.tif;*.gif;*.ico;*.webp",
                )]
            )
        if not img_paths:
            return

        # Use default compression settings
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
                        # Update scrollregion and reset the scroll position
                        self.preview_canvas.config(scrollregion=self.preview_canvas.bbox("all"))
                        self.preview_canvas.xview_moveto(0)
                        self.preview_canvas.yview_moveto(0)
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