import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font
from tkinter.scrolledtext import ScrolledText
import os
from pathlib import Path
import platform
import time
import subprocess
import io
import struct

# Gestione formati estesi con PyMuPDF e python-docx
import pymupdf
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Presentazioni PPTX (pura Python, nessuna dipendenza da LibreOffice)
try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False


import re
import html as html_lib
import zipfile
import xml.etree.ElementTree as ET

# Prova a importare docx2pdf per la conversione nativa dei file Word (opzionale)
try:
    from docx2pdf import convert as convert_docx_to_pdf
    HAS_DOCX2PDF = True
except ImportError:
    HAS_DOCX2PDF = False

# Libreria per esportare in PDF
from reportlab.lib.pagesizes import A4, A3, A2
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# OCR e immagini
from PIL import Image, ImageDraw, ImageFont, ImageTk


class ToolTip:
    """Tooltip generico per i controlli dell'interfaccia."""
    def __init__(self, widget, text, delay=550):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip_window = None
        self.after_id = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, event=None):
        self._cancel()
        self.after_id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self.after_id:
            try:
                self.widget.after_cancel(self.after_id)
            except Exception:
                pass
            self.after_id = None

    def _show(self):
        if self.tip_window or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 6
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        except Exception:
            return
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        try:
            tw.wm_attributes("-topmost", True)
        except Exception:
            pass
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify=tk.LEFT, background="#ffffe0",
                          relief="solid", borderwidth=1, font=("Arial", 8))
        label.pack(ipadx=4, ipady=2)

    def _hide(self, event=None):
        self._cancel()
        if self.tip_window:
            try:
                self.tip_window.destroy()
            except Exception:
                pass
            self.tip_window = None


class AdvancedTextEditor:
    _FONT_SCALE = 0.88  # scala la grandezza originale del PDF con la nostra A4

    def __init__(self, root):
        self.root = root
        self.root.title("Simplewright - Multi-Page Editor")
        self.root.geometry("1400x850")

        self.current_file = None
        self.is_dark = False

        self.active_styles = {
            "bold": False,
            "italic": False,
            "underline": False,
            "alignment": "left"
        }

        self.current_color = "#000000"
        self.recent_colors = ["#ffffff", "#ffffff", "#ffffff", "#ffffff"]
        self.selected_slot_idx = 0
        self.current_bg_color = "#ffffff"

        self.current_popup = None
        self.last_popup_closed_time = 0
        self.active_popup_type = None

        self.base_page_sizes = {
            "A4": (595, 842),
            "A3": (842, 1191),
            "A2": (1191, 1684)
        }
        self.current_format = "A4"
        self.current_orientation = "Verticale"
        self.zoom_factor = 1.0
        self._zoom_job = None
        self._pdf_center_job = None

        self.pages_data = {1: ""}
        self.active_page_id = 1
        self.page_display_names = {}          
        self.page_text_boxes = {1: []}        
        self.page_text_format = {1: []}        
        self.page_text_box_widgets = {1: []}  
        self.page_images = {1: []}
        self.page_image_widgets = {1: []}
        self.active_pdf_image = None
        self._img_after_jobs = {}
        self._img_refs = {}
        self._img_resize_jobs = []
        self.text_page_frames = {}
        self.text_page_editors = {}

        self.embedded_images_cache = {}
        self.font_tag_defs = {}   
        self._font_tag_map = {}   
        self._font_tag_counter = 0

        self.pdf_mode = False
        self.pdf_pages_info = {}
        # DPI di riferimento per la vista PDF: ogni pagina viene ricondotta al
        # foglio A4 standard (595x842 px a zoom 1.0), quindi la scala base e'
        # 72.0 (1 px = 1 pt). Usato come fallback per pagine senza render_dpi.
        self.pdf_render_dpi = 72
        self.pdf_add_mode = False
        self.pdf_photo_refs = {}
        self.pdf_entry_widgets = {}
        self.active_pdf_box = None
        self.pdf_page_offsets = {}

        self.config_path = self.get_config_path()
        self.load_config()

        self.style = ttk.Style()
        self.theme_name = "clam"
        self.style.theme_use(self.theme_name)

        self.pixel_img = tk.PhotoImage(width=1, height=1)
        self.bg_button_img = self.create_highlighter_image()

        self.create_menu()
        self.create_toolbar()
        self.create_workspace()

        self.setup_global_shortcuts()
        self.apply_theme()

        self.update_pages_sidebar()
        self.update_page_layout()

    def _add_tooltip(self, widget, text):
        ToolTip(widget, text)

    def get_config_path(self):
        system_os = platform.system().lower()
        home_dir = os.path.expanduser("~")
        if "windows" in system_os:
            appdata_local_low = os.path.join(home_dir, "AppData", "LocalLow", "universalwriter")
            os.makedirs(appdata_local_low, exist_ok=True)
            return os.path.join(appdata_local_low, "config.txt")
        else:
            linux_config_dir = os.path.join(home_dir, ".universalwriter")
            os.makedirs(linux_config_dir, exist_ok=True)
            return os.path.join(linux_config_dir, "config.txt")

    def load_config(self):
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    for line in f:
                        if line.startswith("theme="):
                            theme_val = line.split("=")[1].strip().lower()
                            self.is_dark = (theme_val == "dark")
            except Exception:
                self.is_dark = False
        else:
            self.save_config()

    def save_config(self):
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                theme_str = "dark" if self.is_dark else "light"
                f.write(f"theme={theme_str}\n")
        except Exception:
            pass

    def create_highlighter_image(self):
        factor = 3
        size = 45 * factor
        img = Image.new("RGBA", (size, size), (255, 255, 255, 255))
        draw = ImageDraw.Draw(img)
        c_magenta, c_yellow, c_cyan = (255, 179, 255, 255), (255, 255, 179, 255), (179, 255, 255, 255)
        for x in range(size * 2):
            if x < (size * 2) // 3: color = c_magenta
            elif x < ((size * 2) // 3) * 2: color = c_yellow
            else: color = c_cyan
            draw.line([(x, 0), (0, x)], fill=color, width=factor)
        text = "Æ"
        try:
            font_path = "georgiai.ttf" if platform.system().lower() == "windows" else "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Italic.ttf"
            pil_font = ImageFont.truetype(font_path, 22 * factor)
        except:
            pil_font = ImageFont.load_default()
        bbox = draw.textbbox((0, 0), text, font=pil_font)
        w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text(((size-w)//2, (size-h)//2 - 4*factor), text, fill=(50, 50, 50, 255), font=pil_font)
        return ImageTk.PhotoImage(img.resize((45, 45), Image.Resampling.LANCZOS))

    def _get_system_fonts(self):
        try:
            self.root.update_idletasks()
            raw = list(font.families(root=self.root))
        except Exception:
            try:
                raw = list(font.families())
            except Exception:
                raw = []

        seen = set()
        cleaned = []
        for f in raw:
            if not f or f.startswith("@"):
                continue
            key = f.strip().lower()
            if key in seen:
                continue
            seen.add(key)
            cleaned.append(f.strip())

        common_fallback = ["Arial", "Times New Roman", "Courier New", "Comic Sans MS",
                            "Verdana", "Tahoma", "Georgia", "Calibri", "Trebuchet MS"]
        for cf in common_fallback:
            if cf.lower() not in seen:
                cleaned.append(cf)
                seen.add(cf.lower())

        cleaned.sort(key=lambda s: s.lower())
        return cleaned if cleaned else ["Arial"]

    def create_menu(self):
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Nuovo", accelerator="Ctrl+N", command=self.new_file)
        file_menu.add_command(label="Apri...", accelerator="Ctrl+O", command=self.open_file)
        file_menu.add_command(label="Salva", accelerator="Ctrl+S", command=self.save_file)
        file_menu.add_command(label="Salva con nome...", accelerator="Ctrl+Shift+S", command=self.save_as_file)
        file_menu.add_separator()
        file_menu.add_command(label="Stampa...", accelerator="Ctrl+P", command=self.print_document)
        file_menu.add_separator()
        file_menu.add_command(label="Esci", accelerator="Ctrl+Q", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        insert_menu = tk.Menu(menubar, tearoff=0)
        insert_menu.add_command(label="Inserisci immagine...", command=self.insert_image_dialog)
        menubar.add_cascade(label="Inserisci", menu=insert_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Attiva/Disattiva Dark Mode", command=self.toggle_dark_mode)
        menubar.add_cascade(label="Vista", menu=view_menu)
        self.root.config(menu=menubar)

    def setup_global_shortcuts(self):
        self.root.bind("<Control-n>", lambda e: self.new_file())
        self.root.bind("<Control-o>", lambda e: self.open_file())
        self.root.bind("<Control-s>", lambda e: self.save_file())
        self.root.bind("<Control-P>", lambda e: self.print_document())
        self.root.bind("<Control-q>", lambda e: self.root.quit())
        self.root.bind("<Control-a>", lambda e: self.select_all())
        self.root.bind("<Control-z>", lambda e: self.trigger_undo())
        self.root.bind("<Control-Alt-x>", lambda e: self.trigger_redo())

    def create_toolbar(self):
        self.toolbar = ttk.Frame(self.root, padding=5)
        self.toolbar.pack(side=tk.TOP, fill=tk.X)

        self.all_fonts = self._get_system_fonts()
        self.font_family = tk.StringVar(value="Arial" if "Arial" in self.all_fonts else self.all_fonts[0])
        self.font_cb = ttk.Combobox(self.toolbar, textvariable=self.font_family, values=self.all_fonts, width=15, state="readonly")
        self.font_cb.pack(side=tk.LEFT, padx=5)
        self.font_cb.bind("<<ComboboxSelected>>", self.change_font)
        self.font_cb.bind("<ButtonPress-1>", self._remember_format_selection, add="+")
        self.font_cb.bind("<KeyPress>", self._remember_format_selection, add="+")

        self.font_size = tk.IntVar(value=12)
        # Grandezza base del testo NON formattato: non deve cambiare quando
        # l'utente cambia il controllo dimensione (quello vale solo per il testo
        # scritto successivamente). Serve anche a non "ricolorare" le lettere
        # gia' presenti quando si cambia grandezza/zoom.
        self.default_font_size = 12
        self.size_cb = ttk.Combobox(self.toolbar, textvariable=self.font_size, values=list(range(8, 73, 2)), width=4, state="readonly")
        self.size_cb.pack(side=tk.LEFT, padx=5)
        self.size_cb.bind("<<ComboboxSelected>>", self.change_font)
        self.size_cb.bind("<ButtonPress-1>", self._remember_format_selection, add="+")
        self.size_cb.bind("<KeyPress>", self._remember_format_selection, add="+")

        ttk.Separator(self.toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        self.btn_bold = ttk.Button(self.toolbar, text="B", width=3, command=self.toggle_bold)
        self.btn_bold.pack(side=tk.LEFT, padx=1)
        self._add_tooltip(self.btn_bold, "Grassetto")
        self.btn_italic = ttk.Button(self.toolbar, text="I", width=3, command=self.toggle_italic)
        self.btn_italic.pack(side=tk.LEFT, padx=1)
        self._add_tooltip(self.btn_italic, "Corsivo")
        self.btn_underline = ttk.Button(self.toolbar, text="U", width=3, command=self.toggle_underline)
        self.btn_underline.pack(side=tk.LEFT, padx=1)
        self._add_tooltip(self.btn_underline, "Sottolineato")

        ttk.Separator(self.toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        self.btn_left = ttk.Button(self.toolbar, text="⌸ L", width=4, command=lambda: self.set_alignment("left"))
        self.btn_left.pack(side=tk.LEFT, padx=1)
        self._add_tooltip(self.btn_left, "Allinea a sinistra")
        self.btn_center = ttk.Button(self.toolbar, text="⌸ C", width=4, command=lambda: self.set_alignment("center"))
        self.btn_center.pack(side=tk.LEFT, padx=1)
        self._add_tooltip(self.btn_center, "Allinea al centro")
        self.btn_right = ttk.Button(self.toolbar, text="⌸ R", width=4, command=lambda: self.set_alignment("right"))
        self.btn_right.pack(side=tk.LEFT, padx=1)
        self._add_tooltip(self.btn_right, "Allinea a destra")

        ttk.Separator(self.toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        self.btn_add_text = ttk.Button(self.toolbar, text="+ Casella di Testo", command=self.toggle_add_text_mode)
        self.btn_add_text.pack(side=tk.LEFT, padx=1)
        self._add_tooltip(self.btn_add_text, "Aggiungi una casella di testo: attiva, poi clicca sul foglio dove vuoi posizionarla")

        ttk.Separator(self.toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        picker_frame = ttk.Frame(self.toolbar)
        picker_frame.pack(side=tk.LEFT, padx=5)

        self.recent_buttons = []
        for i in range(4):
            r, c = i // 2, i % 2
            btn = tk.Button(picker_frame, bg=self.recent_colors[i], image=self.pixel_img,
                            width=18, height=18, compound="center", bd=1,
                            command=lambda idx=i: self.select_recent_slot(idx))
            btn.grid(row=r, column=c, padx=1, pady=1)
            self.recent_buttons.append(btn)
        self.recent_buttons[self.selected_slot_idx].config(relief="sunken", bd=2)

        self.main_color_btn = tk.Button(picker_frame, bg=self.current_color, image=self.pixel_img, width=45, height=45, compound="center", relief="groove", command=lambda: self.toggle_color_popup("text"))
        self.main_color_btn.grid(row=0, column=2, rowspan=2, padx=(6, 0))

        self.main_bg_btn = tk.Button(picker_frame, image=self.bg_button_img, width=45, height=45, compound="center", relief="groove", command=lambda: self.toggle_color_popup("bg"))
        self.main_bg_btn.grid(row=0, column=3, rowspan=2, padx=(4, 0))

        zoom_container = ttk.Frame(self.toolbar)
        zoom_container.pack(side=tk.RIGHT, padx=10)

        self.lbl_zoom_val = ttk.Label(zoom_container, text="100%", font=("Arial", 9, "bold"))
        self.lbl_zoom_val.pack(side=tk.RIGHT, padx=(2, 5))
        self.scale_zoom = ttk.Scale(zoom_container, orient=tk.HORIZONTAL, from_=0.5, to=2.0, value=1.0, length=100, command=self.on_zoom_change)
        self.scale_zoom.pack(side=tk.RIGHT, padx=5)
        ttk.Label(zoom_container, text="Zoom:").pack(side=tk.RIGHT, padx=2)

        ttk.Separator(self.toolbar, orient=tk.VERTICAL).pack(side=tk.RIGHT, fill=tk.Y, padx=10)

        self.orientation_cb = ttk.Combobox(self.toolbar, values=["Verticale", "Orizzontale"], width=10, state="readonly")
        self.orientation_cb.set(self.current_orientation)
        self.orientation_cb.pack(side=tk.RIGHT, padx=5)
        self.orientation_cb.bind("<<ComboboxSelected>>", self.on_layout_change)
        ttk.Label(self.toolbar, text="Orientamento:").pack(side=tk.RIGHT, padx=2)

        self.format_cb = ttk.Combobox(self.toolbar, values=["A4", "A3", "A2"], width=5, state="readonly")
        self.format_cb.set(self.current_format)
        self.format_cb.pack(side=tk.RIGHT, padx=5)
        self.format_cb.bind("<<ComboboxSelected>>", self.on_layout_change)
        ttk.Label(self.toolbar, text="Formato:").pack(side=tk.RIGHT, padx=2)

    def create_workspace(self):
        self.workspace_frame = ttk.Frame(self.root)
        self.workspace_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.top_ruler = ttk.Frame(self.workspace_frame, height=30)
        self.top_ruler.pack(side=tk.TOP, fill=tk.X, padx=(45, 0))

        self.lbl_m_left = ttk.Label(self.top_ruler, text="Margine Sinistro:", font=("Arial", 8))
        self.lbl_m_left.pack(side=tk.LEFT, padx=2)
        self.scale_m_left = ttk.Scale(self.top_ruler, from_=10, to=150, value=30, length=120, command=self.on_margin_change)
        self.scale_m_left.pack(side=tk.LEFT, padx=5)

        self.lbl_m_right = ttk.Label(self.top_ruler, text="Margine Destro:", font=("Arial", 8))
        self.lbl_m_right.pack(side=tk.LEFT, padx=(15, 5))
        self.scale_m_right = ttk.Scale(self.top_ruler, from_=10, to=150, value=30, length=120, command=self.on_margin_change)
        self.scale_m_right.pack(side=tk.LEFT, padx=5)

        self.lower_workspace = ttk.Frame(self.workspace_frame)
        self.lower_workspace.pack(fill=tk.BOTH, expand=True)

        self.left_ruler = ttk.Frame(self.lower_workspace, width=50)
        self.left_ruler.pack(side=tk.LEFT, fill=tk.Y, pady=20)

        ttk.Label(self.left_ruler, text="M. Sup", font=("Arial", 7)).pack(pady=2)
        self.scale_m_top = ttk.Scale(self.left_ruler, orient=tk.VERTICAL, from_=10, to=150, value=40, length=120, command=self.on_margin_change)
        self.scale_m_top.pack(pady=5)

        ttk.Label(self.left_ruler, text="M. Inf", font=("Arial", 7)).pack(pady=2)
        self.scale_m_bottom = ttk.Scale(self.left_ruler, orient=tk.VERTICAL, from_=10, to=150, value=40, length=120, command=self.on_margin_change)
        self.scale_m_bottom.pack(pady=5)

        self.right_sidebar = ttk.LabelFrame(self.lower_workspace, text=" Pagine Documento ", padding=5)
        self.right_sidebar.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0), pady=5)

        self.btn_add_page = ttk.Button(self.right_sidebar, text="+ Aggiungi Pagina", command=self.add_new_page)
        self.btn_add_page.pack(side=tk.TOP, fill=tk.X, pady=5)
        self._add_tooltip(self.btn_add_page, "Aggiungi una nuova pagina sotto quelle esistenti")

        self.sidebar_canvas = tk.Canvas(self.right_sidebar, width=150, bd=0, highlightthickness=0)
        self.sidebar_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.sidebar_scroll = ttk.Scrollbar(self.right_sidebar, orient=tk.VERTICAL, command=self.sidebar_canvas.yview)
        self.sidebar_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.sidebar_canvas.configure(yscrollcommand=self.sidebar_scroll.set)

        self.pages_list_frame = ttk.Frame(self.sidebar_canvas)
        self.sidebar_canvas.create_window((0, 0), window=self.pages_list_frame, anchor="nw")
        self.pages_list_frame.bind("<Configure>", lambda e: self.sidebar_canvas.configure(scrollregion=self.sidebar_canvas.bbox("all")))

        # Barra orizzontale in basso per spostarsi a destra/sinistra, soprattutto
        # quando si ingrandisce (zoom) un documento che passa dalla pipeline PDF.
        self.scrollbar_x = ttk.Scrollbar(self.lower_workspace, orient=tk.HORIZONTAL)
        self.scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.scrollbar_y = ttk.Scrollbar(self.lower_workspace, orient=tk.VERTICAL)
        self.scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas_desktop = tk.Canvas(self.lower_workspace, bg="#e0e0e0", bd=0, highlightthickness=0)
        self.pdf_canvas = tk.Canvas(self.lower_workspace, bg="#e0e0e0", bd=0, highlightthickness=0)

        self.canvas_desktop.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.canvas_desktop.configure(yscrollcommand=self.scrollbar_y.set)
        self.canvas_desktop.configure(xscrollcommand=self.scrollbar_x.set)
        self.scrollbar_y.config(command=self.canvas_desktop.yview)
        self.scrollbar_x.config(command=self.canvas_desktop.xview)
        self.canvas_desktop.configure(yscrollcommand=self._on_desktop_yview, xscrollcommand=self._on_desktop_xview)

        self.pdf_canvas.bind("<Button-1>", self.on_pdf_canvas_click)
        
        self.canvas_desktop.bind("<Configure>", lambda e: self.center_page_on_desktop())
        self.canvas_desktop.bind("<Configure>", lambda e: self._schedule_box_chrome_refresh(), add="+")
        self.pdf_canvas.bind("<Configure>", lambda e: self.center_page_on_desktop())

        for cv in (self.canvas_desktop, self.pdf_canvas):
            cv.bind("<Button-4>", self.handle_mouse_wheel)
            cv.bind("<Button-5>", self.handle_mouse_wheel)
            cv.bind("<MouseWheel>", self.handle_mouse_wheel)

    def handle_mouse_wheel(self, event):
        target_canvas = self.pdf_canvas if self.pdf_mode else self.canvas_desktop
        if event.state & 0x0004:
            if event.num == 4 or event.delta > 0:
                new_zoom = self.zoom_factor + 0.1
            elif event.num == 5 or event.delta < 0:
                new_zoom = self.zoom_factor - 0.1
            else:
                return

            new_zoom = max(0.5, min(2.0, new_zoom))
            self.zoom_factor = new_zoom
            self.scale_zoom.set(new_zoom)
            self.lbl_zoom_val.config(text=f"{int(new_zoom * 100)}%")

            if self._zoom_job is not None:
                self.root.after_cancel(self._zoom_job)
            self._zoom_job = self.root.after(80, self.update_page_layout)
            return "break"
        else:
            if event.num == 4:
                target_canvas.yview_scroll(-3, "units")
            elif event.num == 5:
                target_canvas.yview_scroll(3, "units")
            elif event.delta:
                target_canvas.yview_scroll(int(-1 * (event.delta / 120) * 3), "units")

    def _on_desktop_yview(self, first, last):
        self.scrollbar_y.set(first, last)
        self._schedule_box_chrome_refresh()

    def _on_desktop_xview(self, first, last):
        self.scrollbar_x.set(first, last)
        self._schedule_box_chrome_refresh()

    def _schedule_box_chrome_refresh(self):
        if getattr(self, "_box_chrome_refresh_job", None) is not None:
            try:
                self.root.after_cancel(self._box_chrome_refresh_job)
            except Exception:
                pass
        try:
            self._box_chrome_refresh_job = self.root.after_idle(self._refresh_selected_box_chrome)
        except Exception:
            self._box_chrome_refresh_job = None

    def _refresh_selected_box_chrome(self):
        self._box_chrome_refresh_job = None
        boxw = getattr(self, "active_pdf_box", None)
        if boxw is None or not getattr(boxw, "_rich", False):
            return
        try:
            if not boxw.winfo_exists() or not boxw.is_selected:
                return
            self._position_text_box_chrome(boxw)
        except Exception:
            pass

    def save_current_page_state(self):
        if not self.pdf_mode:
            for pid, ed in list(self.text_page_editors.items()):
                try:
                    self.pages_data[pid] = ed.get("1.0", tk.END + "-1c")
                    self.page_text_format[pid] = self._capture_editor_format(ed)
                except tk.TclError:
                    pass

    def _index_to_char_offset(self, editor, index):
        """Converte un indice 'lin.col' in un offset caratteri (0-based) nel
        testo della pagina ('1.0'+Nc). Fallback manuale se count/chars manca."""
        try:
            res = editor.count("1.0", index, "chars")
            if res:
                return int(res[0])
        except Exception:
            pass
        try:
            line, col = str(index).split(".")
            line, col = int(line), int(col)
        except Exception:
            return 0
        offset = 0
        for ln in range(1, line):
            offset += len(editor.get(f"{ln}.0", f"{ln}.end")) + 1
        return offset + col

    def _is_relevant_tag(self, tag):
        return (tag in self.font_tag_defs
                or tag in ("underline", "left", "center", "right", "rtf_underline",
                           "rtf_left", "rtf_center", "rtf_right")
                or tag.startswith(("color_", "bg_", "rtf_color_", "rtf_bg_")))

    def _capture_editor_format(self, editor):
        """Salva la formattazione dell'editor come runs [start_c, end_c, tag]
        a offset caratteri, cosi' da poterla ri-applicare dopo un rebuild."""
        runs = []
        try:
            for tag in editor.tag_names():
                if not self._is_relevant_tag(tag):
                    continue
                ranges = editor.tag_ranges(tag)
                for i in range(0, len(ranges), 2):
                    try:
                        start_c = self._index_to_char_offset(editor, ranges[i])
                        end_c = self._index_to_char_offset(editor, ranges[i + 1])
                    except Exception:
                        continue
                    if end_c > start_c:
                        runs.append([start_c, end_c, tag])
        except Exception:
            pass
        return runs

    def _restore_page_format(self, page_id, editor):
        """Ri-applica sul nuovo editor i runs di formattazione salvati."""
        try:
            runs = self.page_text_format.get(page_id, [])
            for start_c, end_c, tag in runs:
                try:
                    s = editor.index("1.0 + %dc" % start_c)
                    e = editor.index("1.0 + %dc" % end_c)
                except Exception:
                    continue
                if s == e:
                    continue
                if tag in self.font_tag_defs:
                    fam, sz, bold, italic = self.font_tag_defs[tag]
                    tag_name = self._get_font_tag(editor, fam, sz, bold, italic)
                    try:
                        editor.tag_add(tag_name, s, e)
                    except Exception:
                        pass
                elif tag.startswith("color_"):
                    hexc = tag[len("color_"):]
                    try:
                        editor.tag_configure(tag, foreground="#" + hexc)
                        editor.tag_add(tag, s, e)
                    except Exception:
                        pass
                elif tag.startswith("rtf_color_"):
                    hexc = tag[len("rtf_color_"):]
                    try:
                        editor.tag_configure(tag, foreground="#" + hexc)
                        editor.tag_add(tag, s, e)
                    except Exception:
                        pass
                elif tag.startswith("bg_"):
                    hexc = tag[len("bg_"):]
                    try:
                        editor.tag_configure(tag, background="#" + hexc)
                        editor.tag_add(tag, s, e)
                    except Exception:
                        pass
                elif tag.startswith("rtf_bg_"):
                    hexc = tag[len("rtf_bg_"):]
                    try:
                        editor.tag_configure(tag, background="#" + hexc)
                        editor.tag_add(tag, s, e)
                    except Exception:
                        pass
                elif tag == "rtf_underline":
                    try:
                        editor.tag_configure(tag, underline=True)
                        editor.tag_add(tag, s, e)
                    except Exception:
                        pass
                elif tag in ("underline", "left", "center", "right", "rtf_left", "rtf_center", "rtf_right"):
                    try:
                        if tag.startswith("rtf_") and tag != "rtf_underline":
                            editor.tag_configure(tag, justify=tag[4:])
                        editor.tag_add(tag, s, e)
                    except Exception:
                        pass
        except Exception:
            pass

    def _split_format_at(self, fmt, k):
        """Divide i runs di formattazione [sc, ec, tag] in due liste attorno
        all'offset caratteri k: a sinistra (<=k) e a destra (>k), con offset
        riallineati. Usato quando la paginazione spezza una pagina."""
        left, right = [], []
        for sc, ec, tag in fmt:
            if ec <= k:
                left.append([sc, ec, tag])
            elif sc >= k:
                right.append([sc - k, ec - k, tag])
            else:
                left.append([sc, k, tag])
                if k < ec:
                    right.append([0, ec - k, tag])
        return left, right

    def switch_to_page(self, page_id):
        self.active_page_id = page_id
        self.update_pages_sidebar()

        if self.pdf_mode:
            self.jump_to_pdf_page(page_id)
            return

        if page_id in self.pdf_page_offsets:
            self.root.update_idletasks()
            region = self.canvas_desktop.cget("scrollregion")
            try:
                parts = [float(v) for v in str(region).split()]
                total_h = parts[3] if len(parts) == 4 and parts[3] > 0 else 1.0
            except Exception:
                total_h = 1.0
            y = self.pdf_page_offsets[page_id]
            frac = y / max(1.0, total_h)
            self.canvas_desktop.yview_moveto(max(0.0, min(1.0, frac)))

        if page_id in self.text_page_editors:
            self.text_page_editors[page_id].focus_set()

    def add_new_page(self):
        if self.pdf_mode:
            new_id = max(self.pdf_pages_info.keys()) + 1 if self.pdf_pages_info else 1
            
            # Pagina vuota A4 resa alla scala standard (595x842 px a zoom 1.0),
            # coerente con le pagine caricate da load_pdf (fit nel foglio A4).
            dpi = 72
            img = Image.new("RGB", (int(595 * dpi/72), int(842 * dpi/72)), "white")
            b_io = io.BytesIO()
            img.save(b_io, format="PNG")
            img_bytes = b_io.getvalue()
            
            self.pdf_pages_info[new_id] = {
                "spans": [],
                "img_bytes": img_bytes,
                "width_pt": 595,
                "height_pt": 842,
                "render_dpi": dpi,
                "overlays": [],
                "display_name": f"Pagina Vuota {new_id}"
            }
            self.render_all_pdf_pages()
            self.update_pages_sidebar()
            self.jump_to_pdf_page(new_id)
            return

        self.save_current_page_state()
        new_id = max(self.pages_data.keys()) + 1 if self.pages_data else 1
        self.pages_data[new_id] = ""
        self.page_text_boxes[new_id] = []
        self.page_text_box_widgets[new_id] = []
        self.update_page_layout()
        self.switch_to_page(new_id)

    def insert_page_after(self, page_id):
        """Inserisce una nuova pagina vuota subito dopo page_id, rinumerando e
        spostando quelle successive. Usato per lo sfondamento automatico di
        pagina quando il testo supera il fondo del foglio."""
        self.save_current_page_state()
        sorted_ids = sorted(self.pages_data.keys())

        new_pages_data = {}
        new_text_boxes = {}
        new_text_format = {}
        new_images = {}
        new_display_names = {}
        next_id = 1
        inserted_id = None
        for old_id in sorted_ids:
            new_pages_data[next_id] = self.pages_data[old_id]
            new_text_boxes[next_id] = self.page_text_boxes.get(old_id, [])
            new_text_format[next_id] = self.page_text_format.get(old_id, [])
            new_images[next_id] = self.page_images.get(old_id, [])
            if old_id in self.page_display_names:
                new_display_names[next_id] = self.page_display_names[old_id]
            cur_id = next_id
            next_id += 1
            if old_id == page_id:
                inserted_id = next_id
                new_pages_data[next_id] = ""
                new_text_boxes[next_id] = []
                new_text_format[next_id] = []
                new_images[next_id] = []
                next_id += 1

        if inserted_id is None:
            # page_id non trovato (caso limite): aggiunge semplicemente in coda
            inserted_id = next_id
            new_pages_data[next_id] = ""
            new_text_boxes[next_id] = []
            new_text_format[next_id] = []
            new_images[next_id] = []

        self.pages_data = new_pages_data
        self.page_text_boxes = new_text_boxes
        self.page_text_format = new_text_format
        self.page_images = new_images
        self.page_display_names = new_display_names
        self.page_text_box_widgets = {k: [] for k in self.pages_data.keys()}
        self.page_image_widgets = {k: [] for k in self.pages_data.keys()}
        return inserted_id

    def delete_page(self, page_id):
        if self.pdf_mode:
            return
        if len(self.pages_data) <= 1:
            messagebox.showwarning("Attenzione", "Impossibile eliminare l'unica pagina presente.")
            return
        if not messagebox.askyesno("Conferma eliminazione", f"Eliminare {self._get_page_display_name(page_id, False)}?"):
            return

        self._clear_text_box_widgets(page_id)
        self._clear_image_widgets(page_id)
        if page_id in self.text_page_editors:
            try:
                self.text_page_editors[page_id].destroy()
            except Exception:
                pass
            del self.text_page_editors[page_id]

        self.pages_data.pop(page_id, None)
        self.page_text_boxes.pop(page_id, None)
        self.page_text_format.pop(page_id, None)
        self.page_images.pop(page_id, None)
        self.page_image_widgets.pop(page_id, None)
        self.page_display_names.pop(page_id, None)

        sorted_ids = sorted(self.pages_data.keys())
        new_pages_data = {}
        new_text_boxes = {}
        new_text_format = {}
        new_images = {}
        new_display_names = {}
        for i, old_id in enumerate(sorted_ids):
            new_id = i + 1
            new_pages_data[new_id] = self.pages_data[old_id]
            new_text_boxes[new_id] = self.page_text_boxes.get(old_id, [])
            new_text_format[new_id] = self.page_text_format.get(old_id, [])
            new_images[new_id] = self.page_images.get(old_id, [])
            if old_id in self.page_display_names:
                new_display_names[new_id] = self.page_display_names[old_id]
                
        self.pages_data = new_pages_data
        self.page_text_boxes = new_text_boxes
        self.page_text_format = new_text_format
        self.page_images = new_images
        self.page_display_names = new_display_names
        self.page_text_box_widgets = {k: [] for k in self.pages_data.keys()}
        self.page_image_widgets = {k: [] for k in self.pages_data.keys()}

        if self.active_page_id not in self.pages_data:
            self.active_page_id = max(self.pages_data.keys())
            
        self.update_page_layout()
        self.switch_to_page(self.active_page_id)

    def delete_pdf_page(self, page_id):
        if len(self.pdf_pages_info) <= 1:
            messagebox.showwarning("Attenzione", "Impossibile eliminare l'unica pagina presente.")
            return
        if not messagebox.askyesno("Conferma eliminazione", f"Eliminare {self._get_page_display_name(page_id, True)}?"):
            return

        del self.pdf_pages_info[page_id]
        sorted_ids = sorted(self.pdf_pages_info.keys())
        new_info = {}
        for i, old_id in enumerate(sorted_ids):
            new_info[i + 1] = self.pdf_pages_info[old_id]
        self.pdf_pages_info = new_info

        if self.active_page_id not in self.pdf_pages_info:
            self.active_page_id = max(self.pdf_pages_info.keys())

        self.render_all_pdf_pages()
        self.update_pages_sidebar()

    def _get_page_display_name(self, page_id, is_pdf):
        if is_pdf:
            return self.pdf_pages_info.get(page_id, {}).get("display_name", f"Pagina {page_id}")
        return self.page_display_names.get(page_id, f"Foglio {page_id}")

    def _set_page_display_name(self, page_id, is_pdf, name):
        if is_pdf:
            if page_id in self.pdf_pages_info:
                self.pdf_pages_info[page_id]["display_name"] = name
        else:
            self.page_display_names[page_id] = name

    def _start_rename(self, frame, page_id, is_pdf):
        for w in frame.winfo_children():
            w.destroy()
        var = tk.StringVar(value=self._get_page_display_name(page_id, is_pdf))
        entry = ttk.Entry(frame, textvariable=var, width=14)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        entry.focus_set()
        entry.select_range(0, tk.END)

        committed = {"done": False}

        def commit(event=None):
            if committed["done"]:
                return
            committed["done"] = True
            new_name = var.get().strip() or self._get_page_display_name(page_id, is_pdf)
            self._set_page_display_name(page_id, is_pdf, new_name)
            self.update_pages_sidebar()

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)

    def update_pages_sidebar(self):
        for widget in self.pages_list_frame.winfo_children():
            widget.destroy()

        is_pdf = self.pdf_mode
        keys = sorted(self.pdf_pages_info.keys()) if is_pdf else sorted(self.pages_data.keys())

        for p_id in keys:
            p_frame = ttk.Frame(self.pages_list_frame, padding=1)
            p_frame.pack(fill=tk.X, pady=2)

            name = self._get_page_display_name(p_id, is_pdf)
            btn = ttk.Button(p_frame, text=name, width=8, command=lambda idx=p_id: self.switch_to_page(idx))
            btn.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self._add_tooltip(btn, f"Scorri fino a: {name}")

            btn_rename = tk.Button(p_frame, text="✎", bd=0, width=1, font=("Arial", 8),
                                    command=lambda idx=p_id, f=p_frame, ip=is_pdf: self._start_rename(f, idx, ip))
            btn_rename.pack(side=tk.LEFT, padx=(1, 0))
            self._add_tooltip(btn_rename, "Rinomina")

            can_delete = (len(self.pdf_pages_info) if is_pdf else len(self.pages_data)) > 1
            if can_delete:
                del_cmd = (lambda idx=p_id: self.delete_pdf_page(idx)) if is_pdf else (lambda idx=p_id: self.delete_page(idx))
                btn_del = tk.Button(p_frame, text="🗑", fg="#b3261e", bd=0, width=1, font=("Arial", 8), command=del_cmd)
                btn_del.pack(side=tk.LEFT, padx=(1, 2))
                self._add_tooltip(btn_del, "Elimina pagina")

    def _extract_text_pptx(self, file_path):
        """Estrae il testo di ogni slide come pagina separata, senza LibreOffice."""
        if not HAS_PPTX:
            raise RuntimeError("La libreria 'python-pptx' non è installata (pip install python-pptx).")
        prs = Presentation(file_path)
        pages = []
        for slide in prs.slides:
            lines = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        text = "".join(run.text for run in para.runs) or para.text
                        if text:
                            lines.append(text)
                elif shape.has_table:
                    for row in shape.table.rows:
                        row_text = "\t".join(cell.text for cell in row.cells)
                        if row_text.strip():
                            lines.append(row_text)
            pages.append("\n".join(lines))
        return pages  # una stringa per slide/pagina

    def _extract_text_epub(self, file_path):
        """Estrae il testo di ogni capitolo (in ordine di lettura).
        Parsing manuale via container.xml -> OPF -> spine, letto direttamente
        dallo ZIP: piu' tollerante di ebooklib verso epub reali con XML/OPF
        non perfettamente validi, ed elimina la dipendenza da ebooklib."""
        def strip_html(raw_html):
            raw_html = re.sub(r"(?is)<(script|style).*?</\1>", "", raw_html)
            raw_html = re.sub(r"(?i)<br\s*/?>", "\n", raw_html)
            raw_html = re.sub(r"(?i)</p>", "\n\n", raw_html)
            raw_html = re.sub(r"(?i)</(div|h[1-6]|li)>", "\n", raw_html)
            text = re.sub(r"(?s)<[^>]+>", "", raw_html)
            text = html_lib.unescape(text)
            text = re.sub(r"[ \t]+", " ", text)
            text = re.sub(r"\n{3,}", "\n\n", text).strip()
            return text

        with zipfile.ZipFile(file_path, "r") as z:
            names = set(z.namelist())

            # 1. Trova il file OPF tramite META-INF/container.xml
            opf_path = None
            try:
                with z.open("META-INF/container.xml") as f:
                    croot = ET.parse(f).getroot()
                for rootfile in croot.iter():
                    if rootfile.tag.endswith("rootfile"):
                        opf_path = rootfile.get("full-path")
                        break
            except Exception:
                opf_path = None

            if not opf_path or opf_path not in names:
                # Fallback: cerca un file .opf ovunque nell'archivio
                candidates = [n for n in names if n.lower().endswith(".opf")]
                opf_path = candidates[0] if candidates else None

            spine_files = []
            if opf_path:
                opf_dir = os.path.dirname(opf_path)
                with z.open(opf_path) as f:
                    oroot = ET.parse(f).getroot()

                manifest = {}
                for item in oroot.iter():
                    if item.tag.endswith("item"):
                        item_id = item.get("id")
                        href = item.get("href")
                        media_type = item.get("media-type", "")
                        if item_id and href and ("html" in media_type or "xml" in media_type):
                            path = href if not opf_dir else f"{opf_dir}/{href}"
                            path = path.replace("\\", "/")
                            manifest[item_id] = path

                for itemref in oroot.iter():
                    if itemref.tag.endswith("itemref"):
                        idref = itemref.get("idref")
                        if idref in manifest and manifest[idref] in names:
                            spine_files.append(manifest[idref])

            if not spine_files:
                # Ultimo fallback: prendi tutti gli .xhtml/.html nell'archivio in ordine
                spine_files = sorted(n for n in names if n.lower().endswith((".xhtml", ".html", ".htm")))

            chapters = []
            for path in spine_files:
                try:
                    with z.open(path) as f:
                        raw_html = f.read().decode("utf-8", errors="ignore")
                    text = strip_html(raw_html)
                    if text:
                        chapters.append(text)
                except Exception:
                    continue

        if not chapters:
            raise RuntimeError("Nessun capitolo leggibile trovato nell'EPUB (struttura non riconosciuta).")

        return "\n\n".join(chapters)

    def _extract_text_opendocument(self, file_path):
        """Estrae il testo da .odt/.ods leggendo direttamente content.xml (formato ZIP+XML), senza LibreOffice."""
        ns = {"text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
              "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0"}
        with zipfile.ZipFile(file_path, "r") as z:
            with z.open("content.xml") as f:
                tree = ET.parse(f)
        root = tree.getroot()
        lines = []
        # Paragrafi e intestazioni di testo
        for para in root.iter():
            tag = para.tag.split("}")[-1]
            if tag in ("p", "h"):
                text = "".join(para.itertext())
                lines.append(text)
        return "\n".join(lines)

    def _dominant_docx_font(self, doc):
        """Trova il font e la dimensione piu' usati nel documento .docx, per
        aprirlo mantenendo l'aspetto originale invece del font di default."""
        names = {}
        sizes = {}
        for p in doc.paragraphs:
            for run in p.runs:
                try:
                    if run.font and run.font.name:
                        names[run.font.name] = names.get(run.font.name, 0) + 1
                except Exception:
                    pass
                try:
                    if run.font and run.font.size:
                        pt = int(round(run.font.size.pt))
                        sizes[pt] = sizes.get(pt, 0) + 1
                except Exception:
                    pass
        fam = max(names, key=names.get) if names else None
        size = max(sizes, key=sizes.get) if sizes else None
        return fam, size

    def _dominant_doc_font(self, file_path):
        """Trova il font e la dimensione piu' usati in un file .doc (Word 97-2003)
        leggendo la font table (SttbfFfn) dal WordDocument stream. Se la
        struttura della font table non e' riconoscibile o la lettura fallisce,
        restituisce (None, None) senza errori file_path esterni."""
        try:
            out = self._read_doc_font_table(file_path)
            return out if out else (None, None)
        except Exception:
            return None, None

    def _read_doc_font_table(self, file_path):
        """Legge effettivamente la font table da un .doc. Restituisce
        (font_maggiore, None) o None. Separo questa logica per poter
        fallire in modo pulito senza propagare eccezioni."""
        import olefile
        ole = olefile.OleFileIO(file_path)
        try:
            if not ole.exists("WordDocument"):
                return None
            ws = ole.openstream("WordDocument").read()

            # Ricalcola la posizione del fibRgFcLcb in base alla versione del FIB
            # (nFib -> csw/cslw variabili), cosi' si leggono gli offset corretti
            # anche per documenti Word piu' vecchi, invece di offset fissi.
            try:
                n_fib = struct.unpack_from("<H", ws, 2)[0]
                csw = struct.unpack_from("<H", ws, 32)[0]
                cslw = struct.unpack_from("<H", ws, 34 + csw * 2)[0]
                cb_rg_fc_lcb = struct.unpack_from("<H", ws, 34 + csw * 2 + 2 + cslw * 4)[0]
                fc_start = 34 + csw * 2 + 2 + cslw * 4 + 2
            except Exception:
                return None

            # fcSttbfFfn e' un campo nel fibRgFcLcb; per i FIB piu' comuni
            # e' il campo di indice 4 (fc= trovare coppia con lcb plausibile).
            # Cerca nella tabella una coppia il cui lcb sia plausibile per una
            # string table UTF-16LE (>= 10 byte) e leggila.
            best = None
            for i in range(0, cb_rg_fc_lcb - 7, 8):
                fc = struct.unpack_from("<I", ws, fc_start + i)[0]
                lcb = struct.unpack_from("<I", ws, fc_start + i + 4)[0]
                if 0 < fc < len(ws) and 10 <= lcb <= 4000 and fc + lcb <= len(ws):
                    # verifica che cominci con una struttura Sttbf plausibile
                    probe = ws[fc:fc + min(lcb, 24)]
                    if len(probe) >= 4:
                        cb_entry = struct.unpack_from("<H", probe, 0)[0]
                        if 4 <= cb_entry <= 200:
                            best = (fc, lcb)
                            break

            if best is None:
                return None

            fc_font, lcb_font = best
            font_data = ws[fc_font:fc_font + lcb_font]
            if len(font_data) < 4:
                return None

            cb_entry = struct.unpack_from("<H", font_data, 0)[0]
            entry_start = 4
            names = []
            while entry_start + cb_entry <= len(font_data) and cb_entry > 0:
                entry = font_data[entry_start:entry_start + cb_entry]
                # Sttbf Ffn: ogni entry e' un Ffn dove il nome e' UTF-16LE.
                # Il campo 'cch' (lunghezza nome in caratteri) e' a offset 41
                # (1 byte) in Ffn, il nome segue alla fine della struttura.
                try:
                    if cb_entry >= 42:
                        cch = entry[41]
                        name_bytes = entry[42:42 + cch * 2]
                        name = name_bytes.decode("utf-16-le", errors="ignore").strip()
                        if name and not name.startswith("@") and name.isprintable():
                            names.append(name)
                except Exception:
                    pass

                if names:
                    # Il primo nome restituito e' in genere il font usato da piu'
                    # caratteri in Word (primo della tabella)
                    return (names[0], None)

                entry_start += cb_entry
                if entry_start + 4 > len(font_data):
                    break

            if names:
                return (names[0], None)
            return None
        finally:
            ole.close()

    def _extract_text_doc_legacy(self, file_path):
        """Estrae il testo completo da un file .doc (Word 97-2003) usando
        una catena di metodi in ordine di affidabilita':
          1. catdoc (CLI leggero, estrae il testo semantico correttamente);
          2. LibreOffice headless (soffice --convert-to txt), molto fedele;
          3. parser interno olefile + Piece Table (senza comandi esterni);
          4. scansione euristica del flusso binario (ultima spiaggia).
        Questo garantisce la compatibilita' con la quasi totalita' dei .doc,
        inclusi documenti prodotti da Word vecchi o da altri programmi."""
        import olefile
        import struct

        # --- 1 e 2: comandi esterni affidabili (catdoc / LibreOffice) ---
        text = self._extract_doc_via_external(file_path)
        if text:
            return text.strip()

        # --- 3: parser interno olefile + Piece Table ---
        text = self._extract_doc_via_olefile(file_path)
        if text:
            return text.strip()

        # --- 4: ultimo fallback euristico ---
        return self._fallback_extract_doc_text_raw(file_path)

    def _extract_doc_via_external(self, file_path):
        """Prova a estrarre il testo con catdoc, poi con LibreOffice headless.
        Ritorna il testo estratto o stringa vuota."""
        import shutil
        import tempfile

        # 1) catdoc: leggero e veloce, gia' presente su molti sistemi
        if shutil.which("catdoc"):
            try:
                proc = subprocess.run(
                    ["catdoc", file_path],
                    capture_output=True, timeout=30
                )
                if proc.returncode == 0 and proc.stdout and proc.stdout.strip():
                    return proc.stdout.decode("utf-8", errors="replace").strip()
            except Exception:
                pass

        # 2) LibreOffice headless (soffice/libreoffice)
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice:
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    proc = subprocess.run(
                        [soffice, "--headless", "--convert-to", "txt:Text",
                         "--outdir", tmpdir, file_path],
                        capture_output=True, timeout=90
                    )
                    if proc.returncode == 0:
                        # il file esportato ha lo stesso nome base + .txt
                        base = os.path.splitext(os.path.basename(file_path))[0]
                        out_path = os.path.join(tmpdir, base + ".txt")
                        if os.path.exists(out_path):
                            with open(out_path, "r", encoding="utf-8", errors="replace") as f:
                                return f.read().lstrip("\ufeff").strip()
            except Exception:
                pass

        return ""

    def _convert_doc_to_pdf(self, file_path):
        """Converte un file .doc (Word 97-2003) in PDF usando LibreOffice headless.
        Ritorna (pdf_path, cleanup_callable). Riporta file che contengono campi
        Word (TOC/PAGEREF/HYPERLINK) gia' espansi e le immagini originali,
        esattamente come apparirebbero stampati. Ritorna (None, None) se
        LibreOffice non e' disponibile o la conversione fallisce."""
        import shutil
        import tempfile

        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            return None, None

        tmpdir = tempfile.mkdtemp(prefix="simplewright_doc_")
        try:
            proc = subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf",
                 "--outdir", tmpdir, file_path],
                capture_output=True, timeout=120
            )
            base = os.path.splitext(os.path.basename(file_path))[0]
            pdf_path = os.path.join(tmpdir, base + ".pdf")
            if proc.returncode == 0 and os.path.exists(pdf_path):
                import shutil as _shutil
                return pdf_path, lambda: _shutil.rmtree(tmpdir, ignore_errors=True)
            _shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            import shutil as _shutil
            _shutil.rmtree(tmpdir, ignore_errors=True)
        return None, None

    def _convert_pptx_to_pdf(self, file_path):
        """Converte un file .pptx in PDF usando LibreOffice headless, cosi' il
        contenuto (testo, immagini, tabelle, gruppi) viene reso in modo fedele
        dalla pipeline PDF di PyMuPDF invece che dal parser python-pptx, che su
        certe presentazioni bloccava il caricamento o non mostrava le immagini.
        Ritorna (pdf_path, cleanup_callable), oppure (None, None) se LibreOffice
        manca o la conversione fallisce."""
        import shutil
        import tempfile

        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            return None, None

        tmpdir = tempfile.mkdtemp(prefix="simplewright_pptx_")
        try:
            proc = subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf",
                 "--outdir", tmpdir, file_path],
                capture_output=True, timeout=180
            )
            base = os.path.splitext(os.path.basename(file_path))[0]
            pdf_path = os.path.join(tmpdir, base + ".pdf")
            if proc.returncode == 0 and os.path.exists(pdf_path):
                import shutil as _shutil
                return pdf_path, lambda: _shutil.rmtree(tmpdir, ignore_errors=True)
            _shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            import shutil as _shutil
            _shutil.rmtree(tmpdir, ignore_errors=True)
        return None, None

    def _extract_doc_via_olefile(self, file_path):
        """Estrae il testo da un .doc parsegando il WordDocument stream e la
        Piece Table (Clx) dallo stream tabella, usando esclusivamente olefile
        (nessun comando esterno). Funziona sia per Unicode sia ANSI."""
        import olefile

        try:
            ole = olefile.OleFileIO(file_path)
        except Exception:
            return ""

        try:
            if not ole.exists("WordDocument"):
                return ""
            ws = ole.openstream("WordDocument").read()
            if len(ws) < 32:
                return ""

            # Determina quale stream tabella usare (bit 9 dei flags del FIB)
            flags = struct.unpack_from("<H", ws, 10)[0]
            f_which_tbl_stm = bool(flags & 0x0200)
            table_name_primary = "1Table" if f_which_tbl_stm else "0Table"
            table_name = table_name_primary
            if not ole.exists(table_name):
                table_name = "0Table" if table_name_primary == "1Table" else "1Table"
                if not ole.exists(table_name):
                    return ""
            table_stream = ole.openstream(table_name).read()

            # Ricalcola posizione del fibRgFcLcb secondo nFib
            try:
                n_fib = struct.unpack_from("<H", ws, 2)[0]
                csw = struct.unpack_from("<H", ws, 32)[0]
                cslw_pos = 34 + csw * 2
                cslw = struct.unpack_from("<H", ws, cslw_pos)[0]
                cb_rg_fc_lcb = struct.unpack_from("<H", ws, cslw_pos + 2 + cslw * 4)[0]
                fc_start = cslw_pos + 2 + cslw * 4 + 2
            except Exception:
                return ""

            # Identifica, tra le coppie fc/lcb, quelle per le strutture che
            # servono. Nota: non conosciamo a priori l'indice esatto di fcClx
            # in ogni versione, quindi cerchiamo la Piece Table direttamente
            # nello stream tabella usando la struttura della Clx.

            # --- Trova la Clx nello stream tabella ---
            clx_off = self._find_piece_table_offset(table_stream)
            if clx_off is None:
                return ""

            text = self._extract_doc_text_from_clx(table_stream[clx_off:], ws)
            return text.strip()
        finally:
            ole.close()

    def _find_piece_table_offset(self, table_stream):
        """Cerca nello stream tabella l'offset della Clx (struttura con clxt=2).
        Cerca un punto in cui: byte = 0x02, poi 4 byte di padding (0x00000000),
        poi lcbPlcPcd (4 byte) con valore plausibile, e infine il PLC con una
        sequenza di CP iniziali validi (piccoli e crescenti). Ritorna l'offset
        del byte clxt=2, oppure None."""
        n = len(table_stream)
        for off in range(n - 12):
            if table_stream[off] != 0x02:
                continue
            # padding quasi sempre 0
            if table_stream[off + 1] != 0x00 or table_stream[off + 2] != 0x00:
                continue
            try:
                lcb = struct.unpack_from("<I", table_stream, off + 5)[0]
            except Exception:
                continue
            if lcb < 16 or lcb > n - off or off + 9 + lcb > n:
                continue
            plc = table_stream[off + 9:off + 9 + lcb]
            if len(plc) < 12:
                continue
            n_pieces = (lcb - 4) // 12
            if n_pieces <= 0 or n_pieces > 50000:
                continue
            # verifica la monotonia crescente dei primi CP
            try:
                vals = [struct.unpack_from("<i", plc, 4 + i * 4)[0] for i in range(min(8, n_pieces + 1))]
            except Exception:
                continue
            if not vals or vals[0] < 0 or vals[0] > 100:
                continue
            if any(vals[i] > vals[i + 1] for i in range(len(vals) - 1)):
                continue
            return off
        return None

    def _extract_doc_text_from_clx(self, clx_data, word_stream):
        """Estrae il testo usando la Piece Table (Clx). clx_data e' il resto
        dello stream tabella che inizia con il byte clxt. Gestisce correttamente
        pezzi Unicode e ANSI (compressed), usando i CP per calcolare le lunghezze
        e la fc nel Pcd per la posizione fisica dentro WordDocument."""
        if not clx_data or len(clx_data) < 1 or clx_data[0] != 2:
            return ""

        pos = 1
        n = len(clx_data)

        # salta il padding di 4 byte (inutilizzato) e legge lcbPlcPcd
        if pos + 8 > n:
            return ""
        lcb_piece_table = struct.unpack_from("<I", clx_data, pos + 4)[0]
        if lcb_piece_table <= 0:
            return ""

        plc_start = pos + 8
        plc_end = plc_start + lcb_piece_table
        if plc_end > n:
            plc_end = n
        plc = clx_data[plc_start:plc_end]

        # Il PLC e' composto da: aCps (n+1 valori da 4 byte) + aPcd (n valori 8 byte)
        n_pieces = (len(plc) - 4) // 12
        if n_pieces <= 0:
            return ""

        cps = []
        for i in range(n_pieces + 1):
            cp = struct.unpack_from("<i", plc, 4 + i * 4)[0]
            cps.append(cp)

        # I Pcd iniziano dopo aCps
        pcd_base = 4 + (n_pieces + 1) * 4

        parts = []
        for i in range(n_pieces):
            cp_start = cps[i]
            cp_end = cps[i + 1]
            if cp_end <= cp_start:
                continue
            pcd = plc[pcd_base + i * 8:pcd_base + i * 8 + 8]
            if len(pcd) < 8:
                break

            f_compressed = bool(pcd[0] & 0x01)
            fc_raw = struct.unpack_from("<I", pcd, 2)[0]

            if f_compressed:
                # ANSI: 1 byte per carattere, fc e' l'offset (gia' diviso per 2)
                fc_text = fc_raw >> 1
                char_count = cp_end - cp_start
                start = fc_text
                end = fc_text + char_count
                if start < len(word_stream) and end <= len(word_stream):
                    raw = word_stream[start:end]
                    try:
                        parts.append(raw.decode("cp1252", errors="replace"))
                    except Exception:
                        parts.append(raw.decode("latin-1", errors="replace"))
            else:
                # Unicode: 2 byte per carattere, fc e' l'offset in byte (diviso 2)
                fc_text = fc_raw >> 1
                char_count = cp_end - cp_start
                start = fc_text
                end = fc_text + char_count * 2
                if start < len(word_stream) and end <= len(word_stream):
                    raw = word_stream[start:end]
                    try:
                        parts.append(raw.decode("utf-16-le", errors="replace"))
                    except Exception:
                        pass

        text = "".join(parts)

        # Se dopo la Piece Table abbiamo comunque testo, alcune piattaforme
        # (esempio doc generati da discendenti Word) mettono il testo anche
        # dentro fcMin/fcMac: non facciamo nulla qui, torniamo cio' che c'e'.
        return text.strip()

    def _fallback_extract_doc_text_raw(self, file_path):
        """Ultima spiaggia: legge il WordDocument stream e prova più codifiche
        su tutto il flusso, filtrando il rumore binario. Poco affidabile ma
        meglio di niente quando mancano sia catdoc/LibreOffice sia la Piece
        Table riconoscibile."""
        try:
            import olefile
            ole = olefile.OleFileIO(file_path)
            try:
                if not ole.exists("WordDocument"):
                    return ""
                ws = ole.openstream("WordDocument").read()
            finally:
                ole.close()
        except Exception:
            try:
                with open(file_path, "rb") as f:
                    ws = f.read()
            except Exception:
                return ""

        # prova blocchi UTF-16LE leggibili in tutto lo stream
        utf16_pattern = re.compile(rb"(?:[\x20-\x7e\xa0-\xff]\x00){8,}")
        chunks = []
        for m in utf16_pattern.finditer(ws):
            try:
                decoded = m.group(0).decode("utf-16-le", errors="ignore").strip()
                if decoded:
                    alpha = sum(1 for c in decoded if c.isalpha())
                    if alpha > len(decoded) * 0.3:
                        chunks.append(decoded)
            except Exception:
                continue
        if chunks:
            return "\n".join(chunks).strip()

        # altrimenti prova a decodificare l'intero file come ASCII/CP1252
        for enc in ("cp1252", "latin-1", "cp437"):
            try:
                decoded = ws.decode(enc, errors="replace")
                alpha = sum(1 for c in decoded if c.isalpha())
                if alpha > len(decoded) * 0.2:
                    return decoded.strip()
            except Exception:
                continue
        return ""

    """
    Funzione sostitutiva per _extract_text_rtf in simplewright.py
    Copia-incolla questa funzione sopra a quella vecchia (linea 1442-1501)
    """

    def _extract_text_rtf_with_formatting(self, file_path):
        """Estrae testo e formattazione RTF come run con offset carattere."""
        with open(file_path, "rb") as f:
            raw=f.read()
        data=raw.decode("cp1252", errors="replace")

        # Tabelle font e colori definite nell'header RTF.
        font_table={}
        for m in re.finditer(r"\{\\f(\d+)[^;{}]*\s*([^;{}]+);", data, re.I):
            name=m.group(2).strip()
            if name:
                font_table[int(m.group(1))]=name
        color_table=["#000000"]
        cm=re.search(r"\{\\colortbl\b(.*?)\}", data, re.I|re.S)
        if cm:
            color_table=[]
            for entry in cm.group(1).split(";"):
                r=re.search(r"\\red(\d+)",entry,re.I); g=re.search(r"\\green(\d+)",entry,re.I); b=re.search(r"\\blue(\d+)",entry,re.I)
                if r or g or b:
                    color_table.append(f"#{int(r.group(1)) if r else 0:02X}{int(g.group(1)) if g else 0:02X}{int(b.group(1)) if b else 0:02X}")
                else:
                    color_table.append("#000000")
            if not color_table: color_table=["#000000"]

        fmt={"font":font_table.get(0,self.font_family.get() or "Arial"),"size":max(1,int(self.font_size.get())),"bold":False,"italic":False,"underline":False,"color":"#000000","bg":None,"align":"left"}
        stack=[]; text_chars=[]; skip_depths=set(); depth=0; i=0
        # Pattern.match(data, posizione) evita di copiare data[i:] a ogni
        # controllo: quelle slice rendevano quadratica l'estrazione dei file grandi.
        # RTF ignorable destinations use {\\*\\name ...}. The previous
        # expression expected {\\*name and missed Word's XML/object payloads.
        group_start_re = re.compile(r"\{(?P<ignorable>\\\*)?(?P<name>\\[a-zA-Z]+)")
        hex_escape_re = re.compile(r"\\'([0-9a-fA-F]{2})")
        unicode_escape_re = re.compile(r"\\u(-?\d+)\??")
        control_re = re.compile(r"\\([a-zA-Z]+)(-?\d+)? ?")
        skip_dest={"fonttbl","colortbl","stylesheet","info","pict","object","objdata","header","headerl","headerr","footer","footerl","footerr","footnote","annotation","comment","revtbl","listtable","listoverridetable","generator","xmlnstbl","themedata","colorschememapping","latentstyles","shp","nonshpict","datafield","fldinst","fldrslt"}
        def add(ch):
            if ch: text_chars.append((ch,fmt.copy()))

        while i<len(data):
            ch=data[i]
            if ch=='{':
                depth+=1; stack.append(fmt.copy())
                m=group_start_re.match(data,i)
                if m:
                    destination=m.group("name")[1:].lower()
                    if m.group("ignorable") or destination in skip_dest:
                        skip_depths.add(depth)
                i+=1; continue
            if ch=='}':
                skip_depths.discard(depth)
                if stack: fmt=stack.pop()
                depth=max(0,depth-1); i+=1; continue
            if skip_depths:
                i+=1; continue
            if ch=='\\':
                if i+1<len(data) and data[i+1] in r'\\{}':
                    add(data[i+1]); i+=2; continue
                m=hex_escape_re.match(data,i)
                if m:
                    add(bytes.fromhex(m.group(1)).decode("cp1252",errors="replace")); i=m.end(); continue
                m=unicode_escape_re.match(data,i)
                if m:
                    n=int(m.group(1)); n=n+65536 if n<0 else n; add(chr(n)); i=m.end(); continue
                m=control_re.match(data,i)
                if m:
                    w=m.group(1).lower(); n=int(m.group(2)) if m.group(2) is not None else None
                    if w=='b': fmt['bold']=n!=0
                    elif w=='i': fmt['italic']=n!=0
                    elif w in ('ul','ulw'): fmt['underline']=True
                    elif w in ('ul0','ulnone'): fmt['underline']=False
                    elif w=='fs' and n is not None: fmt['size']=max(1,n/2.0)
                    elif w=='f' and n is not None: fmt['font']=font_table.get(n,fmt['font'])
                    elif w=='cf' and n is not None: fmt['color']=color_table[n] if 0<=n<len(color_table) else '#000000'
                    elif w in ('highlight','chcbpat') and n is not None: fmt['bg']=color_table[n] if 0<=n<len(color_table) else None
                    elif w=='plain':
                        fmt.update({"font":font_table.get(0,self.font_family.get() or "Arial"),"size":max(1,int(self.font_size.get())),"bold":False,"italic":False,"underline":False,"color":"#000000","bg":None})
                    elif w=='ql': fmt['align']='left'
                    elif w=='qc': fmt['align']='center'
                    elif w=='qr': fmt['align']='right'
                    elif w=='qj': fmt['align']='justify'
                    elif w in ('par','line','softline','sect','page'): add('\n')
                    elif w=='tab': add('\t')
                    elif w in ('enspace','qmspace'): add(' ')
                    elif w=='emspace': add('  ')
                    elif w=='emdash': add('—')
                    elif w=='endash': add('–')
                    elif w=='bullet': add('•')
                    elif w=='lquote': add('‘')
                    elif w=='rquote': add('’')
                    elif w=='ldblquote': add('“')
                    elif w=='rdblquote': add('”')
                    # \\binN è seguito da N byte binari, che non fanno parte del testo.
                    elif w=='bin' and n is not None and n>0:
                        i=min(len(data),m.end()+n)
                        continue
                    i=m.end(); continue
                i+=1; continue
            if ch in '\r\n': i+=1; continue
            add(ch); i+=1

        text=''.join(ch for ch,_ in text_chars)
        if not text: return '',[]
        text=text.replace('\r\n','\n').replace('\r','\n').strip('\n')
        runs=[]; rs=0; prev=None
        for idx,(_,f) in enumerate(text_chars):
            key=(f.get('font'),round(float(f.get('size',12)),2),bool(f.get('bold')),bool(f.get('italic')),bool(f.get('underline')),f.get('color'),f.get('bg'),f.get('align'))
            if prev is None: prev=key
            elif key!=prev:
                if idx>rs: runs.append((rs,idx,text_chars[rs][1].copy()))
                rs=idx; prev=key
        if rs<len(text_chars): runs.append((rs,len(text_chars),text_chars[rs][1].copy()))
        runs=[(s,min(e,len(text)),f) for s,e,f in runs if s<len(text) and e>s]
        return text,runs

    def _apply_rtf_tags(self, rtf_runs):
        """Applica le run RTF agli editor, anche quando il testo e' paginato."""
        if not rtf_runs: return
        # Le pagine sono segmenti consecutivi del testo estratto, separati da
        # un newline. Calcolare gli offset cumulativamente evita ricerche
        # ripetute sul testo (e gestisce correttamente anche pagine duplicate).
        page_ranges = []
        cursor = 0
        for pid in sorted(self.pages_data):
            page_text = self.pages_data.get(pid, '')
            page_end = cursor + len(page_text)
            page_ranges.append((pid, cursor, page_end))
            cursor = page_end + 1

        # Le run dell'estrattore sono ordinate per offset. Avanziamo un cursore
        # tra le pagine invece di ricontrollare tutte le run per ogni pagina:
        # su RTF lunghi questo elimina il costo pagine × run.
        ordered_runs = sorted(rtf_runs, key=lambda run: run[0])
        run_index = 0
        for pid,ps,pe in page_ranges:
            ed=self.text_page_editors.get(pid)
            if ed is None: continue
            ed.tag_configure('rtf_underline',underline=True)
            while run_index < len(ordered_runs) and ordered_runs[run_index][1] <= ps:
                run_index += 1
            i = run_index
            while i < len(ordered_runs) and ordered_runs[i][0] < pe:
                a,b,f = ordered_runs[i]
                la=max(a,ps)-ps; lb=min(b,pe)-ps
                try: sidx=ed.index('1.0 + %dc'%la); eidx=ed.index('1.0 + %dc'%lb)
                except Exception:
                    i += 1
                    continue
                try:
                    size=max(1,int(round(float(f.get('size',12)))))
                    ft=self._get_font_tag(ed,f.get('font') or self.font_family.get(),size,bool(f.get('bold')),bool(f.get('italic')))
                    ed.tag_add(ft,sidx,eidx)
                except Exception: pass
                color=f.get('color')
                if color and color.upper()!='#000000':
                    tag='rtf_color_'+color.lstrip('#').upper(); ed.tag_configure(tag,foreground=color); ed.tag_add(tag,sidx,eidx)
                if f.get('bg'):
                    tag='rtf_bg_'+f['bg'].lstrip('#').upper(); ed.tag_configure(tag,background=f['bg']); ed.tag_add(tag,sidx,eidx)
                if f.get('underline'):
                    ed.tag_add('rtf_underline',sidx,eidx)
                align=f.get('align','left')
                if align in ('left','center','right'):
                    tag='rtf_'+align; ed.tag_configure(tag,justify=align); ed.tag_add(tag,sidx,eidx)
                i += 1

    def _apply_auto_links(self, editor=None):
        """Evidenzia gli URL http(s) e li apre con Ctrl+click.

        I tag vengono aggiornati sugli editor attivi senza modificare il testo,
        così selezione, undo e formattazione RTF restano gestiti da Tk Text.
        """
        import webbrowser

        url_pattern = re.compile(r"https?://[^\s<>\"']+", re.IGNORECASE)
        editors = [(None, editor)] if editor is not None else list(self.text_page_editors.items())
        trailing_punctuation = ".,;:!?)]}"

        for page_id, ed in editors:
            if ed is None:
                continue
            try:
                ed.tag_configure("auto_hyperlink", foreground="#0000EE", underline=True)
                # Rimuove solo i tag generati da questa funzione, non quelli RTF.
                for tag_name in ed.tag_names():
                    if tag_name.startswith("auto_hyperlink_"):
                        ed.tag_delete(tag_name)
                ed.tag_remove("auto_hyperlink", "1.0", "end")

                content = ed.get("1.0", "end-1c")
                for match in url_pattern.finditer(content):
                    url = match.group(0).rstrip(trailing_punctuation)
                    if not url:
                        continue
                    start_offset = match.start()
                    end_offset = start_offset + len(url)
                    start = ed.index("1.0 + %dc" % start_offset)
                    end = ed.index("1.0 + %dc" % end_offset)
                    tag_name = "auto_hyperlink_%d" % start_offset
                    ed.tag_configure(tag_name)
                    ed.tag_add("auto_hyperlink", start, end)
                    ed.tag_add(tag_name, start, end)
                    ed.tag_bind(
                        tag_name, "<Control-Button-1>",
                        lambda event, target=url: (webbrowser.open(target), "break")[1],
                    )
            except Exception:
                # Un editor può essere distrutto durante un aggiornamento UI.
                continue

    def _schedule_auto_links(self, page_id=None):
        """Aggiorna i link dopo che Tk ha completato la modifica del testo."""
        try:
            editors = ([self.text_page_editors.get(page_id)] if page_id is not None
                       else list(self.text_page_editors.values()))
            for ed in editors:
                if ed is not None:
                    ed.after_idle(lambda widget=ed: self._apply_auto_links(widget))
        except Exception:
            pass

    def _apply_markdown_styles(self):
        """Applica una formattazione visiva basilare in stile Markdown (titoli
        piu' grandi, grassetto/corsivo/codice inline, separatori, elenchi
        puntati) alle pagine appena caricate da un file .md, invece di
        mostrare i simboli '#', '**', ecc. come semplice testo."""
        fam = self.font_family.get()
        base_size = max(8, int(self.font_size.get()))
        heading_sizes = {1: base_size + 12, 2: base_size + 8, 3: base_size + 4}
        inline_re = re.compile(r"(\*\*.+?\*\*|`[^`]+`|\*[^*\n]+\*|(?<![\w])_[^_\n]+_(?![\w]))")

        for pid, ed in self.text_page_editors.items():
            try:
                ed.tag_configure("md_h1", font=(fam, heading_sizes[1], "bold"))
                ed.tag_configure("md_h2", font=(fam, heading_sizes[2], "bold"))
                ed.tag_configure("md_h3", font=(fam, heading_sizes[3], "bold"))
                ed.tag_configure("md_bold", font=(fam, base_size, "bold"))
                ed.tag_configure("md_italic", font=(fam, base_size, "italic"))
                ed.tag_configure("md_code", font=("Courier New", base_size))
                ed.tag_configure("md_hr", foreground="#999999")

                line_count = int(ed.index("end-1c").split(".")[0])
                for ln in range(1, line_count + 1):
                    line_text = ed.get(f"{ln}.0", f"{ln}.end")
                    stripped = line_text.strip()

                    m = re.match(r"^(#{1,3})\s+(.*)$", stripped)
                    if m:
                        level = len(m.group(1))
                        ed.delete(f"{ln}.0", f"{ln}.end")
                        ed.insert(f"{ln}.0", m.group(2))
                        ed.tag_add(f"md_h{level}", f"{ln}.0", f"{ln}.end")
                        continue

                    if re.match(r"^(-{3,}|\*{3,}|_{3,})$", stripped):
                        ed.delete(f"{ln}.0", f"{ln}.end")
                        ed.insert(f"{ln}.0", "―" * 40)
                        ed.tag_add("md_hr", f"{ln}.0", f"{ln}.end")
                        continue

                    working_text = line_text
                    if re.match(r"^[-*]\s+", stripped):
                        working_text = "•  " + re.sub(r"^[-*]\s+", "", line_text)

                    if inline_re.search(working_text):
                        ed.delete(f"{ln}.0", f"{ln}.end")
                        pos = 0
                        for mm in inline_re.finditer(working_text):
                            if mm.start() > pos:
                                ed.insert(f"{ln}.end", working_text[pos:mm.start()])
                            chunk = mm.group(0)
                            if chunk.startswith("**"):
                                seg, tag = chunk[2:-2], "md_bold"
                            elif chunk.startswith("`"):
                                seg, tag = chunk[1:-1], "md_code"
                            else:
                                seg, tag = chunk[1:-1], "md_italic"
                            start_idx = ed.index(f"{ln}.end")
                            ed.insert(f"{ln}.end", seg)
                            ed.tag_add(tag, start_idx, f"{ln}.end")
                            pos = mm.end()
                        if pos < len(working_text):
                            ed.insert(f"{ln}.end", working_text[pos:])
                    elif working_text != line_text:
                        ed.delete(f"{ln}.0", f"{ln}.end")
                        ed.insert(f"{ln}.0", working_text)
            except Exception:
                continue

    def open_file(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("Tutti i formati supportati", "*.txt *.docx *.doc *.pptx *.rtf *.epub *.odt *.ods *.md *.pdf"),
            ("Documenti PDF e Vettoriali", "*.pdf *.pptx *.epub *.ods"),
            ("Documenti Testuali", "*.docx *.doc *.odt *.rtf *.txt *.md"),
            ("Tutti i file", "*.*")
        ])
        if not file_path:
            return "break"

        ext = os.path.splitext(file_path)[1].lower()

        # Libera l'eventuale PDF temporaneo generato dalla conversione del .doc
        # precedente, per non accumulare file orfani nella cartella temporanea.
        if getattr(self, "_doc_pdf_cleanup", None):
            try:
                self._doc_pdf_cleanup()
            except Exception:
                pass
            self._doc_pdf_cleanup = None

        # Orientamento di default Verticale; viene sovrascritto a "Orizzontale"
        # per i formati intrinsecamente orizzontali (presentazioni .pptx e
        # e-book .epub). Cosi' l'impostazione in alto rispecchia sempre il foglio.
        self.current_orientation = "Verticale"
        try:
            self.orientation_cb.set("Verticale")
        except Exception:
            pass

        # Pulizia dell'area di lavoro
        self.default_font_size = max(1, int(self.font_size.get()))
        for pid in list(self.page_text_box_widgets.keys()):
            self._clear_text_box_widgets(pid)
        self._clear_all_images()
        self.pages_data.clear()
        self.page_text_boxes = {1: []}
        self.page_text_format = {1: []}
        self.page_text_box_widgets = {1: []}
        self.page_images = {1: []}
        self.page_image_widgets = {1: []}
        self.page_display_names = {}
        self.embedded_images_cache.clear()
        self.pdf_pages_info.clear()
        self.pdf_photo_refs.clear()
        self.pdf_entry_widgets.clear()
        self.pdf_page_offsets = {}
        self.active_pdf_box = None

        # FIX CRITICO: distrugge gli editor/frame di pagina del documento
        # precedente PRIMA di caricare il nuovo file. Senza questo,
        # update_page_layout() chiama save_current_page_state() che rilegge
        # il testo dai vecchi widget Tkinter (ancora vivi in memoria) e lo
        # reinietta sopra il contenuto del nuovo file appena aperto: e' la
        # causa della "pagina 2 fantasma", del contenuto sbagliato/mescolato
        # e del disallineamento del testo (frame vecchi con dimensioni diverse
        # che restano a galleggiare sul canvas).
        self.canvas_desktop.delete("text_page_win")
        for pid, frame in list(self.text_page_frames.items()):
            try:
                frame.destroy()
            except Exception:
                pass
        self.text_page_frames = {}
        self.text_page_editors = {}

        try:
            # 1. MODALITA' VISUALE (solo PDF nativo, tramite PyMuPDF)
            if ext == ".pdf":
                self.load_pdf(file_path)
                self.switch_to_pdf_view()
                self.current_file = file_path
                self.update_pages_sidebar()
                self.render_all_pdf_pages()
                self.root.update_idletasks()
                self.pdf_canvas.yview_moveto(0.0)
                self.active_page_id = 1
                self.root.title(f"{file_path} - UniversalWriter [Layout Originale]")

            # 2. PRESENTAZIONI (.pptx) - resa come le pagine PDF (immagine di sfondo
            # + caselle di testo modificabili con font/dimensione reali), invece che
            # come semplice testo su pagine verticali A4: le slide non sono mai
            # verticali, quindi si usa il layout "a pagina fissa" pensato per il PDF,
            # ma alimentato dai dati reali del file .pptx (forma e dimensioni della
            # slide, immagini incluse, font dei singoli blocchi di testo).
            elif ext == ".pptx":
                # Percorso primario: converte la presentazione in PDF con
                # LibreOffice e la mostra tramite la pipeline PDF di PyMuPDF,
                # che rende in modo fedele e stabile ogni elemento (testo,
                # immagini, tabelle, gruppi). Questo e' il percorso che apriva
                # i pptx mostrando anche le immagini.
                pdf_path, cleanup = self._convert_pptx_to_pdf(file_path)
                if pdf_path:
                    self._doc_pdf_cleanup = cleanup
                    self.load_pdf(pdf_path)
                    self.switch_to_pdf_view()
                    self.current_file = file_path
                    # Le slide sono intrinsecamente orizzontali: allinea
                    # l'impostazione di orientamento in alto al foglio reale.
                    self.current_orientation = "Orizzontale"
                    try:
                        self.orientation_cb.set("Orizzontale")
                    except Exception:
                        pass
                    self.update_pages_sidebar()
                    self.render_all_pdf_pages()
                    self.root.update_idletasks()
                    self.pdf_canvas.yview_moveto(0.0)
                    self.active_page_id = 1
                    self.root.title(f"{file_path} - UniversalWriter [Presentazione]")
                else:
                    # Fallback se LibreOffice manca: parser python-pptx.
                    self.load_pptx(file_path)
                    self.switch_to_pdf_view()
                    self.current_file = file_path
                    self.current_orientation = "Orizzontale"
                    try:
                        self.orientation_cb.set("Orizzontale")
                    except Exception:
                        pass
                    self.update_pages_sidebar()
                    self.render_all_pdf_pages()
                    self.root.update_idletasks()
                    self.pdf_canvas.yview_moveto(0.0)
                    self.active_page_id = 1
                    self.root.title(f"{file_path} - UniversalWriter [Presentazione]")

            # 3. EPUB - testo dei capitoli, impaginato come i .txt (35 righe a pagina)
            elif ext == ".epub":
                self.switch_to_text_view()
                # Un epub ha in genere paragrafi lunghi: mettendo la pagina in
                # orizzontale si guadagna spazio utile in larghezza per la lettura.
                self.current_orientation = "Orizzontale"
                try:
                    self.orientation_cb.set("Orizzontale")
                except Exception:
                    pass
                text_content = self._extract_text_epub(file_path).lstrip('\r\n')
                self._load_text_content_into_pages(text_content)
                self.current_file = file_path
                self.update_page_layout()
                self.root.update_idletasks()
                self.canvas_desktop.yview_moveto(0.0)
                self.switch_to_page(1)
                self.root.title(f"{file_path} - UniversalWriter")

            # 4. FOGLI DI CALCOLO (.ods) - testo delle celle, impaginato come i .txt
            elif ext == ".ods":
                self.switch_to_text_view()
                text_content = self._extract_text_opendocument(file_path).lstrip('\r\n')
                self._load_text_content_into_pages(text_content)
                self.current_file = file_path
                self.update_page_layout()
                self.root.update_idletasks()
                self.canvas_desktop.yview_moveto(0.0)
                self.switch_to_page(1)
                self.root.title(f"{file_path} - UniversalWriter")

            # 4-bis. DOC legacy (Word 97-2003) - reso tramite la pipeline PDF:
            # LibreOffice espande i campi Word (TOC/PAGEREF/HYPERLINK) e restituisce
            # il testo pulito, mentre le immagini originali diventano parte dello
            # sfondo della pagina (come per PDF e PPTX). Se LibreOffice manca,
            # si ricade sulla modalita' testuale (estrazione testo puro).
            elif ext == ".doc":
                pdf_path, cleanup = self._convert_doc_to_pdf(file_path)
                if pdf_path:
                    self._doc_pdf_cleanup = cleanup
                    self.load_pdf(pdf_path)
                    self.switch_to_pdf_view()
                    self.current_file = file_path
                    self.update_pages_sidebar()
                    self.render_all_pdf_pages()
                    self.root.update_idletasks()
                    self.pdf_canvas.yview_moveto(0.0)
                    self.active_page_id = 1
                    self.root.title(f"{file_path} - UniversalWriter [Layout Originale]")
                else:
                    text_content = self._extract_text_doc_legacy(file_path)
                    if not text_content.strip():
                        for enc in ("utf-8", "cp1252", "latin-1"):
                            try:
                                with open(file_path, "r", encoding=enc, errors="ignore") as f:
                                    text_content = f.read()
                                break
                            except Exception:
                                continue
                    doc_fam, doc_size = self._dominant_doc_font(file_path)
                    if doc_fam:
                        resolved = self._resolve_pptx_font_family(doc_fam)
                        self.font_family.set(resolved)
                    if doc_size:
                        self.font_size.set(doc_size)
                    text_content = text_content.lstrip('\r\n')
                    self._load_text_content_into_pages(text_content)
                    self.current_file = file_path
                    self.update_page_layout()
                    self.root.update_idletasks()
                    self.canvas_desktop.yview_moveto(0.0)
                    self.switch_to_page(1)
                    self.root.title(f"{file_path} - UniversalWriter")


            # 5. MODALITA' TESTUALE (TXT, MD, RTF, DOCX, ODT) - Aperti come testo puro editabile
            elif ext in [".docx", ".odt", ".txt", ".md", ".rtf"]:
                self.switch_to_text_view()
                text_content = ""

                if ext == ".docx":
                    doc = Document(file_path)
                    text_content = "\n".join([p.text for p in doc.paragraphs])
                    # Mantiene font e dimensione del documento originale invece di
                    # forzare sempre il font/size di default dell'editor.
                    fam, size = self._dominant_docx_font(doc)
                    if fam:
                        resolved = self._resolve_pptx_font_family(fam)
                        self.font_family.set(resolved)
                    if size:
                        self.font_size.set(size)
                elif ext == ".odt":
                    text_content = self._extract_text_opendocument(file_path)
                elif ext == ".rtf":
                    text_content, rtf_tags = self._extract_text_rtf_with_formatting(file_path)
                    # Applica i tag RTF dopo aver caricato il testo
                    self._rtf_tags_pending = rtf_tags
                else:
                    # Prova piu' codifiche prima di arrendersi a "ignore", che
                    # altrimenti puo' scartare silenziosamente gran parte del
                    # testo se il file non e' realmente in UTF-8 (es. txt salvati
                    # da Windows in cp1252/latin-1), facendo sembrare il file "vuoto".
                    text_content = None
                    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
                        try:
                            with open(file_path, "r", encoding=enc) as f:
                                text_content = f.read()
                            break
                        except (UnicodeDecodeError, UnicodeError):
                            continue
                    if text_content is None:
                        with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                            text_content = f.read()

                text_content = text_content.lstrip('\r\n')
                self._load_text_content_into_pages(text_content)

                self.current_file = file_path
                self.update_page_layout()
                if ext == ".md":
                    # Rende visivamente i simboli Markdown piu' comuni (titoli piu'
                    # grandi, grassetto/corsivo/inline code, separatori) invece di
                    # mostrare i simboli grezzi come testo semplice.
                    self._apply_markdown_styles()
                # Applica tag RTF se presenti
                try:
                    if hasattr(self, '_rtf_tags_pending') and self._rtf_tags_pending:
                        self._apply_rtf_tags(self._rtf_tags_pending)
                        self._rtf_tags_pending = None
                except Exception:
                    pass
                # I widget di testo esistono dopo il rendering/impaginazione.
                self._apply_auto_links()
                # Forza esplicitamente lo scroll in cima, senza fidarsi di una
                # scrollregion/posizione residua del documento aperto in precedenza
                # (causa del bug per cui il documento si apriva "a pagina 2").
                self.root.update_idletasks()
                self.canvas_desktop.yview_moveto(0.0)
                self.switch_to_page(1)
                self.root.title(f"{file_path} - UniversalWriter")
                
        except Exception as e:
            messagebox.showerror("Errore di Lettura", f"Impossibile aprire il file in modo nativo:\n{str(e)}")
            
        return "break"

    def _load_text_content_into_pages(self, text_content):
        """Suddivide un testo lungo in pagine da ~35 righe l'una, popolando self.pages_data."""
        self.default_font_size = max(1, int(self.font_size.get()))
        self._clear_all_images()
        lines = text_content.split('\n')
        page_idx = 1
        current_page_text = ""
        self.pages_data = {}
        self.page_text_format = {}
        for line in lines:
            current_page_text += line + "\n"
            if len(current_page_text.split('\n')) > 35:
                self.pages_data[page_idx] = current_page_text.strip()
                page_idx += 1
                current_page_text = ""
        if current_page_text.strip() or page_idx == 1:
            self.pages_data[page_idx] = current_page_text.strip()
        self.page_text_format = {k: [] for k in self.pages_data}
        self.page_images = {k: [] for k in self.pages_data}
        self.page_image_widgets = {k: [] for k in self.pages_data}
        if not self.pages_data:
            self.pages_data = {1: ""}
            self.page_text_format = {1: []}
            self.page_images = {1: []}

    def _auto_paginate_text_pages(self):
        """Ridistribuisce il testo delle pagine in modo che NESSUNA pagina superi
        l'altezza utile del foglio. Misura l'altezza reale del testo usando le
        metriche del font effettivo (stesso stile di _create_rich_text_box):
        ogni riga che non entra nel foglio viene automaticamente spostata in
        una nuova pagina successiva, senza dipendere dalla geometry di un
        widget."""
        if self.pdf_mode or not self.pages_data:
            return
        base_w, base_h = self.base_page_sizes[self.current_format]
        w, h = int(base_w * self.zoom_factor), int(base_h * self.zoom_factor)
        if self.current_orientation == "Orizzontale":
            w, h = h, w
        try:
            m_left = int(self.scale_m_left.get())
            m_right = int(self.scale_m_right.get())
            m_top = int(self.scale_m_top.get())
            m_bottom = int(self.scale_m_bottom.get())
        except Exception:
            m_left = m_right = m_top = m_bottom = 40
        avail_w = max(40, w - m_left - m_right)
        avail_h = max(40, h - m_top - m_bottom)
        try:
            fam = self.font_family.get()
            fsize = max(1, int(self.default_font_size * self.zoom_factor))
        except Exception:
            fam, fsize = "Arial", 12
        try:
            f = font.Font(family=fam, size=fsize)
            line_h = max(10, int(f.metrics("linespace")))
        except Exception:
            line_h = int(fsize * 1.2)
        max_lines = max(1, int(avail_h // line_h))

        new_pages = []
        new_formats = []
        old_page_start = {}
        for p_id in sorted(self.pages_data.keys()):
            old_page_start[p_id] = len(new_pages)
            content = self.pages_data[p_id]
            fmt = self.page_text_format.get(p_id, [])
            if not content.strip():
                new_pages.append(content)
                new_formats.append(fmt)
                continue
            # Stima il wrapping parola per parola. Misurare una riga intera e
            # assegnarla alla pagina successiva quando supera l'altezza faceva
            # sì che un singolo paragrafo EPUB molto lungo rimanesse indiviso.
            lines = content.split("\n")
            page_lines = []
            used = 0
            page_chars = 0
            for line_no, ln in enumerate(lines):
                # Una riga vuota occupa comunque una riga verticale.
                words = re.findall(r"\S+\s*", ln)
                if not words:
                    chunks = [""]
                else:
                    chunks = []
                    current = ""
                    for word in words:
                        candidate = current + word
                        try:
                            fits = not current or f.measure(candidate) <= avail_w
                        except Exception:
                            fits = True
                        if fits:
                            current = candidate
                        else:
                            chunks.append(current.rstrip())
                            current = word
                            # Spezza anche parole singole più larghe della
                            # pagina, evitando righe che non possono andare a capo.
                            while current and f.measure(current) > avail_w:
                                cut = max(1, len(current) // 2)
                                while cut > 1 and f.measure(current[:cut]) > avail_w:
                                    cut -= 1
                                chunks.append(current[:cut])
                                current = current[cut:]
                    if current or not chunks:
                        chunks.append(current.rstrip())

                for chunk_no, chunk in enumerate(chunks):
                    if used >= max_lines and page_lines:
                        left, right = self._split_format_at(fmt, page_chars)
                        new_pages.append("\n".join(page_lines))
                        new_formats.append(left)
                        fmt = right
                        page_lines = []
                        used = 0
                        page_chars = 0
                    page_lines.append(chunk)
                    used += 1
                    page_chars += len(chunk)
                    # I ritorni a capo originari contano come un carattere
                    # nell'indice dei tag di formattazione.
                    if chunk_no == len(chunks) - 1 and line_no < len(lines) - 1:
                        page_chars += 1
            new_pages.append("\n".join(page_lines))
            new_formats.append(fmt)
        self.pages_data = {i + 1: new_pages[i] for i in range(len(new_pages))}
        self.page_text_format = {i + 1: new_formats[i] for i in range(len(new_formats))}
        # riallinea le immagini: ogni vecchia pagina mappa alla prima nuova
        # pagina creata dal suo contenuto.
        reordered_images = {i + 1: [] for i in range(len(new_pages))}
        for old_id, idx in old_page_start.items():
            new_pid = idx + 1
            if new_pid in reordered_images:
                reordered_images[new_pid].extend(self.page_images.get(old_id, []))
        self.page_images = reordered_images
        self.page_image_widgets = {k: [] for k in self.pages_data.keys()}

    def switch_to_pdf_view(self):
        self.pdf_mode = True
        self.pdf_add_mode = False
        self.canvas_desktop.pack_forget()
        self.pdf_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar_y.config(command=self.pdf_canvas.yview)
        self.scrollbar_x.config(command=self.pdf_canvas.xview)
        self.pdf_canvas.configure(yscrollcommand=self.scrollbar_y.set)
        self.pdf_canvas.configure(xscrollcommand=self.scrollbar_x.set)

    def switch_to_text_view(self):
        self.pdf_mode = False
        self.pdf_add_mode = False
        self.pdf_canvas.pack_forget()
        self.canvas_desktop.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar_y.config(command=self.canvas_desktop.yview)
        self.scrollbar_x.config(command=self.canvas_desktop.xview)
        self.canvas_desktop.configure(yscrollcommand=self.scrollbar_y.set)
        self.canvas_desktop.configure(xscrollcommand=self.scrollbar_x.set)
        self.update_page_layout()

    def load_pdf(self, file_path):
        doc = pymupdf.open(file_path)
        # Ogni pagina viene ricondotta al foglio A4 standard dell'editor
        # (595x842 punti, cioe' 595x842 px a zoom 1.0, come la modalita' testo),
        # ridimensionando il contenuto con un "fit contain" per farlo stare
        # interamente dentro il foglio senza distorcerlo. Cosi' un PDF, un ebook
        # o un .doc convertito appaiono sempre alla stessa dimensione di un
        # normale foglio A4, invece di diventare "giganti" a schermo.
        ref_w, ref_h = self.base_page_sizes["A4"]  # 595 x 842 punti/pixel

        for idx, page in enumerate(doc):
            page_id = idx + 1
            pw = page.rect.width or ref_w
            ph = page.rect.height or ref_h
            # Fit "contain": scala per far stare la pagina inalterata dentro A4.
            scale_fit = min(ref_w / pw, ref_h / ph)
            scale_fit = max(0.05, scale_fit)
            # DPI in modo che il pixelato corrisponda a quota foglio A4 standard.
            page_dpi = int(round(72.0 * scale_fit))
            page_dpi = max(20, min(page_dpi, 400))  # PyMuPDF vuole un int

            text_dict = page.get_text("dict")
            spans = []

            for bi, block in enumerate(text_dict.get("blocks", [])):
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        raw_text = span.get("text", "")
                        if not raw_text.strip():
                            continue
                        b = span["bbox"]
                        spans.append({
                            "text": raw_text,
                            "bbox": b,
                            "width": b[2] - b[0],
                            "height": b[3] - b[1],
                            "origin": span["origin"],
                            "size": span["size"],
                            "font": span["font"],
                            "color": span["color"],
                            "flags": span["flags"],
                            "blk": bi,
                        })
                        rect = pymupdf.Rect(b)
                        page.add_redact_annot(rect, fill=(1, 1, 1))

            page.apply_redactions()
            pix = page.get_pixmap(dpi=page_dpi)
            img_bytes = pix.tobytes("png")

            self.pdf_pages_info[page_id] = {
                "spans": spans,
                "boxes": self._group_spans_into_boxes(spans),
                "img_bytes": img_bytes,
                "width_pt": page.rect.width,
                "height_pt": page.rect.height,
                "render_dpi": page_dpi,
                "overlays": [],
            }
        doc.close()
        self.active_page_id = 1

    def _group_spans_into_boxes(self, spans):
        """Raggruppa gli span PDF vicini in caselle di testo per PARAGRAFO.

        L'obiettivo e' unire in un'unica casella il testo che sta insieme
        (stessa riga -> stesso blocco/paragrafo), cosi' le parole con colori o
        grandezze diverse dentro lo stesso blocco non vengono piu' separate in
        caselle che si accavallano. Ogni casella ritorna:
          - "lines":  lista di righe, ogni riga = lista di run ordinati,
          - "runs":   lista appiattita di tutti i run (per export/editing),
          - "bbox":   rettangolo complessivo in punti [x0,y0,x1,y1],
          - "x_pt"/"y_pt": origine in punti, "width_pt"/"height_pt" in punti,
          - campi di formattazione "base" per la toolbar.

        Si sfrutta il raggruppamento in "blocco" (paragrafo) gia' calcolato da
        PyMuPDF quando e' disponibile; altrimenti si usano le regole empiriche
        su spaziature/titoli/indentazione.
        """
        if not spans:
            return []

        # --- 1) Dividi gli span in paragrafi ---
        have_blk = all("blk" in s for s in spans)
        paragraphs = []
        if have_blk:
            by_blk = {}
            order = []
            for s in spans:
                if s["blk"] not in by_blk:
                    by_blk[s["blk"]] = []
                    order.append(s["blk"])
                by_blk[s["blk"]].append(s)
            paragraphs = [by_blk[k] for k in order]
        else:
            srt = sorted(spans, key=lambda s: (round(s["bbox"][1], 1), s["bbox"][0]))
            needles = self._split_paragraph_lines(srt)
            paragraphs = []
            cur = []
            prev = None
            for ln in needles:
                if prev is not None and not self._lines_same_paragraph(prev, ln, needles):
                    paragraphs.append(cur)
                    cur = []
                cur.append(ln)
                prev = ln
            if cur:
                paragraphs.append(cur)
            paragraphs = [[s for ln in pg for s in ln["spans"]] for pg in paragraphs]

        # --- 2) Per ogni paragrafo: dividi in righe e costruisci la casella ---
        boxes = []
        for pspans in paragraphs:
            pspans = sorted(pspans, key=lambda s: (round(s["bbox"][1], 1), s["bbox"][0]))
            plines = []
            for s in pspans:
                y0 = s["bbox"][1]
                placed = False
                for ln in plines:
                    if abs(ln["y_anchor"] - y0) <= 3.0:
                        ln["spans"].append(s)
                        placed = True
                        break
                if not placed:
                    plines.append({"y_anchor": y0, "spans": [s]})
            for ln in plines:
                ln["spans"].sort(key=lambda s: s["bbox"][0])
                ln["y_top"] = min(s["bbox"][1] for s in ln["spans"])
                ln["y_bottom"] = max(s["bbox"][3] for s in ln["spans"])
                ln["x_left"] = min(s["bbox"][0] for s in ln["spans"])
                ln["x_right"] = max(s["bbox"][2] for s in ln["spans"])

            runs = []
            pline_runs = []
            for ln in plines:
                line_runs = []
                for s in ln["spans"]:
                    family, weight, slant = self._font_details_from_span(s)
                    rd = {"text": s["text"], "size": s["size"], "color": s["color"],
                          "family": family, "weight": weight, "slant": slant, "span": s}
                    line_runs.append(rd)
                    runs.append(rd)
                pline_runs.append(line_runs)

            x0 = min(ln["x_left"] for ln in plines)
            y0 = min(ln["y_top"] for ln in plines)
            x1 = max(ln["x_right"] for ln in plines)
            y1 = max(ln["y_bottom"] for ln in plines)
            sizes = [r["size"] for r in runs]
            boxes.append({
                "runs": runs,
                "lines": pline_runs,
                "bbox": [x0, y0, x1, y1],
                "x_pt": x0, "y_pt": y0,
                "width_pt": x1 - x0,
                "height_pt": y1 - y0,
                "custom_family": runs[0]["family"] if runs else "Arial",
                "custom_size": max(1, int(round(max(sizes)))) if sizes else 12,
                "custom_weight": runs[0]["weight"] if runs else "normal",
                "custom_slant": runs[0]["slant"] if runs else "roman",
                "custom_color": runs[0]["color"] if runs else 0x000000,
            })
        return boxes

    def _split_paragraph_lines(self, spans):
        """Raggruppa gli span in righe (stessa baseline). Usato solo come
        fallback delle slide/pptx quando non c'e' il blocco di PyMuPDF."""
        lines = []
        for s in spans:
            y0 = s["bbox"][1]
            placed = False
            for ln in lines:
                if abs(ln["y_anchor"] - y0) <= 3.0:
                    ln["spans"].append(s)
                    placed = True
                    break
            if not placed:
                lines.append({"y_anchor": y0, "spans": [s]})
        for ln in lines:
            ln["spans"].sort(key=lambda s: s["bbox"][0])
            ln["y_top"] = min(s["bbox"][1] for s in ln["spans"])
            ln["y_bottom"] = max(s["bbox"][3] for s in ln["spans"])
            ln["x_left"] = min(s["bbox"][0] for s in ln["spans"])
            ln["x_right"] = max(s["bbox"][2] for s in ln["spans"])
        return lines

    def _lines_same_paragraph(self, a, b, all_lines):
        """Decide se la riga b continua il paragrafo della riga a (fallback)."""
        gap = b["y_top"] - a["y_bottom"]
        if gap <= 0:
            return True
        # grandezza carattere tipica delle righe vicine
        sizes_a = [s["size"] for s in a["spans"]]
        sizes_b = [s["size"] for s in b["spans"]]
        sz_a = max(sizes_a) if sizes_a else 10
        sz_b = max(sizes_b) if sizes_b else 10
        # cambio di grandezza/titolo -> nuovo paragrafo
        if abs(sz_a - sz_b) > 2.0:
            return False
        # interlinea normale ~1.2-1.35x la grandezza
        if gap > sz_a * 1.6:
            return False
        # indentazione diversa (> 5pt) -> nuovo paragrafo/voce
        if abs(a["x_left"] - b["x_left"]) > 5.0:
            return False
        return True

    def _resolve_pptx_font_family(self, raw_name):
        """Prova a mantenere il font originale della slide; se non e' disponibile
        sul sistema ricade su una delle 3 famiglie generiche (come per i PDF)."""
        if raw_name:
            if raw_name in self.all_fonts:
                return raw_name
            for f in self.all_fonts:
                if f.lower() == raw_name.lower():
                    return f
        base = (raw_name or "").lower()
        if "times" in base or "georgia" in base or "serif" in base or "cambria" in base:
            fallback = "Times New Roman"
        elif "courier" in base or "mono" in base or "consolas" in base:
            fallback = "Courier New"
        else:
            fallback = "Arial"
        if fallback not in self.all_fonts:
            fallback = "Arial" if "Arial" in self.all_fonts else self.all_fonts[0]
        return fallback

    @staticmethod
    def _pptx_shape_xfrm_raw(shape):
        """Legge left/top/width/height 'grezzi' della forma (EMU), cosi' come
        scritti nell'XML. Per le forme dentro un gruppo questi valori sono
        espressi nel sistema di coordinate locale del gruppo (chOff/chExt) e
        NON in quello della slide: vanno convertiti con _pptx_group_child_transform."""
        try:
            return (getattr(shape, "left", 0) or 0,
                    getattr(shape, "top", 0) or 0,
                    getattr(shape, "width", 0) or 0,
                    getattr(shape, "height", 0) or 0)
        except Exception:
            return 0, 0, 0, 0

    @staticmethod
    def _pptx_group_child_transform(shape, scale_x, scale_y, trans_x, trans_y):
        """Calcola la trasformazione (scala + traslazione) da applicare ai figli
        di un GroupShape.

        python-pptx NON applica automaticamente l'offset/scala dei gruppi: le
        forme dentro un <p:grpSp> hanno coordinate espresse in un sistema
        locale definito da chOff/chExt, diverso da quello (off/ext) del
        gruppo stesso nella slide. Se non si converte questa trasformazione,
        immagini e testo dentro i gruppi finiscono con coordinate quasi
        sempre sbagliate (spesso fuori dall'area visibile della slide),
        che e' il motivo principale per cui in molte presentazioni le
        immagini/il testo raggruppati risultavano "invisibili"."""
        from pptx.oxml.ns import qn
        try:
            grp_pr = shape._element.grpSpPr
            xfrm = grp_pr.find(qn('a:xfrm'))
            ch_off = xfrm.find(qn('a:chOff'))
            ch_ext = xfrm.find(qn('a:chExt'))
            ch_off_x, ch_off_y = int(ch_off.get('x')), int(ch_off.get('y'))
            ch_ext_cx, ch_ext_cy = int(ch_ext.get('cx')), int(ch_ext.get('cy'))
        except Exception:
            return scale_x, scale_y, trans_x, trans_y

        g_left, g_top, g_w, g_h = AdvancedTextEditor._pptx_shape_xfrm_raw(shape)
        final_left = trans_x + g_left * scale_x
        final_top = trans_y + g_top * scale_y
        final_w = g_w * scale_x
        final_h = g_h * scale_y

        new_scale_x = (final_w / ch_ext_cx) if ch_ext_cx else scale_x
        new_scale_y = (final_h / ch_ext_cy) if ch_ext_cy else scale_y
        new_trans_x = final_left - ch_off_x * new_scale_x
        new_trans_y = final_top - ch_off_y * new_scale_y
        return new_scale_x, new_scale_y, new_trans_x, new_trans_y

    @classmethod
    def _pptx_walk_shapes(cls, shapes, scale_x=1.0, scale_y=1.0, trans_x=0.0, trans_y=0.0):
        """Genera ricorsivamente tutte le forme 'foglia' (non-gruppo) della
        slide, con le coordinate EMU gia' convertite nel sistema assoluto
        della slide, gestendo correttamente i gruppi annidati."""
        for sh in shapes:
            l, t, w, h = cls._pptx_shape_xfrm_raw(sh)
            final_left = trans_x + l * scale_x
            final_top = trans_y + t * scale_y
            final_w = w * scale_x
            final_h = h * scale_y

            if getattr(sh, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
                try:
                    gsx, gsy, gtx, gty = cls._pptx_group_child_transform(
                        sh, scale_x, scale_y, trans_x, trans_y)
                    yield from cls._pptx_walk_shapes(sh.shapes, gsx, gsy, gtx, gty)
                except Exception:
                    pass
                continue

            rotation = getattr(sh, "rotation", 0) or 0
            yield sh, final_left, final_top, final_w, final_h, rotation

    @staticmethod
    def _pptx_shape_fill_image_blob(shape):
        """Estrae l'immagine usata come riempimento ('picture fill') di una
        forma qualunque (es. un rettangolo usato come foto a tutto sfondo in
        un layout). python-pptx classifica queste forme come AUTO_SHAPE:
        prima venivano completamente ignorate perche' solo le forme di tipo
        PICTURE producevano un'immagine."""
        from pptx.oxml.ns import qn
        try:
            sp_pr = shape._element.spPr
            blip_fill = sp_pr.find(qn('a:blipFill'))
            if blip_fill is None:
                return None
            blip = blip_fill.find(qn('a:blip'))
            if blip is None:
                return None
            r_id = blip.get(qn('r:embed')) or blip.get(qn('r:link'))
            if not r_id:
                return None
            return shape.part.related_part(r_id).image.blob
        except Exception:
            return None

    @staticmethod
    def _pptx_slide_background_image(slide):
        """Estrae, se presente, l'immagine impostata come sfondo dell'intera
        slide (fill dello sfondo = immagine), prima ignorata del tutto."""
        from pptx.oxml.ns import qn
        try:
            bg = slide.background
            if bg.fill.type != 6:  # MSO_FILL_TYPE.PICTURE == 6
                return None
            bg_pr = bg._element.find(qn('p:bgPr'))
            blip_fill = bg_pr.find(qn('a:blipFill'))
            blip = blip_fill.find(qn('a:blip'))
            r_id = blip.get(qn('r:embed'))
            return slide.part.related_part(r_id).image.blob
        except Exception:
            return None

    def load_pptx(self, file_path):
        """Carica un file .pptx riusando la stessa pipeline di visualizzazione
        del PDF (self.pdf_pages_info): ogni slide diventa una pagina con
        un'immagine di sfondo (dove vengono incollate le immagini reali della
        slide) e delle caselle di testo modificabili posizionate secondo il
        layout originale, con font e dimensione dei singoli blocchi di testo
        preservati il piu' possibile."""
        if not HAS_PPTX:
            raise RuntimeError("La libreria 'python-pptx' non è installata (pip install python-pptx).")

        prs = Presentation(file_path)
        self.pdf_pages_info = {}

        # 1 EMU = 1/914400 pollici; 1 punto = 1/72 pollice -> pt = emu / 12700
        EMU_PER_PT = 12700.0
        slide_w_pt = (prs.slide_width or 9144000) / EMU_PER_PT
        slide_h_pt = (prs.slide_height or 6858000) / EMU_PER_PT

        target_max_px = 1400
        page_max_pt = max(slide_w_pt, slide_h_pt, 1)
        page_dpi = (target_max_px / page_max_pt) * 72.0
        page_dpi = int(round(max(60, min(page_dpi, 220))))
        scale_px = page_dpi / 72.0

        px_w = max(1, int(round(slide_w_pt * scale_px)))
        px_h = max(1, int(round(slide_h_pt * scale_px)))

        def emu_to_pt(v):
            return (v or 0) / EMU_PER_PT

        def _open_image_safe(blob):
            try:
                pic = Image.open(io.BytesIO(blob))
                pic.load()
                return pic.convert("RGBA")
            except Exception:
                # Formati non rasterizzabili da Pillow (es. immagini vettoriali
                # WMF/EMF incollate da Word/PowerPoint): niente da fare, si salta.
                return None

        def _paste_scaled(bg, pic, left_pt, top_pt, w_pt, h_pt, rotation=0):
            box_w = max(1, int(round(w_pt * scale_px)))
            box_h = max(1, int(round(h_pt * scale_px)))
            try:
                pic = pic.resize((box_w, box_h), Image.Resampling.LANCZOS)
                px, py = int(round(left_pt * scale_px)), int(round(top_pt * scale_px))
                if rotation:
                    pic = pic.rotate(-rotation, expand=True, resample=Image.Resampling.BICUBIC)
                    cx, cy = px + box_w / 2.0, py + box_h / 2.0
                    px = int(round(cx - pic.width / 2.0))
                    py = int(round(cy - pic.height / 2.0))
                bg.paste(pic, (px, py), pic)
            except Exception:
                pass

        for s_idx, slide in enumerate(prs.slides, start=1):
            bg = Image.new("RGB", (px_w, px_h), "white")

            # Sfondo dedicato della slide (se impostato come immagine),
            # prima non gestito affatto.
            bg_blob = self._pptx_slide_background_image(slide)
            if bg_blob:
                bg_pic = _open_image_safe(bg_blob)
                if bg_pic:
                    _paste_scaled(bg, bg_pic, 0, 0, slide_w_pt, slide_h_pt)

            spans = []

            # Scorre TUTTE le forme (anche dentro gruppi annidati, con le
            # coordinate correttamente convertite) e per ciascuna:
            # - se e' un'immagine (o una forma con riempimento a immagine),
            #   la incolla sullo sfondo nella posizione/rotazione corrette;
            # - se ha del testo, crea una casella di testo modificabile.
            # Cosi' sia le immagini sia il testo dentro i gruppi (prima
            # completamente ignorato) vengono recuperati.
            for shape, left_emu, top_emu, width_emu, height_emu, rotation in self._pptx_walk_shapes(slide.shapes):
                left, top = emu_to_pt(left_emu), emu_to_pt(top_emu)
                width = emu_to_pt(width_emu) or 50
                height = emu_to_pt(height_emu) or 20

                img_blob = None
                if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
                    try:
                        img_blob = shape.image.blob
                    except Exception:
                        img_blob = None
                else:
                    # Forme (rettangoli, placeholder, ecc.) riempite con
                    # un'immagine invece che di un colore pieno.
                    img_blob = self._pptx_shape_fill_image_blob(shape)

                if img_blob:
                    pic = _open_image_safe(img_blob)
                    if pic:
                        _paste_scaled(bg, pic, left, top, width or 50, height or 20, rotation)
                    # Una forma-immagine non ha mai testo modificabile associato.
                    continue

                if not getattr(shape, "has_text_frame", False):
                    continue

                paragraphs = shape.text_frame.paragraphs
                line_top = top
                for para in paragraphs:
                    runs = list(para.runs)
                    para_text = "".join(r.text for r in runs) if runs else para.text
                    if not para_text.strip():
                        # riga vuota: avanza comunque per non sovrapporre le righe successive
                        line_top += max(10, (para.font.size.pt if para.font and para.font.size else 18) * 1.2)
                        continue

                    rep_run = runs[0] if runs else None
                    rep_font = rep_run.font if rep_run is not None else para.font

                    size_pt = None
                    if rep_font is not None and rep_font.size is not None:
                        try:
                            size_pt = rep_font.size.pt
                        except Exception:
                            size_pt = None
                    if size_pt is None:
                        size_pt = 28 if s_idx == 1 and line_top == top else 18
                    size_pt = max(6, min(size_pt, 96))

                    raw_font_name = rep_font.name if (rep_font is not None and rep_font.name) else None
                    family = self._resolve_pptx_font_family(raw_font_name)
                    bold = bool(rep_font.bold) if rep_font is not None and rep_font.bold is not None else False
                    italic = bool(rep_font.italic) if rep_font is not None and rep_font.italic is not None else False

                    color_int = 0x000000
                    try:
                        if rep_font is not None and rep_font.color and rep_font.color.type is not None:
                            rgb = rep_font.color.rgb
                            color_int = int(str(rgb), 16)
                    except Exception:
                        color_int = 0x000000

                    line_h = size_pt * 1.25
                    bbox = [left, line_top, left + width, line_top + line_h]

                    span = {
                        "text": para_text,
                        "bbox": bbox,
                        "width": width,
                        "height": line_h,
                        "origin": (left, line_top + line_h),
                        "size": size_pt,
                        "font": raw_font_name or family,
                        "color": color_int,
                        "flags": (16 if bold else 0) | (2 if italic else 0),
                        # Precalcolati per mantenere fedelmente font/dimensione reali
                        # della slide invece della mappatura generica dei PDF.
                        "custom_font": family,
                        "custom_size": int(round(size_pt)),
                        "custom_weight": "bold" if bold else "normal",
                        "custom_slant": "italic" if italic else "roman",
                    }
                    spans.append(span)
                    line_top += line_h

            b_io = io.BytesIO()
            bg.save(b_io, format="PNG")
            img_bytes = b_io.getvalue()

            self.pdf_pages_info[s_idx] = {
                "spans": spans,
                "boxes": self._group_spans_into_boxes(spans),
                "img_bytes": img_bytes,
                "width_pt": slide_w_pt,
                "height_pt": slide_h_pt,
                "render_dpi": page_dpi,
                "overlays": [],
                "display_name": f"Slide {s_idx}",
            }

        if not self.pdf_pages_info:
            raise RuntimeError("Nessuna slide leggibile trovata nella presentazione.")

        self.active_page_id = 1

    def _font_details_from_span(self, span):
        flags = span.get("flags", 0)
        bold = bool(flags & 16)
        italic = bool(flags & 2)
        raw_name = (span.get("font") or "Helvetica")
        base = raw_name.split("+")[-1].split(",")[0].split("-")[0].lower()
        if "times" in base or "georgia" in base or "serif" in base:
            family = "Times New Roman"
        elif "courier" in base or "mono" in base or "consolas" in base:
            family = "Courier New"
        else:
            family = "Arial"
        if family not in self.all_fonts:
            family = "Arial" if "Arial" in self.all_fonts else self.all_fonts[0]
        weight = "bold" if bold else "normal"
        slant = "italic" if italic else "roman"
        return family, weight, slant

    def _int_color_to_hex(self, color_int):
        r = (color_int >> 16) & 255
        g = (color_int >> 8) & 255
        b = color_int & 255
        return "#%02x%02x%02x" % (r, g, b)

    def _int_color_to_rgb01(self, color_int):
        r = (color_int >> 16) & 255
        g = (color_int >> 8) & 255
        b = color_int & 255
        return (r / 255.0, g / 255.0, b / 255.0)

    def render_all_pdf_pages(self):
        info_check = self.pdf_pages_info
        if not info_check:
            return
        self.pdf_canvas.delete("all")
        self.pdf_entry_widgets = {}
        self.active_pdf_box = None
        self.pdf_page_offsets = {}

        gap = 24
        z = self.zoom_factor
        y_cursor = 0
        max_w = 0

        # Prima passata: dimensioni (gia' zoomate) per allineare e centrare tutte
        # le pagine sullo stesso bordo sinistro.
        zoomed_dims = {}
        for page_id in sorted(self.pdf_pages_info.keys()):
            info = self.pdf_pages_info[page_id]
            img = Image.open(io.BytesIO(info["img_bytes"]))
            zoomed_dims[page_id] = (max(1, int(round(img.width * z))),
                                    max(1, int(round(img.height * z))))

        canvas_width = max(self.pdf_canvas.winfo_width(), 1)
        all_max_w = max((d[0] for d in zoomed_dims.values()), default=1)
        # Centra il blocco di pagine quando entra nel canvas, altrimenti lascia
        # un piccolo margine sinistro (e la barra orizzontale permette lo scroll).
        x_off = max(20, (canvas_width - all_max_w) // 2)

        for page_id in sorted(self.pdf_pages_info.keys()):
            info = self.pdf_pages_info[page_id]
            scale = info.get("render_dpi", self.pdf_render_dpi) / 72.0
            img = Image.open(io.BytesIO(info["img_bytes"]))
            if z != 1.0:
                img = img.resize(zoomed_dims[page_id][0:2], Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self.pdf_photo_refs[page_id] = photo

            page_w, page_h = img.width, img.height
            max_w = max(max_w, page_w)

            self.pdf_page_offsets[page_id] = y_cursor
            self.pdf_canvas.create_image(x_off, y_cursor, anchor="nw", image=photo)

            self.pdf_canvas.create_line(x_off, y_cursor, x_off + page_w, y_cursor, fill="#999999")
            self.pdf_canvas.create_line(x_off, y_cursor, x_off, y_cursor + page_h, fill="#999999")
            self.pdf_canvas.create_line(x_off + page_w, y_cursor, x_off + page_w, y_cursor + page_h, fill="#999999")
            self.pdf_canvas.create_line(x_off, y_cursor + page_h, x_off + page_w, y_cursor + page_h, fill="#999999", dash=(6, 4))

            self.pdf_canvas.create_text(
                x_off + 8, y_cursor + 6, anchor="nw",
                text=info.get("display_name", f"Pagina {page_id}"),
                fill="#666666", font=("Arial", 8, "bold")
            )

            self.pdf_entry_widgets[page_id] = []

            for box in info.get("boxes", []):
                box["_x_pt"] = box.get("_x_pt", box.get("x_pt", 0))
                box["_y_pt"] = box.get("_y_pt", box.get("y_pt", 0))
                x0, y0 = box["_x_pt"], box["_y_pt"]
                cx = x_off + x0 * scale * z
                cy = y_cursor + y0 * scale * z

                container, entry, grip, win_id = self._create_rich_text_box(
                    self.pdf_canvas, cx, cy, box, scale * z,
                    page_id, (scale * z, x_off, y_cursor, page_w, page_h)
                )
                container._win_id = win_id
                self.pdf_entry_widgets[page_id].append((win_id, container, entry, box))

            for ov in info["overlays"]:
                self._create_overlay_entry(page_id, ov, scale, y_offset=y_cursor, x_offset=x_off, z=z, page_w=page_w)

            y_cursor += page_h + gap

        total_h = max(y_cursor, 100)
        total_w = x_off + all_max_w + 40
        self.pdf_canvas.config(scrollregion=(0, 0, total_w, total_h))

        if self.active_page_id in self.pdf_page_offsets:
            self.jump_to_pdf_page(self.active_page_id)


    def jump_to_pdf_page(self, page_id):
        if page_id not in self.pdf_page_offsets:
            return
        self.root.update_idletasks()
        region = self.pdf_canvas.cget("scrollregion")
        try:
            parts = [float(v) for v in str(region).split()]
            total_h = parts[3] if len(parts) == 4 and parts[3] > 0 else 1.0
        except Exception:
            total_h = 1.0
        y = self.pdf_page_offsets[page_id]
        frac = y / max(1.0, total_h)
        self.pdf_canvas.yview_moveto(max(0.0, min(1.0, frac)))

    def _create_resizable_box(self, parent_canvas, x, y, w, h, text, family, pt_size, weight, slant, color_hex, data_obj, page_id=1, page_meta=None):
        """Crea una casella di testo editabile "invisibile" sopra il foglio.
        - Il riquadro/bordo sono nascosti di default: resta visibile solo il testo.
        - Larghezza minima = larghezza del testo col font/dimensione usati.
        - Quando selezionata (click sul testo) compaiono: una X rossa in alto a
          destra (elimina la casella), una maniglia in alto a sinistra (trascina
          per spostare) e una maniglia in basso a destra (ridimensiona)."""
        scale_z, x_off, y_off, page_w = (page_meta or (1.0, 0, 0, 595))

        box_font = font.Font(family=family, size=pt_size, weight=weight, slant=slant)
        text_w = box_font.measure(text) if text else 0

        # Riduci automaticamente il font se la riga di testo non entra nel foglio:
        # scala la grandezza del file originale con il nostro A4 standard, cosi'
        # il testo non esce mai fuori dal bordo del foglio.
        page_right = x_off + page_w
        available_w = max(20, page_right - x)
        # Itero fino a che la riga (testo + padding) entra nel foglio: riduco
        # il font scalando la grandezza del file originale con il nostro A4.
        for _ in range(4):
            if text_w + 6 <= available_w:
                break
            # Piccolo margine extra per evitare che l'arrotondamento "half-to-even"
            # riporti il font alla grandezza di partenza (es. round(9.5)=10).
            reduce_by = (available_w / max(1, text_w + 6)) * 0.97
            new_pt = max(6, int(pt_size * reduce_by))
            if new_pt >= pt_size:
                break
            pt_size = new_pt
            box_font = font.Font(family=family, size=pt_size, weight=weight, slant=slant)
            text_w = box_font.measure(text) if text else 0

        # Larghezza minima = larghezza del testo alla dimensione/font correnti,
        # cosi' una casella e' lunga quanto il testo che contiene (e non usa mai
        # la larghezza del bbox originale del PDF, che puo' sforare il foglio).
        pad_x = 6
        pad_y = 4
        line_h = max(10, int(round(pt_size * 1.25)))
        fit_w = max(text_w, 10) + pad_x
        fit_h = line_h + pad_y
        box_w = max(int(fit_w), 60)
        box_h = max(int(h), fit_h)

        # Se il testo (ridotto) non entra comunque, tronca la casella al bordo
        # destro del foglio: cosi' nessuna casella sfora mai il foglio a nessuno
        # zoom (si preserva la proporzione testo/foglio dell'A4).
        limit_w = int(page_right - x)
        if box_w > limit_w and limit_w >= 20:
            box_w = limit_w

        data_obj["_reduced_size"] = pt_size
        data_obj["custom_size"] = pt_size

        container = tk.Frame(parent_canvas, bg="#ffffff", bd=0, highlightthickness=0,
                             width=box_w, height=box_h)
        container.pack_propagate(False)

        entry = tk.Entry(container, bd=0, highlightthickness=0, bg="#ffffff",
                         fg=color_hex, font=box_font, insertbackground=color_hex)
        entry.insert(0, text)
        entry.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=pad_x // 2, pady=pad_y // 2)
        entry.box_font = box_font
        entry.data_obj = data_obj
        entry.container = container
        entry.page_id = page_id
        entry.is_selected = False

        # --- Chrome (nascosto di default) ---
        btn_close = tk.Label(parent_canvas, text="✕", bg="#d32f2f", fg="#ffffff",
                             cursor="hand2", font=("Arial", 8, "bold"))
        self._add_tooltip(btn_close, "Elimina casella di testo")
        btn_close.bind("<Button-1>", lambda e, pid=page_id, d=data_obj: self._delete_pdf_box(pid, d))

        drag_lbl = tk.Label(parent_canvas, text="✥", bg="#90caf9", fg="#0d47a1",
                            cursor="fleur", font=("Arial", 8))
        self._add_tooltip(drag_lbl, "Trascina per spostare la casella")

        grip = tk.Label(parent_canvas, text="◢", bg="#90caf9", fg="#0d47a1",
                        cursor="sizing", font=("Arial", 7))
        self._add_tooltip(grip, "Trascina per ridimensionare")

        container._chrome = (btn_close, drag_lbl, grip)
        entry._chrome = (btn_close, drag_lbl, grip)
        self._set_chrome_visible(entry, False)
        entry._meta = (scale_z, x_off, y_off)

        # --- Selezione / deselezione ---
        def on_entry_click(_e, en=entry):
            self.select_pdf_box(en)
        entry.bind("<Button-1>", on_entry_click)
        entry.bind("<FocusIn>", lambda e, en=entry: self.select_pdf_box(en))
        entry.bind("<KeyRelease>", lambda e, en=entry: (en.data_obj.__setitem__("text", en.get()),
                                                        self._autosize_box(en)))

        # --- Trascinamento (maniglia in alto a sinistra) ---
        def start_drag(event):
            container._drag_x0 = event.x_root
            container._drag_y0 = event.y_root
            container._drag_cx0, container._drag_cy0 = parent_canvas.coords(container._win_id)

        def do_drag(event):
            dx = event.x_root - container._drag_x0
            dy = event.y_root - container._drag_y0
            nx = container._drag_cx0 + dx
            ny = container._drag_cy0 + dy
            parent_canvas.coords(container._win_id, nx, ny)
            self._position_text_box_chrome(entry)
            # salva la nuova posizione in punti per il ri-render
            sz, xo, yo = entry._meta
            data_obj["_x_pt"] = (nx - xo) / sz
            data_obj["_y_pt"] = (ny - yo) / sz
            if data_obj.get("_moved") is False:
                data_obj["_moved"] = True

        drag_lbl.bind("<Button-1>", start_drag)
        drag_lbl.bind("<B1-Motion>", do_drag)

        # --- Ridimensionamento (maniglia in basso a destra) ---
        def start_resize(event):
            container._gx0 = event.x_root
            container._gy0 = event.y_root
            container._gw0 = container.winfo_width()
            container._gh0 = container.winfo_height()

        def do_resize(event):
            new_w = max(max(10, text_w) + pad_x, container._gw0 + (event.x_root - container._gx0))
            new_h = max(fit_h, container._gh0 + (event.y_root - container._gy0))
            container.config(width=int(new_w), height=int(new_h))
            parent_canvas.itemconfigure(container._win_id, width=int(new_w), height=int(new_h))
            self._position_text_box_chrome(entry)
            sz, xo, yo = entry._meta
            if data_obj.get("_moved"):
                data_obj["width_pt"] = new_w / sz
                data_obj["height_pt"] = new_h / sz
            else:
                data_obj["width_pt"] = new_w / sz
                data_obj["height_pt"] = new_h / sz

        grip.bind("<Button-1>", start_resize)
        grip.bind("<B1-Motion>", do_resize)

        return container, entry, grip

    def _create_rich_text_box(self, parent_canvas, x, y, box, scale_z, page_id=1, page_meta=None):
        """Crea una casella di testo RICCA (tk.Text) per una casella-paragrafo.

        Dentro la stessa casella vengono resi piu' "run" (parole/parti) con
        colori, grandezze e grassetto diversi (formattazione mista), cosi' il
        testo vicino non viene spezzato in caselle che si accavallano. I tag
        tk.Text restano attaccati ai caratteri anche quando si modifica il testo.
        """
        scale_z, x_off, y_off, page_w, page_h = (page_meta or (scale_z, 0, 0, 595, 842))
        available_w = max(60, int((x_off + page_w) - x))
        available_h = max(60, int((y_off + page_h) - y))

        # --- Calcola, per ogni riga, i run con font scalato per far entrare la
        # riga nel foglio (la "grandezza del file originale" viene ridotta alla
        # nostra A4). Cosi' il testo non esce/resta nascosto e non e' sproporzionato.
        # ogni voce: (list[(rd_ref, px)], ln_h)
        sized_lines = []
        for line in box["lines"]:
            if not line:
                sized_lines.append(([], 12))
                continue
            rf = []
            for rd in line:
                # riduzione globale (scala la grandezza originale con la nostra A4)
                rf.append((rd, max(6, int(round(rd["size"] * scale_z * self._FONT_SCALE)))))
            # riduce il font della riga finche' non entra
            for _ in range(4):
                w = 0
                for _rd, px in rf:
                    f = font.Font(family=_rd["family"], size=px, weight=_rd["weight"], slant=_rd["slant"])
                    w += f.measure(_rd["text"]) if _rd["text"] else 0
                if w <= available_w - 6:
                    break
                factor = ((available_w - 6) / max(1, w)) * 0.97
                nxt = [max(4, int(px * factor)) for _rd, px in rf]
                if all(n == px for (_rd, px), n in zip(rf, nxt)):
                    break
                rf = [(_rd, n) for (_rd, px), n in zip(rf, nxt)]
            hmax = max((px for _rd, px in rf), default=10)
            ln_h = max(12, int(round(hmax * 1.25)))
            sized_lines.append((rf, ln_h))

        line_widths = []
        total_h = 0
        for rf, ln_h in sized_lines:
            w = 0
            for _rd, px in rf:
                f = font.Font(family=_rd["family"], size=px, weight=_rd["weight"], slant=_rd["slant"])
                w += f.measure(_rd["text"]) if _rd["text"] else 0
            line_widths.append(max(w, 20))
            total_h += ln_h

        pad_x = 8
        pad_y = 6

        # --- Riduce il font in altezza se il paragrafo non entra nel foglio:
        # cosi' anche in basso il testo viene preso dentro (come a destra).
        # Iterativo, come il fit orizzontale: si restringe finche' non entra
        # oppure non si riesce piu' a ridurre. ---
        if available_h >= 60:
            for _ in range(5):
                if total_h + pad_y <= available_h:
                    break
                hf = ((available_h - pad_y) / max(1, total_h)) * 0.97
                new_total = 0
                new_sized = []
                for rf, ln_h in sized_lines:
                    new_px = [max(4, int(px * hf)) for _rd, px in rf]
                    new_h = max(8, int(round(max(new_px, default=10) * 1.25)))
                    new_sized.append(([(rd, p) for (rd, _o), p in zip(rf, new_px)], new_h))
                    new_total += new_h
                if new_total >= total_h:
                    break
                sized_lines = new_sized
                total_h = new_total
                line_widths = []
                for rf, ln_h in sized_lines:
                    w = 0
                    for _rd, px in rf:
                        f = font.Font(family=_rd["family"], size=px, weight=_rd["weight"], slant=_rd["slant"])
                        w += f.measure(_rd["text"]) if _rd["text"] else 0
                    line_widths.append(max(w, 20))

        content_w = (max(line_widths) if line_widths else 40)
        box_w = content_w + pad_x
        limit_w = int((x_off + page_w) - x)
        if limit_w >= 30 and box_w > limit_w:
            box_w = limit_w
        box_h = max(total_h + pad_y, 24)
        box_w = max(box_w, int(box.get("min_w", 0)))
        box_h = max(box_h, int(box.get("min_h", 0)))

        container = tk.Frame(parent_canvas, bg="#ffffff", bd=0, highlightthickness=0,
                             width=box_w, height=box_h)
        container.pack_propagate(False)

        textw = tk.Text(container, wrap="none", bd=0, highlightthickness=0,
                        bg="#ffffff", insertbackground="#000000",
                        spacing1=0, spacing3=0, padx=2, pady=2, cursor="xterm")
        textw.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # --- Inserisce testo con i tag (run) usando i font ridotti ---
        textw.tagstash = {}
        textw._tag_to_run = {}
        for li, (rf, ln_h) in enumerate(sized_lines):
            if li > 0:
                textw.insert("end", "\n")
            for ri, (rd, px) in enumerate(rf):
                tag = f"rr{li}_{ri}"
                textw.insert("end", rd["text"], (tag,))
                textw.tagstash[tag] = (px, rd["color"], rd["family"], rd["weight"], rd["slant"])
                textw._tag_to_run[tag] = rd
        for tag, fmt in textw.tagstash.items():
            px, color_int, family, weight, slant = fmt
            f = font.Font(family=family, size=px, weight=weight, slant=slant)
            textw.tag_configure(tag, font=f, foreground=self._int_color_to_hex(color_int),
                                underline=bool(textw._tag_to_run[tag].get("underline", False)))
        # run vuote (es. casella creata a mano): posiziona il range-tag in alto
        # a sinistra cosi' la prima digitazione eredita la formattazione del tag.
        for _t in textw.tagstash:
            try:
                if not textw.tag_ranges(_t):
                    textw.tag_add(_t, "1.0", "1.0")
            except Exception:
                pass
        textw.configure(state="disabled")
        textw.configure(state="normal")

        textw.data_obj = box
        textw.container = container
        textw.page_id = page_id
        textw.is_selected = False
        textw._rich = True
        textw._meta = (scale_z, x_off, y_off)
        textw._box_w0, textw._box_h0 = box_w, box_h
        textw._min_w = content_w + pad_x
        textw._min_h = box_h


        # --- Chrome (nascosto di default) ---
        btn_close = tk.Label(parent_canvas, text="✕", bg="#d32f2f", fg="#ffffff",
                             cursor="hand2", font=("Arial", 8, "bold"))
        self._add_tooltip(btn_close, "Elimina casella di testo")
        btn_close.bind("<Button-1>", lambda e, pid=page_id, d=box: self._delete_pdf_box(pid, d))

        drag_lbl = tk.Label(parent_canvas, text="✥", bg="#90caf9", fg="#0d47a1",
                            cursor="fleur", font=("Arial", 8))
        self._add_tooltip(drag_lbl, "Trascina per spostare la casella")

        grip = tk.Label(parent_canvas, text="◢", bg="#90caf9", fg="#0d47a1",
                        cursor="sizing", font=("Arial", 7))
        self._add_tooltip(grip, "Trascina per ridimensionare")

        # I controlli della casella ricca devono essere window item del canvas
        # della pagina, così scorrono insieme alla casella e restano cliccabili.
        for control in (btn_close, drag_lbl, grip):
            control._canvas_win_id = parent_canvas.create_window(
                x, y, anchor="nw", window=control, state="hidden"
            )

        container._chrome = (btn_close, drag_lbl, grip)
        textw._chrome = (btn_close, drag_lbl, grip)
        textw._win_id = None
        self._set_chrome_visible(textw, False)
        _bg = box.get("background")
        if _bg:
            container.configure(bg=_bg)
            textw.configure(bg=_bg)

        # --- Selezione / deselezione + aggiornamento toolbar dal clic ---
        def on_click(_e, t=textw):
            self.select_pdf_box(t)
            self._sync_toolbar_from_rich(t)
        textw.bind("<Button-1>", on_click)
        textw.bind("<FocusIn>", on_click)
        textw.bind("<ButtonRelease-1>", lambda _e, t=textw: self._update_saved_format_selection(t), add="+")

        # --- Modifica: salva il testo nel box e mantiene i tag ---
        def on_key(_e, t=textw):
            self._flush_rich_text(t)
        textw.bind("<KeyRelease>", on_key)
        textw.bind("<KeyRelease>", lambda _e, t=textw: self._update_saved_format_selection(t), add="+")

        # --- Trascinamento (maniglia in alto a sinistra) ---
        def start_drag(event):
            container._drag_x0 = event.x_root
            container._drag_y0 = event.y_root
            container._drag_cx0, container._drag_cy0 = parent_canvas.coords(container._win_id)

        def do_drag(event):
            dx = event.x_root - container._drag_x0
            dy = event.y_root - container._drag_y0
            nx = container._drag_cx0 + dx
            ny = container._drag_cy0 + dy
            parent_canvas.coords(container._win_id, nx, ny)
            sz, xo, yo = textw._meta
            box["_x_pt"] = (nx - xo) / sz
            box["_y_pt"] = (ny - yo) / sz
            self._position_text_box_chrome(textw)

        drag_lbl.bind("<Button-1>", start_drag)
        drag_lbl.bind("<B1-Motion>", do_drag)

        # --- Ridimensionamento (maniglia in basso a destra) ---
        def start_resize(event):
            container._gx0 = event.x_root
            container._gy0 = event.y_root
            container._gw0 = container.winfo_width()
            container._gh0 = container.winfo_height()

        def do_resize(event):
            min_w = getattr(textw, "_min_w", 40)
            min_h = getattr(textw, "_min_h", 30)
            new_w = max(min_w, container._gw0 + (event.x_root - container._gx0))
            new_h = max(min_h, container._gh0 + (event.y_root - container._gy0))
            container.config(width=int(new_w), height=int(new_h))
            parent_canvas.itemconfigure(container._win_id, width=int(new_w), height=int(new_h))
            sz = textw._meta[0]
            box["width_pt"] = new_w / sz
            box["height_pt"] = new_h / sz
            self._position_text_box_chrome(textw)

        grip.bind("<Button-1>", start_resize)
        grip.bind("<B1-Motion>", do_resize)

        win_id = parent_canvas.create_window(x, y, anchor="nw", window=container,
                                             width=box_w, height=box_h)
        container._win_id = win_id
        textw._win_id = win_id
        # Le caselle manuali restano selezionate dopo la creazione e il bordo
        # comunica lo stato senza interferire con l'area di testo.
        if box.get("_manual") and getattr(self, "_new_text_box_to_select", None) is box:
            self._new_text_box_to_select = None
            self.root.after_idle(lambda widget=textw: self.select_pdf_box(widget))
        return container, textw, grip, win_id

    def _set_chrome_visible(self, entry_widget, visible):
        chrome = getattr(entry_widget, "_chrome", None)
        if not chrome:
            return
        if getattr(entry_widget, "_rich", False):
            if visible:
                self._position_text_box_chrome(entry_widget)
            else:
                try:
                    canvas = entry_widget.container.master
                    for wdg in chrome:
                        canvas.itemconfigure(wdg._canvas_win_id, state="hidden")
                except Exception:
                    pass
            return
        _M = {"close": (1.0, 0.0, -2, 0, "ne"),
              "drag": (0.0, 0.0, 0, 0, "nw"),
              "grip": (1.0, 1.0, 0, 0, "se")}
        for wdg, key in zip(chrome, ("close", "drag", "grip")):
            if visible:
                rx, ry, px, py, anc = _M[key]
                wdg.place(relx=rx, rely=ry, x=px, y=py, anchor=anc)
            else:
                wdg.place_forget()

    def _position_text_box_chrome(self, textw):
        """Posiziona i controlli come oggetti del canvas, agganciati alla casella."""
        try:
            container = textw.container
            canvas = container.master
            x, y = canvas.coords(container._win_id)
            canvas.update_idletasks()
            width = max(container.winfo_width(), int(canvas.itemcget(container._win_id, "width") or 0))
            close, drag, grip = textw._chrome
            widths = [w.winfo_reqwidth() for w in (close, drag, grip)]
            heights = [w.winfo_reqheight() for w in (close, drag, grip)]
            gap = 2
            total_w = sum(widths) + gap * 2
            # Centra la barra sul bordo superiore destro; se la casella e'
            # vicina al bordo alto, la barra scende appena sotto il bordo.
            left = int(x + width - total_w)
            bar_y = max(0, int(y) - max(heights))
            positions = []
            cursor_x = left
            for widget, ww in zip((close, drag, grip), widths):
                positions.append((widget, cursor_x, bar_y))
                cursor_x += ww + gap
            for widget, px, py in positions:
                canvas.coords(widget._canvas_win_id, px, py)
                canvas.itemconfigure(widget._canvas_win_id,
                                     state="normal" if textw.is_selected else "hidden")
            for control in (close, drag, grip):
                control.lift()
        except Exception:
            pass

    def _hide_all_box_chrome(self):
        for _pid, items in self.pdf_entry_widgets.items():
            for _wid, _c, entry, _d in items:
                if entry.is_selected:
                    entry.is_selected = False
                    self._set_chrome_visible(entry, False)
        for items in self.page_text_box_widgets.values():
            for container in items:
                for child in container.winfo_children():
                    if getattr(child, "_rich", False) and child.is_selected:
                        child.is_selected = False
                        self._set_chrome_visible(child, False)
                        container.configure(highlightthickness=0, highlightbackground="#000000")
        if (self.active_pdf_box is not None and
                getattr(self.active_pdf_box, "_rich", False)):
            self.active_pdf_box = None

    def select_pdf_box(self, entry_widget):
        self._deselect_page_image()
        self._hide_all_box_chrome()
        entry_widget.is_selected = True
        self._set_chrome_visible(entry_widget, True)
        self.active_pdf_box = entry_widget
        if getattr(entry_widget, "_rich", False):
            try:
                entry_widget.container.configure(
                    highlightthickness=1, highlightbackground="#000000",
                    highlightcolor="#000000",
                )
            except Exception:
                pass
        try:
            self.font_family.set(entry_widget.box_font.actual()["family"])
            fs = entry_widget.box_font.actual()["size"]
            if fs < 0:
                fs = abs(fs)
            self.font_size.set(fs)
            self.active_styles["bold"] = (entry_widget.box_font.actual()["weight"] == "bold")
            self.active_styles["italic"] = (entry_widget.box_font.actual()["slant"] == "italic")
            self._update_style_buttons()
        except Exception:
            pass

    def _sync_toolbar_from_rich(self, textw):
        """Aggiorna la toolbar con la formattazione del run sotto il cursore."""
        try:
            idx = textw.index("insert")
            if not idx:
                return
            tag = textw.tag_names(idx)
            tag = [t for t in tag if t.startswith("rr")]
            if not tag:
                return
            px, color_int, family, weight, slant = textw.tagstash[tag[0]]
            self.font_family.set(family)
            self.font_size.set(px)
            self.active_styles["bold"] = (weight == "bold")
            self.active_styles["italic"] = (slant == "italic")
            self.active_styles["underline"] = bool(textw._tag_to_run[tag[0]].get("underline", False))
            self.current_color = self._int_color_to_hex(color_int)
            self._update_style_buttons()
            self._update_color_button()
        except Exception:
            pass

    def _rich_target_runs(self, textw):
        """Ritorna i run (dict) da formattare: quelli della selezione se esiste,
        altrimenti il singolo run sotto il cursore."""
        try:
            # selezione attiva?
            has_sel = False
            try:
                textw.index("sel.first")
                has_sel = True
            except tk.TclError:
                has_sel = False
            if has_sel:
                sel = textw.tag_ranges("sel")
                tags = set()
                # colleziona i tag rr che toccano la selezione (intervalli
                # [a,b) vs [sel.first, sel.last), confronto tagliato all'estremo
                # destro escluso: un run che inizia proprio sul bordo destro
                # della selezione non e' selezionato)
                for tag in textw.tag_names():
                    if tag.startswith("rr"):
                        tr = textw.tag_ranges(tag)
                        if tr and self._ranges_overlap_strict(tr[0], tr[1], sel[0], sel[1]):
                            tags.add(tag)
                found = [textw._tag_to_run[t] for t in tags if t in textw._tag_to_run]
                if found:
                    return found
                # Selezione su testo ancora senza tag (es. casella appena creata):
                # formatta comunque tutta la casella.
                return list(textw._tag_to_run.values())
            idx = textw.index("insert")
            for probe in (idx, "insert -1c", "insert -2c"):
                try:
                    for tag in textw.tag_names(textw.index(probe)):
                        if tag.startswith("rr") and tag in textw._tag_to_run:
                            return [textw._tag_to_run[tag]]
                except tk.TclError:
                    pass
            # fallback: ultimo run della casella
            if textw._tag_to_run:
                last = list(textw._tag_to_run.values())[-1]
                return [last]
        except Exception:
            pass
        return []

    def _ranges_overlap(self, a1, b1, a2, b2):
        try:
            return self._idx_ge(b1, a2) and self._idx_ge(b2, a1)
        except Exception:
            return True

    def _ranges_overlap_strict(self, a1, b1, a2, b2):
        """Intervalli [a1,b1) e [a2,b2) in intesa 'bordo destro escluso':
        si sovrappongono se a1 < b2 e a2 < b1 (cosi' un estremo condiviso
        sul bordo destro non conta come sovrapposizione)."""
        try:
            return self._idx_lt(a1, b2) and self._idx_lt(a2, b1)
        except Exception:
            return False

    def _idx_lt(self, a, b):
        la, ca = str(a).split("."); lb, cb = str(b).split(".")
        if int(la) != int(lb):
            return int(la) < int(lb)
        return int(ca) < int(cb)

    def _idx_ge(self, a, b):
        la, ca = str(a).split("."); lb, cb = str(b).split(".")
        if int(la) != int(lb):
            return int(la) > int(lb)
        return int(ca) >= int(cb)

    def _box_all_runs(self, textw):
        """Ritorna TUTTI i run (dict) di una casella-ricca, per applicare una
        formattazione (es. la dimensione del font) a tutta la casella,
        indipendentemente dalla selezione."""
        return list(textw._tag_to_run.values())

    def _rich_underlined_runs(self, textw):
        """Ritorna i run (dict) della casella con la proprieta' 'underline' attiva."""
        return [rd for _tag, rd in textw._tag_to_run.items() if rd.get("underline")]

    def _rich_has_selection(self, textw):
        try:
            textw.index("sel.first")
            return True
        except tk.TclError:
            return False

    def _rich_split_runs_at_selection(self, textw):
        """Spezza i run della casella ai confini della selezione, cosi' la
        formattazione applicata tocca solo i caratteri selezionati (come nel
        testo delle pagine). Il testo non cambia, cambiano solo i run.
        Ritorna True se almeno un run e' stato spezzato."""
        try:
            if not self._rich_has_selection(textw):
                return False
            top = textw.index("sel.first")
            bottom = textw.index("sel.last")
        except Exception:
            return False
        try:
            box = textw.data_obj
            lines = box.get("lines", [])
            offs = []
            acc = 0
            for line in lines:
                offs.append(acc)
                acc += sum(len(r.get("text", "")) for r in line) + 1
            def abs_pos(idx):
                li, ci = idx.split(".")
                return offs[min(int(li) - 1, len(offs) - 1)] + int(ci)
            top = abs_pos(top)
            bottom = abs_pos(bottom)
            cursor = 0
            changed = False
            new_lines = []
            for line in lines:
                line_len = sum(len(r.get("text", "")) for r in line)
                if not line:
                    new_lines.append(line)
                    cursor += 1
                    continue
                line_start = cursor
                line_end = cursor + line_len
                if top >= line_end or bottom <= line_start:
                    new_lines.append(line)
                    cursor = line_end + 1
                    continue
                # spezza la riga ai confini (top-start, bottom-start)
                local_pieces = []
                pos = 0
                for rd in line:
                    local_pieces.append((rd, pos, pos + len(rd.get("text", ""))))
                    pos += len(rd.get("text", ""))
                cuts = {0, pos}
                for rd0, a, b in local_pieces:
                    ca = max(a, top - line_start)
                    cb = min(b, bottom - line_start)
                    if 0 <= ca <= pos:
                        cuts.add(ca)
                    if 0 <= cb <= pos:
                        cuts.add(cb)
                cuts = sorted(cuts)
                new_runs_line = []
                for i in range(len(cuts) - 1):
                    ca, cb = cuts[i], cuts[i + 1]
                    if ca >= cb:
                        continue
                    for rd, a, b in local_pieces:
                        if a <= ca and cb <= b:
                            nd = dict(rd)
                            nd["text"] = rd.get("text", "")[ca - a: cb - a]
                            new_runs_line.append(nd)
                            break
                if sum(len(r.get("text", "")) for r in new_runs_line) == line_len:
                    if len(new_runs_line) != len(line):
                        changed = True
                    new_lines.append(new_runs_line)
                else:
                    new_lines.append(line)
                cursor = line_end + 1
            if changed:
                box["lines"] = new_lines
                box["runs"] = [r for line in new_lines for r in line]
                self._retag_rich_textw(textw, new_lines)
            return changed
        except Exception:
            return False

    def _apply_rich_format(self, textw, family=None, size_px=None, bold=None, italic=None, underline=None, color=None, target_runs=None):
        """Applica la formattazione richiesta ai run sotto selezione/cursore di
        una casella-ricca, aggiornando i tag, la toolbar e i dati del box.
        Se target_runs e' fornito, formatta solo quei run (es. tutta la casella
        per la dimensione del font)."""
        if textw is None or not getattr(textw, "_rich", False):
            return
        scale_z = textw._meta[0]
        if target_runs is None:
            # formattazione legata a selezione/cursore: spezza i run ai confini
            # della selezione cosi' le modifiche toccano solo i caratteri scelti
            self._rich_split_runs_at_selection(textw)
        for rd in (target_runs if target_runs is not None else self._rich_target_runs(textw)):
            accum = dict(rd)
            if family is not None:
                accum["family"] = family
            if size_px is not None:
                # salva in punti per il ri-render; px = punti*scale_z*_FONT_SCALE
                accum["size"] = size_px / (scale_z * self._FONT_SCALE)
            if bold is not None:
                accum["weight"] = "bold" if bold else "normal"
            if italic is not None:
                accum["slant"] = "italic" if italic else "roman"
            if underline is not None:
                accum["underline"] = bool(underline)
            if color is not None:
                accum["color"] = int(color.replace("#", ""), 16)
            rd.update(accum)
            # aggiorna il tag corrispondente a questo run
            for tag, rref in textw._tag_to_run.items():
                if rref is rd:
                    try:
                        px = max(6, int(round(rd["size"] * scale_z * self._FONT_SCALE)))
                        f = font.Font(family=rd["family"], size=px, weight=rd["weight"], slant=rd["slant"])
                        textw.tag_configure(tag, font=f, foreground=self._int_color_to_hex(rd["color"]),
                                            underline=bool(rd.get("underline", False)))
                        textw.tagstash[tag] = (px, rd["color"], rd["family"], rd["weight"], rd["slant"])
                    except Exception:
                        pass
        # aggiorna toolbar
        self._sync_toolbar_from_rich(textw)
        # aggiorna anche box custom (per il rendering se la casella e' "base")
        rd0 = self._rich_target_runs(textw)
        box = textw.data_obj
        if rd0:
            r = rd0[0]
            box["custom_family"] = r.get("family", box.get("custom_family", "Arial"))
            box["custom_weight"] = r.get("weight", box.get("custom_weight", "normal"))
            box["custom_slant"] = r.get("slant", box.get("custom_slant", "roman"))
            box["custom_size"] = max(6, int(round(r.get("size", 12) * scale_z * self._FONT_SCALE)))
            box["custom_color"] = r.get("color", box.get("custom_color", 0))

    def _flush_rich_text(self, textw):
        """Salva il testo modificato della casella-ricca nel box dati.

        I tag tk.Text restano attaccati ai caratteri durante la modifica, quindi
        qui si limitano ad aggiornare le stringhe di testo dei run e del box.
        """
        try:
            box = textw.data_obj
            content = textw.get("1.0", "end-1c")
            # ricostruisce le righe: ogni riga del Text -> run allineati
            txt_lines = content.split("\n")
            orig_lines = box["lines"]
            new_lines = []
            new_runs = []
            for i, line_text in enumerate(txt_lines):
                if i < len(orig_lines) and orig_lines[i]:
                    src_runs = orig_lines[i]
                    total_src = sum(len(r["text"]) for r in src_runs)
                    # distribuisce il nuovo testo in modo proporzionale
                    runs_here = []
                    pos = 0
                    for r in src_runs:
                        frac = (len(r["text"]) / total_src) if total_src else (1.0 / max(1, len(src_runs)))
                        seg_len = max(0, int(round(len(line_text) * frac)))
                        seg = line_text[pos:pos + seg_len]
                        r = dict(r)
                        r["text"] = seg
                        runs_here.append(r)
                        new_runs.append(r)
                        pos += seg_len
                    # eventuale testo eccedente lo si appende all'ultimo run
                    if pos < len(line_text) and runs_here:
                        runs_here[-1]["text"] += line_text[pos:]
                        new_runs[-1]["text"] = runs_here[-1]["text"]
                    new_lines.append(runs_here)
                else:
                    runs_here = []
                    if line_text:
                        base = orig_lines[i - 1][-1] if (i > 0 and orig_lines[i - 1]) else None
                        if base is None:
                            base = orig_lines[0][0] if orig_lines and orig_lines[0] else None
                        if base is not None:
                            rd = dict(base)
                            rd["text"] = line_text
                            runs_here = [rd]
                            new_runs.append(rd)
                    new_lines.append(runs_here)
            box["lines"] = new_lines
            box["runs"] = new_runs
            self._retag_rich_textw(textw, new_lines)
        except Exception:
            pass

    def _retag_rich_textw(self, textw, new_lines):
        """Ri-applica i tag dei run su TUTTO il contenuto del text widget.

        Soprattutto per le caselle create a mano (che nascono vuote) il testo
        digitato entra SENZA tag: senza questo ripasso, grassetto/corsivo/
        colore non avrebbero effetto. Ricostruisce tag, tagstash e font-tag
        per farli combaciare con i dati del box."""
        for tag in list(textw._tag_to_run.keys()):
            try:
                textw.tag_remove(tag, "1.0", "end")
            except Exception:
                pass
        textw._tag_to_run = {}
        scale_z = textw._meta[0]
        for li, runs_here in enumerate(new_lines):
            col = 0
            for ri, rd in enumerate(runs_here):
                tag = "rr%d_%d" % (li, ri)
                textw._tag_to_run[tag] = rd
                seg = rd.get("text", "")
                if seg:
                    try:
                        textw.tag_add(tag, "%d.%d" % (li + 1, col), "%d.%d" % (li + 1, col + len(seg)))
                    except Exception:
                        pass
                try:
                    px = max(6, int(round(rd["size"] * scale_z * self._FONT_SCALE)))
                    f = font.Font(family=rd["family"], size=px, weight=rd["weight"], slant=rd["slant"])
                    textw.tag_configure(tag, font=f, foreground=self._int_color_to_hex(rd["color"]),
                                        underline=bool(rd.get("underline", False)))
                    textw.tagstash[tag] = (px, rd["color"], rd["family"], rd["weight"], rd["slant"])
                except Exception:
                    pass
                col += len(seg)

    def _update_color_button(self):
        try:
            self.main_color_btn.configure(bg=self.current_color)
        except Exception:
            pass

    def _autosize_box(self, entry_widget):
        """Riadatta la casella alla larghezza del nuovo testo digitato."""
        container = entry_widget.container
        if container is None:
            return
        tw = entry_widget.box_font.measure(entry_widget.get())
        pad_x = 6
        fit_w = max(tw, 10) + pad_x
        cur_h = container.winfo_height()
        container.config(width=fit_w)
        if container._win_id:
            self.pdf_canvas.itemconfigure(container._win_id, width=fit_w, height=cur_h)
        entry_widget.data_obj["width_pt"] = fit_w / entry_widget._meta[0]
        entry_widget.data_obj["height_pt"] = cur_h / entry_widget._meta[0]

    def _delete_pdf_box(self, page_id, data_obj):
        info = self.pdf_pages_info.get(page_id)
        if not info:
            return
        # rimuove sia dagli span che dagli overlay
        if data_obj in info.get("spans", []):
            info["spans"].remove(data_obj)
        if data_obj in info.get("overlays", []):
            info["overlays"].remove(data_obj)
        # casella-paragrafo: rimuove il box e i relativi span
        if data_obj in info.get("boxes", []):
            info["boxes"].remove(data_obj)
            for rd in data_obj.get("runs", []):
                sp = rd.get("span")
                if sp is not None and sp in info.get("spans", []):
                    info["spans"].remove(sp)
        if getattr(self, "active_pdf_box", None) is not None and getattr(self.active_pdf_box, "data_obj", None) is data_obj:
            self.active_pdf_box = None
        try:
            info["spans_removed"] = True
        except Exception:
            pass
        self.render_all_pdf_pages()

    def set_active_pdf_box(self, entry_widget):
        self.select_pdf_box(entry_widget)

    def _create_overlay_entry(self, page_id, ov, scale, y_offset=0, x_offset=0, z=1.0, page_w=595):
        ov["_x_pt"] = ov.get("_x_pt", ov.get("x", 0))
        ov["_y_pt"] = ov.get("_y_pt", ov.get("y", 0))
        cx = x_offset + ov["_x_pt"] * scale * z
        cy = y_offset + ov["_y_pt"] * scale * z
        pt_size = max(6, int(round(float(ov.get("size", 12.0)) * z)))

        ov["_scale"] = scale * z
        ov["custom_font"] = ov.get("custom_font", "Arial")
        ov["custom_size"] = pt_size
        ov["custom_weight"] = ov.get("custom_weight", "normal")
        ov["custom_slant"] = ov.get("custom_slant", "roman")

        container, entry, grip = self._create_resizable_box(
            self.pdf_canvas, cx, cy, 180, pt_size + 14,
            ov["text"], ov["custom_font"], pt_size, ov["custom_weight"], ov["custom_slant"], "#000000", ov,
            page_id, (scale * z, x_offset, y_offset, page_w)
        )
        win_id = self.pdf_canvas.create_window(cx, cy, anchor="nw", window=container,
                                               width=container.winfo_reqwidth(), height=container.winfo_reqheight())
        container._win_id = win_id
        self.pdf_entry_widgets[page_id].append((win_id, container, entry, ov))

    def toggle_add_text_mode(self):
        self.pdf_add_mode = not self.pdf_add_mode
        if self.pdf_mode:
            self.pdf_canvas.config(cursor="crosshair" if self.pdf_add_mode else "")
        else:
            for ed in self.text_page_editors.values():
                ed.config(cursor="crosshair" if self.pdf_add_mode else "")
        self.btn_add_text.state(["pressed"] if self.pdf_add_mode else ["!pressed"])

    def on_editor_click_for_textbox(self, event, page_id):
        if self.pdf_mode or not self.pdf_add_mode:
            return
        ed = self.text_page_editors.get(page_id)
        if not ed: return
        info = ed.place_info()
        try:
            ex = int(float(info.get("x", 0)))
            ey = int(float(info.get("y", 0)))
        except Exception:
            ex, ey = 0, 0
        x = ex + event.x
        y = ey + event.y
        self._add_text_box_to_page(page_id, x, y)
        self.toggle_add_text_mode()
        return "break"

    def _add_text_box_to_page(self, page_id, x, y):
        try:
            sz = int(self.font_size.get())
        except Exception:
            sz = 12
        fam = "Arial"
        try:
            color_int = int(str(self.current_color).lstrip("#") or "000000", 16)
        except Exception:
            color_int = 0
        run = {
            "text": "",
            "size": sz,
            "family": fam,
            "weight": "bold" if self.active_styles["bold"] else "normal",
            "slant": "italic" if self.active_styles["italic"] else "roman",
            "color": color_int,
        }
        # Stessa struttura dati usata dalle caselle-ricche dei PDF, cosi' la
        # casella creata manualmente e' identica a quella dei PDF (stesso
        # widget, stesso chrome nascosto, stesso ridimensionamento).
        box = {
            "lines": [[run]],
            "runs": [run],
            "bbox": [x, y, x + 170, y + 45],
            "_x_pt": x, "_y_pt": y,
            "x_pt": x, "y_pt": y,
            "x": x, "y": y,
            "width": 170, "height": 45,
            "min_w": 170, "min_h": 45,
            "background": "#ffffff",
            "text": "",
            "family": fam,
            "size": sz,
            "custom_family": fam,
            "custom_size": sz,
            "custom_weight": run["weight"],
            "custom_slant": run["slant"],
            "custom_color": color_int,
            "_manual": True,
        }
        self.page_text_boxes.setdefault(page_id, []).append(box)
        self._new_text_box_to_select = box
        self._mount_text_box(page_id, box)

    def _mount_text_box(self, page_id, box):
        p_frame = self.text_page_frames.get(page_id)
        if not p_frame:
            return
        page_canvas = p_frame.master
        z = self.zoom_factor
        base_w, base_h = self.base_page_sizes[self.current_format]
        if self.current_orientation == "Orizzontale":
            pw, ph = int(base_h * z), int(base_w * z)
        else:
            pw, ph = int(base_w * z), int(base_h * z)
        x0 = box.get("_x_pt", box.get("x_pt", box.get("x", 10))) or 10
        y0 = box.get("_y_pt", box.get("y_pt", box.get("y", 10))) or 10
        cx = x0 * z
        cy = y0 * z
        container, textw, grip, win_id = self._create_rich_text_box(
            page_canvas, cx, cy, box, z, page_id,
            (z, 0, 0, pw, ph)
        )
        container._win_id = win_id
        self.page_text_box_widgets.setdefault(page_id, []).append(container)
        box["width_pt"] = (box.get("width", 170) or 170) / z
        box["height_pt"] = (box.get("height", 45) or 45) / z
        # La X di chiusura elimina la casella qui in modalita' testo, non nel
        # pdf_pages_info (che non la contiene).
        for lbl in getattr(container, "_chrome", ()):
            try:
                if lbl.cget("text") == "✕":
                    lbl.unbind("<Button-1>")
                    lbl.bind("<Button-1>", lambda e, pid=page_id, b=box, c=container:
                             self._remove_text_box(pid, b, c))
            except Exception:
                pass
        # I comandi sono figli del canvas pagina, così possono stare fuori
        # dalla casella senza coprirne il testo.
        rich_editor = next((w for w in container.winfo_children()
                            if getattr(w, "_rich", False)), None)
        if rich_editor is not None:
            close_btn = rich_editor._chrome[0]
            close_btn.bind("<Button-1>", lambda e, pid=page_id, b=box, c=container:
                            self._remove_text_box(pid, b, c))
            self._position_text_box_chrome(rich_editor)
        return container

    def _remove_text_box(self, page_id, box, container):
        rich_editor = next((w for w in container.winfo_children()
                            if getattr(w, "_rich", False)), None)
        if rich_editor is not None:
            for control in getattr(rich_editor, "_chrome", ()):
                try:
                    control.destroy()
                except Exception:
                    pass
        try:
            self.page_text_boxes.get(page_id, []).remove(box)
        except ValueError:
            pass
        try:
            container.destroy()
        except Exception:
            pass
        widgets = self.page_text_box_widgets.get(page_id, [])
        if container in widgets:
            widgets.remove(container)
        if self.active_pdf_box is not None and getattr(self.active_pdf_box, "data_obj", None) is box:
            self.active_pdf_box = None

    def _clear_text_box_widgets(self, page_id):
        for w in self.page_text_box_widgets.get(page_id, []):
            try:
                for child in w.winfo_children():
                    if getattr(child, "_rich", False):
                        for control in getattr(child, "_chrome", ()):
                            try:
                                control.destroy()
                            except Exception:
                                pass
                w.destroy()
            except Exception:
                pass
        self.page_text_box_widgets[page_id] = []

    def _render_text_boxes_for_page(self, page_id):
        if self.pdf_mode:
            return
        self._clear_text_box_widgets(page_id)
        for box in self.page_text_boxes.get(page_id, []):
            self._mount_text_box(page_id, box)

    def _deselect_text_box(self, event=None):
        """Deseleziona le caselle manuali quando si clicca sul testo del foglio."""
        if self.active_pdf_box is not None and getattr(self.active_pdf_box, "_rich", False):
            self._hide_all_box_chrome()
        return None

    # ----------------------------------------------------------------------
    # ----------------------------------------------------------------------
    # IMMAGINI SUL FOGLIO (testo): oggetti visibili sopra l'editor, sempre
    # sotto le caselle di testo, con bordo di selezione, spostabili e
    # ridimensionabili. GIF animate in loop senza lag.
    #
    # NOTA: l'editor di testo e' un widget opaco che Tk disegna SEMPRE sopra
    # gli item di un canvas (anche con z-order). Per rendere l'immagine
    # visibile sul foglio, ogni immagine viene disegnata su un piccolo canvas
    # creato come "window" del p_canvas DOPO la finestra dell'editor: cosi'
    # l'immagine sta fisicamente sopra l'editor (e sotto le caselle di testo,
    # che vengono montate dopo).
    # ----------------------------------------------------------------------

    def _image_filetypes(self):
        return [
            ("Immagini supportate", "*.png *.jpg *.jpeg *.gif *.bmp *.webp *.tiff *.tif *.ico"),
            ("Tutti i file", "*.*"),
        ]

    def insert_image_dialog(self):
        if self.pdf_mode:
            messagebox.showinfo("Inserisci immagine",
                                "Inserisci l'immagine dalla vista testo (menu Inserisci).")
            return
        path = filedialog.askopenfilename(parent=self.root, title="Inserisci immagine",
                                          filetypes=self._image_filetypes())
        if not path:
            return
        self._add_image_to_page(self.active_page_id, path)

    def _add_image_to_page(self, page_id, path):
        """Aggiunge un oggetto immagine al foglio della pagina (in punti, per
        restare coerente con zoom/orientamento). La posizione iniziale e'
        centrata in alto; la larghezza e' il 50% della pagina."""
        try:
            with Image.open(path) as img:
                nat_w, nat_h = img.size
        except Exception:
            messagebox.showerror("Errore", "Impossibile leggere l'immagine selezionata.")
            return
        base_w, _ = self.base_page_sizes[self.current_format]
        target_w = base_w * 0.5
        ratio = target_w / nat_w if nat_w else 1
        w_pt = int(round(target_w))
        h_pt = max(1, int(round(nat_h * ratio)))
        x_pt = int((base_w - w_pt) / 2)
        y_pt = 45
        record = {
            "path": path,
            "x_pt": x_pt,
            "y_pt": y_pt,
            "w_pt": w_pt,
            "h_pt": h_pt,
        }
        self.page_images.setdefault(page_id, []).append(record)
        self.update_page_layout()

    def _cancel_image_animation(self, page_id):
        jobs = self._img_after_jobs.pop(page_id, [])
        for jid in jobs:
            try:
                self.root.after_cancel(jid)
            except Exception:
                pass

    def _clear_image_widgets(self, page_id):
        self._cancel_image_animation(page_id)
        for jid in self._img_resize_jobs:
            try:
                self.root.after_cancel(jid)
            except Exception:
                pass
        self._img_resize_jobs = []
        refs = self._img_refs.pop(page_id, {})
        for net_id, photos in list(refs.items()):
            for ph in (photos if isinstance(photos, (list, tuple)) else [photos]):
                try:
                    ph._image_no_ref = True
                except Exception:
                    pass
        for rec in self.page_image_widgets.get(page_id, []):
            try:
                rec.get("_widget").destroy()
            except Exception:
                pass
        self.page_image_widgets[page_id] = []
        self._img_refs.pop(page_id, None)

    def _load_image_frames(self, path):
        """Carica i frame originali (risoluzione nativa, RGBA) di un'immagine.
        Ritorna (frames_pil, n_frames, delay_ms, is_animated)."""
        try:
            img = Image.open(path)
            img.seek(0)
            frames = []
            is_anim = False
            delay = 100
            try:
                n = img.n_frames
            except Exception:
                n = 1
            if getattr(img, "is_animated", False) and n > 1:
                is_anim = True
                for i in range(n):
                    try:
                        img.seek(i)
                        fr = img.convert("RGBA")
                        if i == 0:
                            try:
                                delay = int(img.info.get("duration", 100)) or 100
                            except Exception:
                                delay = 100
                        frames.append(fr)
                    except Exception:
                        break
            else:
                frames = [img.convert("RGBA")]
            img.close()
            return frames, len(frames), delay, is_anim
        except Exception:
            return [], 0, 0, False

    def _build_photos(self, frames, w_px, h_px):
        """Converte i frame PIL in PhotoImage ridimensionati a w_px x h_px."""
        photos = []
        for fr in frames:
            if w_px > 0 and h_px > 0:
                fr = fr.resize((w_px, h_px), Image.Resampling.LANCZOS)
            photos.append(ImageTk.PhotoImage(fr))
        return photos

    def _load_image_photo(self, path, w_px, h_px):
        """Carica i frame dell'immagine ridimensionati a w_px x h_px.
        Ritorna (lista_photo, n_frames, delay_ms, is_animated, frames_pil)."""
        frames, n, delay, is_anim = self._load_image_frames(path)
        if not frames:
            return [], 0, 0, False, []
        photos = self._build_photos(frames, w_px, h_px)
        return photos, len(photos), delay, is_anim, frames

    def _rescale_current(self, record):
        """Ridimensiona i PhotoImage dell'oggetto dai frame originali salvati,
        senza riaprire il file. Usato per il resize fluido dell'immagine."""
        st = record.get("_state")
        if not st:
            return
        w = max(1, st["w"])
        h = max(1, st["h"])
        frames = record.get("_frames") or []
        if not frames:
            return
        photos = self._build_photos(frames, w, h)
        st["photos"] = photos
        st["frame"] = min(st["frame"], len(photos) - 1)
        ic = st.get("img_canvas")
        try:
            ic.itemconfigure(st["img_id"], image=photos[st["frame"]])
        except Exception:
            pass
        return photos

    def _draw_page_image(self, page_id, record, p_canvas):
        """Crea il canvas-oggetto dell'immagine e lo piazza come finestra sul
        p_canvas DOPO l'editor, quindi visibile sopra. Ritorna il record
        aggiornato (con _state)."""
        z = self.zoom_factor
        w_px = max(1, int(record["w_pt"] * z))
        h_px = max(1, int(record["h_pt"] * z))
        photos, n, delay, is_anim, frames = self._load_image_photo(record["path"], w_px, h_px)
        if not photos:
            return None
        record["_frames"] = frames
        x = int(record["x_pt"] * z)
        y = int(record["y_pt"] * z)

        # il canvas oggetto: bianco come la pagina, con l'immagine disegnata
        chrome_margin = 14
        img_canvas = tk.Canvas(p_canvas, width=w_px + chrome_margin * 2,
                               height=h_px + chrome_margin * 2, bg="#ffffff",
                               highlightthickness=0, bd=0)
        base_tag = "imgbase"
        img_id = img_canvas.create_image(chrome_margin, chrome_margin, anchor="nw",
                                         image=photos[0], tags=(base_tag,))
        img_canvas.tag_raise(base_tag)

        # piazzato come window sopra il testo della pagina e sotto le caselle mobili
        win_id = p_canvas.create_window((x - chrome_margin, y - chrome_margin),
                                        window=img_canvas, anchor="nw",
                                        tags=("page_image",))

        record["_state"] = {
            "photos": photos, "n": n, "delay": delay, "is_anim": is_anim,
            "frame": 0, "x": x, "y": y, "w": w_px, "h": h_px,
            "img_canvas": img_canvas, "img_id": img_id, "win_id": win_id,
            "chrome_margin": chrome_margin,
        }
        record["_items"] = [win_id]
        record.setdefault("_chrome", [])
        record["_page"] = page_id
        record["_canvas"] = p_canvas
        record["_killed"] = False

        page_editor = self.text_page_editors.get(page_id)

        # se clicchi l'immagine: seleziona; altrimenti inoltra il clic all'editor
        img_canvas.bind("<ButtonPress-1>",
                        lambda e, pid=page_id, rec=record, cv=p_canvas, ic=img_canvas:
                        self._image_press(pid, rec, cv, ic, e))
        img_canvas.bind("<B1-Motion>",
                        lambda e, pid=page_id, rec=record, cv=p_canvas, ic=img_canvas:
                        self._image_move(pid, rec, cv, ic, e))
        img_canvas.bind("<ButtonRelease-1>",
                        lambda e, pid=page_id, rec=record, cv=p_canvas, ic=img_canvas:
                        self._image_release(pid, rec, cv, ic, e))
        # inoltra la rotellina del mouse all'editor sotto, per far scorrere la pagina
        for evseq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            img_canvas.bind(evseq,
                            lambda e, ed=page_editor: self._forward_wheel(e, ed), add="+")

        # conserva il riferimento per la pulizia
        self.page_image_widgets.setdefault(page_id, []).append(record)

        if is_anim and n > 1:
            old_job = record.get("_anim_job")
            if old_job is not None:
                try:
                    self.root.after_cancel(old_job)
                except Exception:
                    pass
                jobs = self._img_after_jobs.get(page_id, [])
                if old_job in jobs:
                    jobs.remove(old_job)
            self._image_animate(page_id, record, p_canvas)
        return record

    def _forward_wheel(self, event, editor):
        """Inoltra la rotellina del mouse all'editor di testo sottostante,
        cosi' la pagina continua a scorrere anche sopra un'immagine."""
        try:
            if editor is not None and editor.winfo_exists():
                if event.num == 4:
                    editor.yview_scroll(-1, "units")
                elif event.num == 5:
                    editor.yview_scroll(1, "units")
                elif getattr(event, "delta", None):
                    editor.yview_scroll(-1 if event.delta > 0 else 1, "units")
        except Exception:
            pass
        return "break"

    def _image_animate(self, page_id, record, p_canvas):
        st = record.get("_state")
        if not st or not st.get("is_anim") or record.get("_killed"):
            return
        ic = st.get("img_canvas")
        if ic is None or not ic.winfo_exists():
            return
        st["frame"] = (st["frame"] + 1) % st["n"]
        try:
            ic.itemconfigure(st["img_id"], image=st["photos"][st["frame"]])
        except Exception:
            return
        jid = self.root.after(st["delay"], lambda: self._image_animate(page_id, record, p_canvas))
        jobs = self._img_after_jobs.setdefault(page_id, [])
        old = record.get("_anim_job")
        if old in jobs:
            jobs.remove(old)
        jobs.append(jid)
        record["_anim_job"] = jid

    def _render_page_images(self, page_id, page_canvas):
        """Rende le immagini della pagina: ogni immagine e' un piccolo canvas
        piazzato come finestra DOPO l'editor (quindi sopra di esso) e prima
        delle caselle di testo."""
        self._clear_image_widgets(page_id)
        for record in self.page_images.get(page_id, []):
            self._draw_page_image(page_id, record, page_canvas)

    def _image_press(self, page_id, record, p_canvas, img_canvas, event):
        # Se era già selezionata, non cancellare e ricreare le maniglie sotto
        # il puntatore: il click può essere proprio l'inizio di un resize.
        if self.active_pdf_image is not record:
            self._deselect_page_image()
            self.active_pdf_image = record
            self._draw_image_chrome(page_id, record, p_canvas, img_canvas)
        record["_press_x_root"] = event.x_root
        record["_press_y_root"] = event.y_root
        st = record["_state"]
        record["_press_img_x"] = st["x"]
        record["_press_img_y"] = st["y"]

    def _image_move(self, page_id, record, p_canvas, img_canvas, event):
        st = record.get("_state")
        if not st:
            return
        # Le maniglie hanno un proprio binding di ridimensionamento. Il binding
        # generale del canvas riceve lo stesso evento: non deve trasformarlo
        # anche in uno spostamento dell'immagine.
        if record.get("_resizing"):
            return
        z = self.zoom_factor
        # Coordinate schermo stabili: event.x/y sono relativi al canvas
        # dell'immagine, che si sposta sotto il puntatore durante il drag.
        dx = event.x_root - record.get("_press_x_root", event.x_root)
        dy = event.y_root - record.get("_press_y_root", event.y_root)
        nx = record.get("_press_img_x", st["x"]) + dx
        ny = record.get("_press_img_y", st["y"]) + dy
        record["x_pt"] = nx / z
        record["y_pt"] = ny / z
        st["x"] = int(nx)
        st["y"] = int(ny)
        try:
            margin = st.get("chrome_margin", 14)
            p_canvas.coords(st["win_id"], st["x"] - margin, st["y"] - margin)
        except Exception:
            pass
        if self.active_pdf_image is record:
            self._update_image_chrome_geometry(record, img_canvas)

    def _update_image_chrome_geometry(self, record, img_canvas=None):
        """Aggiorna le maniglie esistenti senza ricrearle durante il drag."""
        try:
            st = record["_state"]
            canvas = img_canvas or st["img_canvas"]
            ids = record.get("_chrome", [])
            if len(ids) < 8:
                return
            w, h = st["w"], st["h"]
            m = st.get("chrome_margin", 14)
            hs = 10
            canvas.coords(ids[0], m, m, w + m, h + m)
            canvas.coords(ids[1], m - 5, m - 5, w + m + 5, h + m + 5)
            canvas.coords(ids[2], w + m + 1, m - 11, w + m + 13, m + 1)
            canvas.coords(ids[3], w + m + 6, m - 6)
            for iid, (hx, hy) in zip(ids[4:], ((m - hs, m - hs), (w + m, m - hs),
                                               (m - hs, h + m), (w + m, h + m))):
                canvas.coords(iid, hx, hy, hx + hs, hy + hs)
            canvas.tag_raise("pageimg_chrome")
        except Exception:
            pass

    def _draw_image_chrome(self, page_id, record, p_canvas, img_canvas=None):
        if img_canvas is None:
            img_canvas = record.get("_state", {}).get("img_canvas")
        self._delete_chrome(record, img_canvas)
        st = record["_state"]
        w, h = st["w"], st["h"]
        m = st.get("chrome_margin", 14)
        tag = "pageimg_chrome"
        # bordo nero attorno all'immagine
        border = img_canvas.create_rectangle(m, m, w + m, h + m,
                                             outline="#000000", width=1, tags=(tag,))
        # cornice esterna tratteggiata
        selsq = img_canvas.create_rectangle(m - 5, m - 5, w + m + 5, h + m + 5,
                                            outline="#000000", dash=(4, 2),
                                            width=1, tags=(tag,))
        # pulsante per ELIMINARE l'immagine (in alto a destra)
        close = img_canvas.create_text(w + m + 6, m - 6, text="\u2715", fill="#ffffff",
                                       font=("Arial", 10, "bold"), tags=(tag,))
        close_bg = img_canvas.create_oval(w + m + 1, m - 11, w + m + 13, m + 1, fill="#d32f2f",
                                          outline="#d32f2f", tags=(tag,))
        # maniglie d'angolo per distorcere/ridimensionare
        handles = []
        hs = 10
        positions = {
            "nw": (m - hs, m - hs),
            "ne": (w + m, m - hs),
            "sw": (m - hs, h + m),
            "se": (w + m, h + m),
        }
        for corner, (hx, hy) in positions.items():
            hd = img_canvas.create_rectangle(hx, hy, hx + hs, hy + hs,
                                             fill="#1a73e8", outline="#ffffff",
                                             width=1, tags=(tag,))
            img_canvas.tag_bind(hd, "<ButtonPress-1>",
                                lambda e, c=corner, pid=page_id, rec=record, cv=p_canvas, ic=img_canvas:
                                self._image_resize_press(pid, rec, cv, ic, e, corner=c))
            img_canvas.tag_bind(hd, "<B1-Motion>",
                                lambda e, pid=page_id, rec=record, cv=p_canvas, ic=img_canvas:
                                self._image_resize_move(pid, rec, cv, ic, e))
            img_canvas.tag_bind(hd, "<ButtonRelease-1>",
                                lambda e, pid=page_id, rec=record, cv=p_canvas, ic=img_canvas:
                                self._image_release(pid, rec, cv, ic, e))
            handles.append(hd)
        record["_chrome"] = [border, selsq, close_bg, close] + handles
        # l'immagine resta sotto il chrome (le maniglie sopra)
        try:
            img_canvas.tag_raise("pageimg_chrome")
            img_canvas.tag_lower(record["_state"]["img_id"])
            img_canvas.tag_raise(close_bg)
            img_canvas.tag_raise(close)
        except Exception:
            pass
        img_canvas.tag_bind(close, "<Button-1>",
                            lambda e, pid=page_id, rec=record, cv=p_canvas, ic=img_canvas:
                            self._remove_page_image(pid, rec, cv, ic))

    def _delete_chrome(self, record, img_canvas=None):
        if img_canvas is None:
            img_canvas = record.get("_state", {}).get("img_canvas")
        for iid in record.get("_chrome", []):
            try:
                img_canvas.delete(iid)
            except Exception:
                pass
        record["_chrome"] = []

    def _redraw_chrome(self, page_id, record, p_canvas, img_canvas=None):
        if self.active_pdf_image is record:
            self._draw_image_chrome(page_id, record, p_canvas, img_canvas)

    def _image_resize_press(self, page_id, record, p_canvas, img_canvas, event, corner="se"):
        st = record.get("_state")
        if not st:
            return
        self.active_pdf_image = record
        record["_resizing"] = True
        record["_rsz_corner"] = corner
        record["_rsz_x0"] = event.x_root
        record["_rsz_y0"] = event.y_root
        record["_rsz_w0"] = st["w"]
        record["_rsz_h0"] = st["h"]
        record["_rsz_x0_pt"] = record["x_pt"]
        record["_rsz_y0_pt"] = record["y_pt"]

    def _image_resize_move(self, page_id, record, p_canvas, img_canvas, event):
        st = record.get("_state")
        if not st:
            return
        z = self.zoom_factor
        dx = event.x_root - record.get("_rsz_x0", event.x_root)
        dy = event.y_root - record.get("_rsz_y0", event.y_root)
        dw = dx
        dh = dy
        corner = record.get("_rsz_corner", "se")
        old_w = record.get("_rsz_w0", st["w"])
        old_h = record.get("_rsz_h0", st["h"])
        new_w = max(20, int(old_w - dx if "w" in corner else old_w + dx))
        new_h = max(20, int(old_h - dy if "n" in corner else old_h + dy))
        base_x_pt = record.get("_rsz_x0_pt", record["x_pt"])
        base_y_pt = record.get("_rsz_y0_pt", record["y_pt"])
        # per gli angoli sx/top l'ancora resta l'angolo opposto: muovi la posizione
        if "w" in corner:
            record["x_pt"] = base_x_pt + dx / z
        else:
            record["x_pt"] = base_x_pt
        if "n" in corner:
            record["y_pt"] = base_y_pt + dy / z
        else:
            record["y_pt"] = base_y_pt
        record["w_pt"] = new_w / z
        record["h_pt"] = new_h / z
        # aggiorna subito (leggero): dimensioni del widget e posizione
        x = int(record["x_pt"] * z)
        y = int(record["y_pt"] * z)
        st["w"] = new_w
        st["h"] = new_h
        st["x"] = x
        st["y"] = y
        try:
            margin = st.get("chrome_margin", 14)
            img_canvas.config(width=new_w + margin * 2, height=new_h + margin * 2)
            p_canvas.coords(st["win_id"], x - margin, y - margin)
        except Exception:
            pass
        # ridimensiona realmente l'immagine in modo throttlato (senza riaprire file)
        job = self.root.after(30, lambda: self._rescale_current(record))
        self._img_resize_jobs.append(job)
        if self.active_pdf_image is record:
            self._update_image_chrome_geometry(record, img_canvas)

    def _image_release(self, page_id, record, p_canvas, img_canvas, event):
        st = record.get("_state")
        if st:
            try:
                self._rescale_current(record)
            except Exception:
                pass
        record["_resizing"] = False

    def _rerender_page_image(self, page_id, record, p_canvas, img_canvas=None):
        """Ridisegna completamente un oggetto immagine (usato a fine resize o
        su render successivi). Riapre solo il file quando serve; durante il
        resize si usa _rescale_current (in memoria, fluido)."""
        for iid in record.get("_items", []):
            try:
                p_canvas.delete(iid)
            except Exception:
                pass
        try:
            record.get("_state", {}).get("img_canvas").destroy()
        except Exception:
            pass
        record["_killed"] = True
        new_rec = self._draw_page_image(page_id, record, p_canvas)
        if new_rec is None:
            return
        if self.active_pdf_image is record:
            self._draw_image_chrome(page_id, record, p_canvas, record["_state"]["img_canvas"])

    def _remove_page_image(self, page_id, record, p_canvas=None, img_canvas=None):
        if p_canvas is None:
            p_canvas = record.get("_canvas")
        for iid in record.get("_items", []):
            try:
                p_canvas.delete(iid)
            except Exception:
                pass
        self._delete_chrome(record, img_canvas)
        try:
            record.get("_state", {}).get("img_canvas").destroy()
        except Exception:
            pass
        try:
            refs = self._img_refs.get(page_id, {})
            for key in [k for k, v in refs.items() if v is record.get("_state", {}).get("photos")]:
                del refs[key]
        except Exception:
            pass
        record["_killed"] = True
        records = self.page_images.get(page_id, [])
        if record in records:
            records.remove(record)
        try:
            wl = self.page_image_widgets.get(page_id, [])
            if record in wl:
                wl.remove(record)
        except Exception:
            pass
        if self.active_pdf_image is record:
            self.active_pdf_image = None
        self.update_page_layout()

    def _deselect_page_image(self):
        rec = self.active_pdf_image
        if rec is None:
            return
        self._delete_chrome(rec, rec.get("_state", {}).get("img_canvas"))
        self.active_pdf_image = None

    def _page_canvas_click_deselect(self, p_canvas):
        """Click su una zona vuota del foglio: deseleziona l'oggetto immagine."""
        rec = self.active_pdf_image
        if rec is None:
            return
        self._deselect_page_image()

    def _clear_all_images(self):
        for pid in list(self.pages_data.keys()):
            self._clear_image_widgets(pid)
        self.page_images = {k: [] for k in self.pages_data}
        self.page_image_widgets = {k: [] for k in self.pages_data}
        self.active_pdf_image = None

    def on_pdf_canvas_click(self, event):
        if not self.pdf_mode:
            return
        # Click su spazio vuoto del foglio: deseleziona la casella attiva.
        self._hide_all_box_chrome()
        if not self.pdf_add_mode:
            return
        cx = self.pdf_canvas.canvasx(event.x)
        cy = self.pdf_canvas.canvasy(event.y)
        z = self.zoom_factor

        page_id = min(self.pdf_pages_info.keys()) if self.pdf_pages_info else 1
        for pid in sorted(self.pdf_page_offsets.keys()):
            if self.pdf_page_offsets[pid] <= cy:
                page_id = pid
            else:
                break

        y_off = self.pdf_page_offsets.get(page_id, 0)
        info = self.pdf_pages_info[page_id]
        # Ricalcola lo stesso x_off usato in render_all_pdf_pages per centrare le
        # pagine, cosi' il punto cliccato viene convertito nelle coordinate giuste.
        img = Image.open(io.BytesIO(info["img_bytes"]))
        page_w = max(1, int(round(img.width * z)))
        for _pid, _info in self.pdf_pages_info.items():
            _img = Image.open(io.BytesIO(_info["img_bytes"]))
            all_max_w = max(all_max_w, max(1, int(round(_img.width * z))))
        canvas_width = max(self.pdf_canvas.winfo_width(), 1)
        x_off = max(20, (canvas_width - all_max_w) // 2)

        # Scala della pagina effettiva (render_dpi/72), come in render_all_pdf_pages.
        base_scale = info.get("render_dpi", self.pdf_render_dpi) / 72.0
        x_pt = (cx - x_off) / (base_scale * z)
        y_pt = (cy - y_off) / (base_scale * z)

        ov = {
            "x": x_pt, "y": y_pt, "text": "", "size": float(self.font_size.get()),
            "color": 0x000000, "width": 150, "height": 30,
            "custom_font": self.font_family.get(),
            "custom_size": int(self.font_size.get()),
            "custom_weight": "bold" if self.active_styles["bold"] else "normal",
            "custom_slant": "italic" if self.active_styles["italic"] else "roman"
        }
        info["overlays"].append(ov)
        self._create_overlay_entry(page_id, ov, base_scale, y_offset=y_off, x_offset=x_off, z=z)
        self.toggle_add_text_mode()

    def save_as_file(self):
        file_path = filedialog.asksaveasfilename(
            filetypes=[
                ("Tutti i formati supportati", "*.txt *.docx *.doc *.pptx *.rtf *.epub *.odt *.ods *.md *.pdf"),
                ("Documenti Word", "*.docx"),
                ("Documenti PDF", "*.pdf"),
                ("OpenDocument Text", "*.odt"),
                ("OpenDocument Spreadsheet", "*.ods"),
                ("Presentazioni", "*.pptx"),
                ("E-book EPUB", "*.epub"),
                ("Markdown", "*.md"),
                ("Rich Text Format", "*.rtf"),
                ("File di Testo", "*.txt"),
                ("Tutti i file", "*.*")
            ]
        )
        if not file_path:
            return

        ext = os.path.splitext(file_path)[1].lower()
        if not ext:
            file_path += ".docx"

        self.export_universal(file_path)

    def export_universal(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()

        if not self.pdf_mode:
            self.save_current_page_state()

        try:
            if ext in [".pdf", ".pptx", ".epub", ".ods"]:
                self._export_to_pdf_universal(file_path)
            elif ext in [".docx", ".doc", ".odt"]:
                self._export_to_docx_universal(file_path)
            elif ext == ".rtf":
                self._export_to_rtf_universal(file_path)
            elif ext in [".txt", ".md"]:
                self._export_to_txt_universal(file_path)
            else:
                self._export_to_txt_universal(file_path)

            messagebox.showinfo("Esportazione Universale", f"Il file è stato salvato con successo in:\n{file_path}")
        except Exception as e:
            try:
                fallback_path = file_path + ".txt"
                self._export_to_txt_universal(fallback_path)
                messagebox.showwarning("Esportazione di Emergenza", f"Errore ({str(e)}). Salvato come TXT in:\n{fallback_path}")
            except Exception as inner_e:
                messagebox.showerror("Errore Critico", f"Impossibile completare l'esportazione:\n{str(inner_e)}")

    def _export_to_pdf_universal(self, file_path):
        if self.pdf_mode:
            out_doc = pymupdf.open()
            for page_id in sorted(self.pdf_pages_info.keys()):
                info = self.pdf_pages_info[page_id]
                page = out_doc.new_page(width=info["width_pt"], height=info["height_pt"])

                img_tmp_path = os.path.join(os.path.expanduser("~"), f"_export_bg_{page_id}.png")
                with open(img_tmp_path, "wb") as f:
                    f.write(info["img_bytes"])
                page.insert_image(page.rect, filename=img_tmp_path)

                for span in info["spans"]:
                    text = span["text"]
                    if not text.strip():
                        continue
                    try:
                        f_name = "helv"
                        c_font = span.get("custom_font", "").lower()
                        if "times" in c_font or "serif" in c_font: f_name = "times"
                        elif "courier" in c_font or "mono" in c_font: f_name = "courier"

                        is_b = (span.get("custom_weight") == "bold")
                        is_i = (span.get("custom_slant") == "italic")
                        if f_name == "helv":
                            if is_b and is_i: f_name = "helv-boldoblique"
                            elif is_b: f_name = "helv-bold"
                            elif is_i: f_name = "helv-oblique"
                        elif f_name == "times":
                            if is_b and is_i: f_name = "times-bolditalic"
                            elif is_b: f_name = "times-bold"
                            elif is_i: f_name = "times-italic"
                        elif f_name == "courier":
                            if is_b and is_i: f_name = "courier-boldoblique"
                            elif is_b: f_name = "courier-bold"
                            elif is_i: f_name = "courier-oblique"

                        page.insert_text(
                            span["origin"], text,
                            fontsize=float(span.get("custom_size", span["size"])), fontname=f_name,
                            color=self._int_color_to_rgb01(span["color"]),
                        )
                    except Exception:
                        pass

                for ov in info["overlays"]:
                    text = ov["text"]
                    if not text.strip():
                        continue
                    try:
                        f_name = "helv"
                        c_font = ov.get("custom_font", "").lower()
                        if "times" in c_font or "serif" in c_font: f_name = "times"
                        elif "courier" in c_font or "mono" in c_font: f_name = "courier"

                        is_b = (ov.get("custom_weight") == "bold")
                        is_i = (ov.get("custom_slant") == "italic")
                        if f_name == "helv":
                            if is_b and is_i: f_name = "helv-boldoblique"
                            elif is_b: f_name = "helv-bold"
                            elif is_i: f_name = "helv-oblique"
                        elif f_name == "times":
                            if is_b and is_i: f_name = "times-bolditalic"
                            elif is_b: f_name = "times-bold"
                            elif is_i: f_name = "times-italic"
                        elif f_name == "courier":
                            if is_b and is_i: f_name = "courier-boldoblique"
                            elif is_b: f_name = "courier-bold"
                            elif is_i: f_name = "courier-oblique"

                        sz = float(ov.get("size", 12.0))
                        baseline_y = ov["y"] + sz
                        page.insert_text(
                            (ov["x"], baseline_y), text,
                            fontsize=sz, fontname=f_name, color=(0, 0, 0),
                        )
                    except Exception:
                        pass

                if os.path.exists(img_tmp_path):
                    os.remove(img_tmp_path)

            out_doc.save(file_path)
            out_doc.close()
        else:
            pdf_sizes = {"A4": A4, "A3": A3, "A2": A2}
            chosen_size = pdf_sizes.get(self.current_format, A4)
            if self.current_orientation == "Orizzontale":
                from reportlab.lib.pagesizes import landscape
                chosen_size = landscape(chosen_size)

            doc = SimpleDocTemplate(file_path, pagesize=chosen_size)
            styles = getSampleStyleSheet()
            custom_style = ParagraphStyle('Univ_Style', parent=styles['Normal'], fontName='Helvetica', fontSize=self.font_size.get(), leading=self.font_size.get() + 4)

            story = []
            for p_id in sorted(self.pages_data.keys()):
                editor = self.text_page_editors.get(p_id)
                if editor is None:
                    styled_lines = [(line, custom_style) for line in self.pages_data[p_id].split("\n")]
                else:
                    line_runs = [[]]
                    line_styles = [None]
                    for segment, fmt in self._iter_editor_style_segments(editor):
                        chunks = segment.split("\n")
                        for chunk_index, chunk in enumerate(chunks):
                            if chunk:
                                family = fmt["family"].lower()
                                face = "Times-Roman" if ("times" in family or "serif" in family) else ("Courier" if ("courier" in family or "mono" in family) else "Helvetica")
                                if fmt["bold"] and fmt["italic"]:
                                    face = {"Helvetica":"Helvetica-BoldOblique","Times-Roman":"Times-BoldItalic","Courier":"Courier-BoldOblique"}[face]
                                elif fmt["bold"]:
                                    face = {"Helvetica":"Helvetica-Bold","Times-Roman":"Times-Bold","Courier":"Courier-Bold"}[face]
                                elif fmt["italic"]:
                                    face = {"Helvetica":"Helvetica-Oblique","Times-Roman":"Times-Italic","Courier":"Courier-Oblique"}[face]
                                safe = html_lib.escape(chunk)
                                if fmt["underline"]:
                                    safe = f"<u>{safe}</u>"
                                safe = f'<font face="{face}" size="{fmt["size"]}" color="{fmt["color"]}">{safe}</font>'
                                if fmt["background"]:
                                    safe = f'<span backColor="{fmt["background"]}">{safe}</span>'
                                line_runs[-1].append(safe)
                                line_styles[-1] = fmt
                            if chunk_index < len(chunks) - 1:
                                line_runs.append([])
                                line_styles.append(None)
                    styled_lines = []
                    for runs, fmt in zip(line_runs, line_styles):
                        if fmt:
                            align = {"left": 0, "center": 1, "right": 2}.get(fmt["align"], 0)
                            para_style = ParagraphStyle(
                                f"Univ_Style_{p_id}_{len(styled_lines)}", parent=custom_style,
                                alignment=align, fontSize=fmt["size"], leading=fmt["size"] + 4,
                            )
                        else:
                            para_style = custom_style
                        styled_lines.append(("".join(runs), para_style))
                for line, para_style in styled_lines:
                    if not line.strip():
                        story.append(Spacer(1, 12))
                    else:
                        story.append(Paragraph(line, para_style))
                if p_id != max(self.pages_data.keys()):
                    story.append(PageBreak())
            doc.build(story)

    def _export_to_docx_universal(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        doc = Document()
        if self.pdf_mode:
            for page_id in sorted(self.pdf_pages_info.keys()):
                info = self.pdf_pages_info[page_id]
                p_text_parts = [span["text"] for span in info["spans"] if span["text"].strip()]
                for ov in info["overlays"]:
                    if ov["text"].strip():
                        p_text_parts.append(ov["text"])
                doc.add_paragraph("\n".join(p_text_parts))
                if page_id != max(self.pdf_pages_info.keys()):
                    doc.add_page_break()
        else:
            page_ids = sorted(self.pages_data.keys())
            for page_index, p_id in enumerate(page_ids):
                editor = self.text_page_editors.get(p_id)
                if editor is not None:
                    self._append_styled_page_to_docx(doc, editor)
                else:
                    for line in self.pages_data[p_id].split("\n"):
                        doc.add_paragraph(line)
                if page_index < len(page_ids) - 1:
                    doc.add_page_break()
        # python-docx sa scrivere SOLO .docx: per .odt/.doc il file va convertito
        # con LibreOffice, altrimenti verrebbe salvato come docx con un'estensione
        # sbagliata (per .odt il risultato era un documento corrotto/cancellato).
        if ext in (".odt", ".doc"):
            import tempfile, shutil
            tmpdir = tempfile.mkdtemp(prefix="simplewright_exp_")
            tmp_docx = os.path.join(tmpdir, os.path.basename(os.path.splitext(file_path)[0]) + ".docx")
            doc.save(tmp_docx)
            soffice = shutil.which("soffice") or shutil.which("libreoffice")
            if soffice:
                try:
                    proc = subprocess.run(
                        [soffice, "--headless", "--convert-to", ext.lstrip("."),
                         "--outdir", os.path.dirname(file_path) or ".", tmp_docx],
                        capture_output=True, timeout=120
                    )
                    produced = os.path.join(
                        os.path.dirname(file_path) or ".",
                        os.path.splitext(os.path.basename(tmp_docx))[0] + ext
                    )
                    if proc.returncode == 0 and os.path.exists(produced):
                        if os.path.abspath(produced) != os.path.abspath(file_path):
                            shutil.move(produced, file_path)
                finally:
                    shutil.rmtree(tmpdir, ignore_errors=True)
            else:
                # senza LibreOffice non possiamo produrre un .odt/.doc vero:
                # salviamo comunque un .docx nella posizione richiesta
                shutil.copy(tmp_docx, file_path)
                shutil.rmtree(tmpdir, ignore_errors=True)
            return
        doc.save(file_path)

    def _iter_editor_style_segments(self, editor):
        """Yield contiguous (text, style) segments using Tk tag boundaries."""
        text = editor.get("1.0", "end-1c")
        events = {}
        relevant_tags = [tag for tag in editor.tag_names() if self._is_relevant_tag(tag)]
        for tag in relevant_tags:
            ranges = editor.tag_ranges(tag)
            for i in range(0, len(ranges), 2):
                start = self._index_to_char_offset(editor, ranges[i])
                end = self._index_to_char_offset(editor, ranges[i + 1])
                if 0 <= start < end <= len(text):
                    events.setdefault(start, []).append((1, tag))
                    events.setdefault(end, []).append((0, tag))

        active = set()
        points = sorted(set((0, len(text), *events.keys())))
        base_family = self.font_family.get() or "Arial"
        try:
            base_size = int(self.default_font_size or 12)
        except Exception:
            base_size = 12
        for i, start in enumerate(points[:-1]):
            for is_start, tag in events.get(start, ()):
                if not is_start:
                    active.discard(tag)
            for is_start, tag in events.get(start, ()):
                if is_start:
                    active.add(tag)
            end = points[i + 1]
            if end <= start:
                continue
            family, size, bold, italic = base_family, base_size, False, False
            color = "#000000"
            background = None
            underline = False
            alignment = "left"
            for tag in relevant_tags:
                if tag not in active:
                    continue
                if tag in self.font_tag_defs:
                    family, size, bold, italic = self.font_tag_defs[tag]
                elif tag.startswith("rtf_color_"):
                    color = "#" + tag[len("rtf_color_"):]
                elif tag.startswith("color_"):
                    color = "#" + tag[len("color_"):]
                elif tag.startswith("rtf_bg_"):
                    background = "#" + tag[len("rtf_bg_"):]
                elif tag.startswith("bg_"):
                    background = "#" + tag[len("bg_"):]
                elif tag in ("underline", "rtf_underline"):
                    underline = True
                elif tag in ("left", "center", "right", "rtf_left", "rtf_center", "rtf_right"):
                    alignment = tag.replace("rtf_", "")
            yield text[start:end], {
                "family": family, "size": max(1, int(round(float(size)))),
                "bold": bool(bold), "italic": bool(italic),
                "underline": underline, "color": color.upper(),
                "background": background.upper() if background else None,
                "align": alignment,
            }

    def _append_styled_page_to_docx(self, doc, editor):
        """Append editor content to DOCX while preserving inline formatting."""
        paragraph = doc.add_paragraph()
        for segment, style in self._iter_editor_style_segments(editor):
            try:
                paragraph.alignment = {
                    "center": WD_ALIGN_PARAGRAPH.CENTER,
                    "right": WD_ALIGN_PARAGRAPH.RIGHT,
                    "left": WD_ALIGN_PARAGRAPH.LEFT,
                }.get(style["align"], WD_ALIGN_PARAGRAPH.LEFT)
            except Exception:
                pass
            pieces = segment.split("\n")
            for piece_index, piece in enumerate(pieces):
                if piece:
                    run = paragraph.add_run(piece)
                    run.font.name = style["family"]
                    run.font.size = Pt(style["size"])
                    run.bold = style["bold"]
                    run.italic = style["italic"]
                    run.underline = style["underline"]
                    try:
                        run.font.color.rgb = RGBColor.from_string(style["color"].lstrip("#"))
                    except Exception:
                        pass
                    if style["background"]:
                        try:
                            shading = OxmlElement("w:shd")
                            shading.set(qn("w:fill"), style["background"].lstrip("#"))
                            run._element.get_or_add_rPr().append(shading)
                        except Exception:
                            pass
                if piece_index < len(pieces) - 1:
                    paragraph = doc.add_paragraph()

    def _export_to_rtf_universal(self, file_path):
        """Esporta l'editor in un RTF reale, conservando font, dimensioni,
        grassetto, corsivo, sottolineato e colori."""
        self.save_current_page_state()
        page_segments = []
        fonts = []
        colors = []
        for pid in sorted(self.text_page_editors):
            segments = list(self._iter_editor_style_segments(self.text_page_editors[pid]))
            page_segments.append(segments)
            for _, style in segments:
                if style["family"] not in fonts:
                    fonts.append(style["family"])
                color = style["color"].upper()
                if color != "#000000" and color not in colors:
                    colors.append(color)
                background = style["background"]
                if background and background not in colors:
                    colors.append(background)
        if not fonts:
            fonts = [self.font_family.get() or "Arial"]
        font_index = {family: i for i, family in enumerate(fonts)}
        color_index = {"#000000": 0}
        color_index.update({color: i for i, color in enumerate(colors, 1)})
        out = ["{\\rtf1\\ansi\\deff0{\\fonttbl"]
        for i, family in enumerate(fonts):
            safe_family = family.replace("\\", "").replace("{", "").replace("}", "")
            out.append(f"{{\\f{i} {safe_family};}}")
        out.append("}{\\colortbl ;")
        for color in colors:
            rgb = color.lstrip("#")
            out.append(f"\\red{int(rgb[0:2], 16)}\\green{int(rgb[2:4], 16)}\\blue{int(rgb[4:6], 16)};")
        out.append("}")

        previous_style = None
        page_ids = sorted(self.text_page_editors)
        for page_index, segments in enumerate(page_segments):
            for segment, style in segments:
                style_key = (style["family"], style["size"], style["bold"],
                             style["italic"], style["underline"], style["color"],
                             style["background"], style["align"])
                for char in segment:
                    if char == "\n":
                        align_control = {"center": "\\qc", "right": "\\qr", "left": "\\ql"}.get(style["align"], "\\ql")
                        out.append(align_control + "\\par\n")
                        previous_style = None
                        continue
                    if style_key != previous_style:
                        align_control = {"center": "\\qc", "right": "\\qr", "left": "\\ql"}.get(style["align"], "\\ql")
                        bold_control = "\\b" if style["bold"] else "\\b0"
                        italic_control = "\\i" if style["italic"] else "\\i0"
                        underline_control = "\\ul" if style["underline"] else "\\ul0"
                        highlight_control = "\\highlight%d" % color_index.get(style["background"], 0)
                        out.append(
                            f"{align_control}\\f{font_index.get(style['family'], 0)}"
                            f"\\fs{max(1, style['size'] * 2)}"
                            f"{bold_control}{italic_control}{underline_control}{highlight_control}"
                            f"\\cf{color_index.get(style['color'], 0)} "
                        )
                        previous_style = style_key
                    if char == "\t":
                        out.append("\\tab ")
                    elif char in "{}\\":
                        out.append("\\" + char)
                    elif ord(char) > 127:
                        codepoint = ord(char)
                        if codepoint <= 0xFFFF:
                            signed = codepoint if codepoint < 0x8000 else codepoint - 0x10000
                            out.append(f"\\u{signed}?")
                        else:
                            value = codepoint - 0x10000
                            high = 0xD800 + (value >> 10)
                            low = 0xDC00 + (value & 0x3FF)
                            out.append(f"\\u{high - 0x10000}?\\u{low - 0x10000}?")
                    else:
                        out.append(char)
            if page_index < len(page_ids) - 1:
                out.append("\\page ")
                previous_style = None
        out.append("}")
        Path(file_path).write_text("".join(out), encoding="ascii", errors="strict")

    def _export_to_txt_universal(self, file_path):
        all_text_accumulated = []
        if self.pdf_mode:
            for page_id in sorted(self.pdf_pages_info.keys()):
                info = self.pdf_pages_info[page_id]
                page_lines = [span["text"] for span in info["spans"] if span["text"].strip()]
                for ov in info["overlays"]:
                    if ov["text"].strip():
                        page_lines.append(ov["text"])
                all_text_accumulated.append("\n".join(page_lines))
        else:
            for p_id in sorted(self.pages_data.keys()):
                all_text_accumulated.append(self.pages_data[p_id])

        with open(file_path, "w", encoding="utf-8") as f:
            f.write("\n\n--- [INTERRUZIONE PAGINA] ---\n\n".join(all_text_accumulated))

    def print_document(self):
        self.save_current_page_state()
        temp_pdf = os.path.join(os.path.expanduser("~"), "UniversalWriter_PrintJob.pdf")
        try:
            self._export_to_pdf_universal(temp_pdf)
            system_os = platform.system().lower()
            if "windows" in system_os:
                os.startfile(temp_pdf, "print")
            elif "linux" in system_os:
                try: subprocess.run(["xdg-open", temp_pdf], check=True)
                except: subprocess.run(["lp", temp_pdf], check=True)
            else:
                subprocess.run(["open", "-a", "Preview", temp_pdf], check=True)
        except Exception as e:
            messagebox.showerror("Errore di Stampa", str(e))

    def on_layout_change(self, event=None):
        self.current_format = self.format_cb.get()
        self.current_orientation = self.orientation_cb.get()
        self.update_page_layout()

    def on_zoom_change(self, event=None):
        self.zoom_factor = float(self.scale_zoom.get())
        self.lbl_zoom_val.config(text=f"{int(self.zoom_factor * 100)}%")
        self.update_page_layout()

    def on_margin_change(self, event=None):
        if self.pdf_mode:
            return

        base_w, base_h = self.base_page_sizes[self.current_format]
        w = int(base_w * self.zoom_factor)
        h = int(base_h * self.zoom_factor)
        if self.current_orientation == "Orizzontale":
            w, h = h, w

        m_left = int(self.scale_m_left.get())
        m_right = int(self.scale_m_right.get())
        m_top = int(self.scale_m_top.get())
        m_bottom = int(self.scale_m_bottom.get())

        min_writable = 40
        max_h_margin = max(0, h - min_writable)
        max_w_margin = max(0, w - min_writable)

        if (m_top + m_bottom) > max_h_margin and (m_top + m_bottom) > 0:
            factor = max_h_margin / (m_top + m_bottom)
            m_top, m_bottom = int(m_top * factor), int(m_bottom * factor)
        if (m_left + m_right) > max_w_margin and (m_left + m_right) > 0:
            factor = max_w_margin / (m_left + m_right)
            m_left, m_right = int(m_left * factor), int(m_right * factor)

        avail_w = max(min_writable, w - m_left - m_right)
        avail_h = max(min_writable, h - m_top - m_bottom)

        for pid, ed in self.text_page_editors.items():
            ed.place_configure(x=m_left, y=m_top, width=avail_w, height=avail_h)

    def setup_tags_for_editor(self, editor):
        editor.tag_configure("bold", font=(self.font_family.get(), max(1, int(self.default_font_size * self.zoom_factor)), "bold") if hasattr(self, 'font_family') else None)
        try:
            editor.tag_configure("bold", font=("", 0, "bold"))
        except Exception:
            pass
        try:
            editor.tag_configure("italic", font=("", 0, "italic"))
        except Exception:
            pass
        try:
            editor.tag_configure("underline", underline=True)
        except Exception:
            pass
        editor.tag_configure("left", justify=tk.LEFT)
        editor.tag_configure("center", justify=tk.CENTER)
        editor.tag_configure("right", justify=tk.RIGHT)

    def check_page_overflow(self, event, editor, page_id):
        nav_keys = {
            "BackSpace", "Delete", "Left", "Right", "Up", "Down",
            "Home", "End", "Prior", "Next",
            "Shift_L", "Shift_R", "Control_L", "Control_R",
            "Alt_L", "Alt_R", "Caps_Lock", "Tab",
        }
        if event.keysym in nav_keys:
            return

        self.root.update_idletasks()
        try:
            editor.update_idletasks()
            bbox = editor.bbox("end-1c")
            if bbox is None:
                return

            text_bottom = bbox[1] + bbox[3]
            editor_height = editor.winfo_height()

            if text_bottom > editor_height - 2:
                content = editor.get("1.0", "end-1c")
                all_lines = content.split("\n")

                insert_idx = editor.index("insert")
                insert_line = int(insert_idx.split(".")[0]) - 1

                if len(all_lines) > 0:
                    overflow_start = max(1, insert_line)
                    overflow_text = "\n".join(all_lines[overflow_start:])
                    keep_text = "\n".join(all_lines[:overflow_start])

                    new_page_id = self.insert_page_after(page_id)
                    self.pages_data[new_page_id] = overflow_text
                    editor.delete(f"{overflow_start + 1}.0", "end")
                    editor.mark_set("insert", "end-1c")

                    self.root.after_idle(lambda: self._refresh_after_overflow(new_page_id))
        except Exception:
            pass

    def _refresh_after_overflow(self, new_page_id):
        self.save_current_page_state()
        self.update_page_layout()
        self.switch_to_page(new_page_id)

    def update_page_layout(self):
        if self.pdf_mode:
            self.render_all_pdf_pages()
            return

        self.save_current_page_state()
        # Se il testo eccede l'altezza della pagina, spezza automaticamente in
        # piu' pagine PRIMA di ridisegnare gli editor.
        self._auto_paginate_text_pages()
        self.canvas_desktop.delete("text_page_win")

        base_w, base_h = self.base_page_sizes[self.current_format]
        w = int(base_w * self.zoom_factor)
        h = int(base_h * self.zoom_factor)
        if self.current_orientation == "Orizzontale":
            w, h = h, w

        m_left = int(self.scale_m_left.get())
        m_right = int(self.scale_m_right.get())
        m_top = int(self.scale_m_top.get())
        m_bottom = int(self.scale_m_bottom.get())
        avail_w = max(40, w - m_left - m_right)
        avail_h = max(40, h - m_top - m_bottom)

        gap = 30
        y_offset = 30
        self.pdf_page_offsets = {}

        self.text_page_frames = {}
        self.text_page_editors = {}

        for p_id in sorted(self.pages_data.keys()):
            p_canvas = tk.Canvas(self.canvas_desktop, bg="#ffffff", bd=0, highlightthickness=0, width=w, height=h)

            p_canvas.create_line(0, 0, w, 0, fill="#999999")
            p_canvas.create_line(0, 0, 0, h, fill="#999999")
            p_canvas.create_line(w, 0, w, h, fill="#999999")
            p_canvas.create_line(0, h, w, h, fill="#999999", dash=(6, 4))

            page_bg = "#ffffff"
            p_frame = tk.Frame(p_canvas, bg=page_bg, bd=0)
            p_canvas.create_window((0, 0), window=p_frame, anchor="nw", width=w, height=h,
                                   tags=("page_content",))

            current_f_size = int(self.default_font_size * self.zoom_factor)
            editor = tk.Text(
                p_frame, wrap=tk.WORD,
                font=(self.font_family.get(), max(1, current_f_size)),
                undo=True, maxundo=-1, bd=0, highlightthickness=0,
                bg=page_bg,
                padx=0, pady=0, spacing1=0, spacing2=0, spacing3=0,
            )
            editor.place(x=m_left, y=m_top, width=avail_w, height=avail_h)

            content = self.pages_data.get(p_id, "")
            editor.insert("1.0", content)

            self.setup_tags_for_editor(editor)
            self._restore_page_format(p_id, editor)

            editor.bind("<Key>", lambda e, ed=editor, pid=p_id: self.check_page_overflow(e, ed, pid), add="+")
            editor.bind("<KeyRelease>", lambda e, pid=p_id: self.on_text_key_release(e, pid), add="+")
            editor.bind("<KeyRelease>", lambda e, pid=p_id: self._schedule_auto_links(pid), add="+")
            editor.bind("<ButtonRelease-1>", lambda e, pid=p_id: self._sync_toolbar_to_cursor(pid))
            editor.bind("<ButtonRelease-1>", lambda _e, ed=editor: self._update_saved_format_selection(ed), add="+")
            editor.bind("<KeyRelease>", lambda _e, ed=editor: self._update_saved_format_selection(ed), add="+")
            editor.bind("<Button-1>", lambda e, pid=p_id: self.on_editor_click_for_textbox(e, pid))
            editor.bind("<Button-1>", self._deselect_text_box, add="+")
            editor.bind("<FocusIn>", lambda e, pid=p_id: [setattr(self, "active_page_id", pid), setattr(self, "active_pdf_box", None), self._deselect_page_image()], add="+")

            editor.bind("<Button-4>", self.handle_mouse_wheel)
            editor.bind("<Button-5>", self.handle_mouse_wheel)
            editor.bind("<MouseWheel>", self.handle_mouse_wheel)

            self.text_page_frames[p_id] = p_frame
            self.text_page_editors[p_id] = editor
            self._apply_auto_links(editor)

            self.pdf_page_offsets[p_id] = y_offset
            self.canvas_desktop.create_window((50, y_offset), window=p_canvas, anchor="nw", tags="text_page_win")

            # Ordine originale e affidabile dei widget: immagini visibili sopra
            # il testo della pagina, caselle mobili in primo piano.
            self._render_page_images(p_id, p_canvas)
            p_canvas.bind("<Button-1>", lambda e, cv=p_canvas: self._page_canvas_click_deselect(cv))

            self._render_text_boxes_for_page(p_id)
            y_offset += h + gap

        total_h = y_offset
        self.canvas_desktop.config(scrollregion=(0, 0, w + 100, total_h))
        self.center_page_on_desktop()
        self.update_pages_sidebar()

    def center_page_on_desktop(self):
        if self.pdf_mode:
            # In modalita' PDF le pagine sono rese (e centrate) da
            # render_all_pdf_pages: qui basta ri-scatenare il render su resize
            # della finestra, con un piccolo debounce per non ri-renderizzare a
            # ogni pixel durante il ridimensionamento dal vivo.
            if getattr(self, "_pdf_center_job", None) is not None:
                self.root.after_cancel(self._pdf_center_job)
            self._pdf_center_job = self.root.after(120, self.render_all_pdf_pages)
            return

        canvas_width = self.canvas_desktop.winfo_width()
        base_w, _ = self.base_page_sizes[self.current_format]
        w = int(base_w * self.zoom_factor)
        if self.current_orientation == "Orizzontale":
            _, base_h = self.base_page_sizes[self.current_format]
            w = int(base_h * self.zoom_factor)

        new_x = max(20, (canvas_width - w) // 2)
        for item in self.canvas_desktop.find_withtag("text_page_win"):
            y = self.canvas_desktop.coords(item)[1]
            self.canvas_desktop.coords(item, new_x, y)

    def toggle_color_popup(self, popup_type):
        if self.current_popup and self.current_popup.winfo_exists():
            same_type = (self.active_popup_type == popup_type)
            self.close_popup()
            if same_type: return
        if time.time() - self.last_popup_closed_time < 0.15: return
        self.open_color_popup(popup_type)

    def open_color_popup(self, popup_type):
        self.active_popup_type = popup_type
        self.current_popup = tk.Toplevel(self.root)
        self.current_popup.wm_overrideredirect(True)
        self.current_popup.config(bd=2, relief="ridge")
        target_btn = self.main_color_btn if popup_type == "text" else self.main_bg_btn
        x, y = target_btn.winfo_rootx(), target_btn.winfo_rooty() + target_btn.winfo_height()
        self.current_popup.geometry(f"+{x}+{y}")
        self.current_popup.bind("<FocusOut>", lambda e: self.close_popup())
        self.current_popup.focus_set()

        paint_matrix = [
            ["#ffffff", "#f0f0f0", "#e0e0e0", "#d0d0d0", "#c0c0c0", "#a0a0a0", "#808080", "#606060", "#404040", "#202020"],
            ["#ffcccc", "#ff9999", "#ff6666", "#ff3333", "#ff0000", "#cc0000", "#990000", "#660000", "#330000", "#1a0000"],
            ["#ffddcc", "#ffbb99", "#ff9966", "#ff7733", "#ff5500", "#cc4400", "#993300", "#662200", "#331100", "#000000"],
            ["#ffffcc", "#ffff99", "#ffff66", "#ffff33", "#ffff00", "#cccc00", "#999900", "#666600", "#333300", "#1a1a00"],
            ["#ccffcc", "#99ff99", "#66ff66", "#33ff33", "#00ff00", "#00cc00", "#009900", "#006600", "#003300", "#001a00"],
            ["#ccffff", "#99ffff", "#66ffff", "#33ffff", "#00ffff", "#00cccc", "#009999", "#006666", "#003333", "#001a1a"],
            ["#e5ccff", "#cc99ff", "#b266ff", "#9933ff", "#7f00ff", "#6600cc", "#4c0099", "#330066", "#190033", "#0d001a"]
        ]

        callback = self.set_text_color if popup_type == "text" else self.set_text_background_color
        for r_idx, row in enumerate(paint_matrix):
            for c_idx, color in enumerate(row):
                btn = tk.Button(self.current_popup, bg=color, image=self.pixel_img, width=18, height=18, command=lambda c=color: [callback(c), self.close_popup()])
                btn.grid(row=r_idx, column=c_idx, padx=1, pady=1)

    def close_popup(self):
        if self.current_popup:
            try: self.current_popup.destroy()
            except tk.TclError: pass
            self.current_popup = None
            self.active_popup_type = None
            self.last_popup_closed_time = time.time()

    def apply_theme(self):
        if self.is_dark:
            self.canvas_desktop.config(bg="#2d2d2d")
            for ed in self.text_page_editors.values():
                ed.config(bg="#ffffff", fg="#000000", insertbackground="black")
        else:
            self.canvas_desktop.config(bg="#e0e0e0")
            for ed in self.text_page_editors.values():
                ed.config(bg="#ffffff", fg="#000000", insertbackground="black")

    def toggle_dark_mode(self):
        self.is_dark = not self.is_dark
        self.apply_theme()
        self.save_config()

    def _get_font_tag(self, editor, family, size, bold, italic):
        key = (family, int(size), bool(bold), bool(italic))
        tag_name = self._font_tag_map.get(key)
        if not tag_name:
            self._font_tag_counter += 1
            tag_name = f"fnt{self._font_tag_counter}"
            self._font_tag_map[key] = tag_name
            self.font_tag_defs[tag_name] = key
        weight = "bold" if bold else "normal"
        slant = "italic" if italic else "roman"
        px = max(1, int(int(size) * self.zoom_factor))
        try:
            editor.tag_configure(tag_name, font=(family, px, f"{weight} {slant}".strip()))
        except Exception:
            pass
        return tag_name

    def apply_font_state_to_range(self, editor, start, end, family, size, bold, italic):
        for t in list(self.font_tag_defs.keys()):
            editor.tag_remove(t, start, end)
        tag_name = self._get_font_tag(editor, family, size, bold, italic)
        editor.tag_add(tag_name, start, end)

    def get_format_at(self, editor, index):
        try:
            tags = editor.tag_names(index)
        except tk.TclError:
            tags = []
        for t in tags:
            if t in self.font_tag_defs:
                return self.font_tag_defs[t]
        fam = self.font_family.get()
        try:
            sz = self.default_font_size
        except Exception:
            sz = 12
        return (fam, sz, False, False)

    def _update_style_buttons(self):
        self.btn_bold.state(["pressed"] if self.active_styles["bold"] else ["!pressed"])
        self.btn_italic.state(["pressed"] if self.active_styles["italic"] else ["!pressed"])
        self.btn_underline.state(["pressed"] if self.active_styles["underline"] else ["!pressed"])

    def _sync_toolbar_to_cursor(self, page_id=None):
        if self.pdf_mode:
            return
        if page_id is None:
            page_id = self.active_page_id
        ed = self.text_page_editors.get(page_id)
        if not ed: return

        try:
            start_sel = ed.index("sel.first")
            idx = start_sel
        except tk.TclError:
            return

        fam, sz, bold, italic = self.get_format_at(ed, idx)
        try:
            underline = "underline" in ed.tag_names(idx)
        except tk.TclError:
            underline = False

        self.font_family.set(fam)
        self.font_size.set(sz)
        self.active_styles["bold"] = bold
        self.active_styles["italic"] = italic
        self.active_styles["underline"] = underline
        self._update_style_buttons()

    def select_recent_slot(self, idx):
        for btn in self.recent_buttons: btn.config(relief="flat", bd=1)
        self.selected_slot_idx = idx
        self.recent_buttons[idx].config(relief="sunken", bd=2)
        if self.recent_colors[idx] != "#ffffff": self.apply_color_to_state(self.recent_colors[idx])

    def set_text_color(self, color):
        if not color: return
        self.recent_colors[self.selected_slot_idx] = color
        self.recent_buttons[self.selected_slot_idx].config(bg=color)
        self.apply_color_to_state(color)

    def apply_color_to_state(self, color):
        self.current_color = color
        self.main_color_btn.config(bg=color)
        if self.active_pdf_box is not None:
            try:
                if getattr(self.active_pdf_box, "_rich", False):
                    self._apply_rich_format(self.active_pdf_box, color=color)
                else:
                    self.active_pdf_box.config(fg=color)
                    self.active_pdf_box.data_obj["color"] = int(color.replace("#", ""), 16)
            except Exception:
                pass
            return

        ed = self.text_page_editors.get(self.active_page_id)
        if not ed: return

        try:
            start, end = ed.index("sel.first"), ed.index("sel.last")
            tag_name = f"color_{color.replace('#', '')}"
            ed.tag_configure(tag_name, foreground=color)
            ed.tag_add(tag_name, start, end)
        except tk.TclError:
            ed.config(fg=color)

    def set_text_background_color(self, color):
        if not color: return
        self.current_bg_color = color
        if self.active_pdf_box is not None:
            try:
                t = self.active_pdf_box
                if getattr(t, "_rich", False) and getattr(t, "container", None) is not None:
                    t.container.configure(bg=color)
                    t.configure(bg=color)
                    t.data_obj["background"] = color
                    return
            except Exception:
                pass
        ed = self.text_page_editors.get(self.active_page_id)
        if not ed: return
        try:
            start, end = ed.index("sel.first"), ed.index("sel.last")
        except tk.TclError:
            return
        tag_name = f"bg_{color.replace('#', '')}"
        ed.tag_configure(tag_name, background=color)
        ed.tag_add(tag_name, start, end)

    def _update_saved_format_selection(self, editor):
        try:
            editor._format_selection = (editor.index("sel.first"), editor.index("sel.last"))
        except tk.TclError:
            editor._format_selection = None

    def _remember_format_selection(self, event=None):
        """Keep the current selection when a toolbar combobox takes focus."""
        editor = self.active_pdf_box if self.active_pdf_box is not None else None
        if editor is None:
            editor = self.text_page_editors.get(self.active_page_id)
        if editor is None:
            return
        try:
            editor._format_selection = (editor.index("sel.first"), editor.index("sel.last"))
        except tk.TclError:
            # Do not replace a saved range here: the toolbar may already have
            # taken focus and cleared Tk's visible selection.
            pass

    def _get_format_selection(self, editor):
        try:
            return editor.index("sel.first"), editor.index("sel.last")
        except tk.TclError:
            selection = getattr(editor, "_format_selection", None)
            if selection:
                try:
                    if editor.compare(selection[0], "<", selection[1]):
                        return selection
                except Exception:
                    pass
        return None

    def change_font(self, event=None):
        if self.active_pdf_box is not None:
            if getattr(self.active_pdf_box, "_rich", False):
                textw = self.active_pdf_box
                fam = self.font_family.get()
                try:
                    sz = int(self.font_size.get())
                except Exception:
                    sz = 12
                selection = self._get_format_selection(textw)
                if selection:
                    try:
                        textw.tag_add("sel", selection[0], selection[1])
                    except Exception:
                        pass
                    # Selezione attiva: la grandezza cambia SOLO la selezione
                    # (stesso comportamento del testo delle pagine).
                    self._apply_rich_format(textw, family=fam, size_px=sz)
                    textw._format_selection = None
                    return
                underlined = self._rich_underlined_runs(textw)
                if underlined:
                    # Senza selezione, se c'e' testo sottolineato cambia solo
                    # quello (stesso comportamento del testo delle pagine).
                    self._apply_rich_format(textw, family=fam, size_px=sz,
                                            target_runs=underlined)
                    return
                # Default per le caselle: la dimensione si applica a tutta la casella.
                self._apply_rich_format(textw,
                                        family=fam,
                                        size_px=sz,
                                        target_runs=self._box_all_runs(textw))
                return
            try:
                fam = self.font_family.get()
                sz = int(self.font_size.get())
                self.active_pdf_box.box_font.config(family=fam, size=sz)
                self.active_pdf_box.data_obj["custom_font"] = fam
                self.active_pdf_box.data_obj["custom_size"] = sz
            except Exception:
                pass
            return

        fam = self.font_family.get()
        try:
            sz = int(self.font_size.get())
        except Exception:
            sz = 12

        ed = self.text_page_editors.get(self.active_page_id)
        if ed:
            selection = self._get_format_selection(ed)
            if selection:
                start, end = selection
                _, _, bold, italic = self.get_format_at(ed, start)
                self.apply_font_state_to_range(ed, start, end, fam, sz, bold, italic)
                ed._format_selection = None
            else:
                # Nessuna selezione: applica la nuova grandezza/font SOLO al
                # testo sottolineato della pagina corrente, conservando
                # grassetto/corsivo dei singoli tratti; il resto resta invariato.
                self._apply_size_to_underlined(ed, fam, sz)
                self.update_page_layout()
                return

        self.update_page_layout()

    def _apply_size_to_underlined(self, editor, fam, sz):
        """Applica la nuova grandezza/font solo ai tratti sottolineati,
        preservando grassetto/corsivo gia' presenti in ogni sottotag font."""
        if editor is None:
            return
        underline_ranges = []
        for tag in ("underline", "rtf_underline"):
            try:
                ranges = editor.tag_ranges(tag)
                underline_ranges.extend((ranges[i], ranges[i + 1]) for i in range(0, len(ranges), 2))
            except Exception:
                continue
        for s, e in underline_ranges:
            # tratti gia' formattati (tag font) dentro [s,e): tengono i propri
            # attributi bold/italic
            pieces = []
            for t in list(self.font_tag_defs.keys()):
                try:
                    tr = editor.tag_ranges(t)
                except Exception:
                    continue
                for j in range(0, len(tr), 2):
                    a, b = tr[j], tr[j + 1]
                    if editor.compare(b, "<=", s) or editor.compare(e, "<=", a):
                        continue
                    na = a if editor.compare(a, ">", s) else s
                    nb = b if editor.compare(b, "<", e) else e
                    if editor.compare(na, "<", nb):
                        _, _, bold, italic = self.get_format_at(editor, na)
                        pieces.append([na, nb, bold, italic])
            out = []
            if pieces:
                pieces.sort(key=lambda p: editor.index(p[0]))
                joined = [list(pieces[0])]
                for na, nb, bo, it in pieces[1:]:
                    if editor.compare(na, "<=", joined[-1][1]):
                        if editor.compare(nb, ">", joined[-1][1]):
                            joined[-1][1] = nb
                    else:
                        joined.append([na, nb, bo, it])
                ptr = s
                for ca, cb, bo, it in joined:
                    if editor.compare(ptr, "<", ca):
                        out.append((ptr, ca, False, False))
                    out.append((ca, cb, bo, it))
                    ptr = cb
                if editor.compare(ptr, "<", e):
                    out.append((ptr, e, False, False))
            else:
                out.append((s, e, False, False))
            for a, b, bo, it in out:
                self.apply_font_state_to_range(editor, a, b, fam, sz, bo, it)

    def toggle_style(self, style):
        if self.active_pdf_box is not None:
            try:
                if getattr(self.active_pdf_box, "_rich", False):
                    cur = self._sync_toolbar_from_rich(self.active_pdf_box) or None
                    if style == "bold":
                        want = not self.active_styles["bold"]
                        self._apply_rich_format(self.active_pdf_box, bold=want)
                    elif style == "italic":
                        want = not self.active_styles["italic"]
                        self._apply_rich_format(self.active_pdf_box, italic=want)
                    elif style == "underline":
                        want = not self.active_styles["underline"]
                        self._apply_rich_format(self.active_pdf_box, underline=want)
                    return
                current_w = self.active_pdf_box.box_font.actual()["weight"]
                current_s = self.active_pdf_box.box_font.actual()["slant"]

                if style == "bold":
                    new_w = "normal" if current_w == "bold" else "bold"
                    self.active_pdf_box.box_font.config(weight=new_w)
                    self.active_pdf_box.data_obj["custom_weight"] = new_w
                    self.active_styles["bold"] = (new_w == "bold")
                elif style == "italic":
                    new_s = "roman" if current_s == "italic" else "italic"
                    self.active_pdf_box.box_font.config(slant=new_s)
                    self.active_pdf_box.data_obj["custom_slant"] = new_s
                    self.active_styles["italic"] = (new_s == "italic")
                self._update_style_buttons()
            except Exception as e:
                pass
            return

        ed = self.text_page_editors.get(self.active_page_id)
        if not ed: return

        fam = self.font_family.get()
        try:
            sz = int(self.font_size.get())
        except Exception:
            sz = 12

        self.active_styles[style] = not self.active_styles[style]

        try:
            start, end = ed.index("sel.first"), ed.index("sel.last")
            has_sel = True
        except tk.TclError:
            has_sel = False

        if has_sel:
            bold = self.active_styles["bold"]
            italic = self.active_styles["italic"]
            underline = self.active_styles["underline"]
            if style == "underline":
                if underline:
                    ed.tag_add("underline", start, end)
                else:
                    ed.tag_remove("underline", start, end)
            else:
                self.apply_font_state_to_range(ed, start, end, fam, sz, bold, italic)

        self._update_style_buttons()
        ed.focus_set()

    def on_text_key_release(self, event, page_id):
        if self.pdf_mode:
            return

        ed = self.text_page_editors.get(page_id)
        if not ed: return

        nav_or_modifier_keys = {
            "Left", "Right", "Up", "Down", "Home", "End", "Prior", "Next",
            "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R",
            "Caps_Lock", "Tab"
        }
        if event.keysym in nav_or_modifier_keys or event.keysym in ("BackSpace", "Delete"):
            return

        try:
            ed.index("sel.first")
            return
        except tk.TclError:
            pass

        try:
            end_idx = ed.index("insert")
            start_idx = ed.index("insert -1c")
            fam = self.font_family.get()
            try:
                sz = int(self.font_size.get())
            except Exception:
                sz = 12
            self.apply_font_state_to_range(ed, start_idx, end_idx, fam, sz,
                                            self.active_styles["bold"], self.active_styles["italic"])

            if self.active_styles["underline"]:
                ed.tag_add("underline", start_idx, end_idx)
            else:
                ed.tag_remove("underline", start_idx, end_idx)
        except tk.TclError:
            pass

    def toggle_bold(self): self.toggle_style("bold")
    def toggle_italic(self): self.toggle_style("italic")
    def toggle_underline(self): self.toggle_style("underline")

    def set_alignment(self, align):
        ed = self.text_page_editors.get(self.active_page_id)
        if not ed: return
        try:
            start, end = ed.index("sel.first"), ed.index("sel.last")
            for a in ["left", "center", "right"]:
                ed.tag_remove(a, start, end)
            ed.tag_add(align, start, end)
        except tk.TclError:
            self.active_styles["alignment"] = align
            try:
                line_start = ed.index("insert linestart")
                line_end = ed.index("insert lineend")
                for a in ["left", "center", "right"]:
                    ed.tag_remove(a, line_start, line_end)
                ed.tag_add(align, line_start, line_end)
            except Exception:
                pass

    def select_all(self):
        if self.active_pdf_box is not None:
            self.active_pdf_box.select_range(0, tk.END)
            return "break"
        ed = self.text_page_editors.get(self.active_page_id)
        if ed:
            ed.tag_add("sel", "1.0", tk.END)
        return "break"

    def trigger_undo(self):
        ed = self.text_page_editors.get(self.active_page_id)
        if ed:
            try: ed.edit_undo()
            except tk.TclError: pass
        return "break"

    def trigger_redo(self):
        ed = self.text_page_editors.get(self.active_page_id)
        if ed:
            try: ed.edit_redo()
            except tk.TclError: pass
        return "break"

    def new_file(self):
        self.switch_to_text_view()
        self.default_font_size = max(1, int(self.font_size.get()))
        for pid in list(self.page_text_box_widgets.keys()):
            self._clear_text_box_widgets(pid)
        self._clear_all_images()
        self.pages_data = {1: ""}
        self.page_text_boxes = {1: []}
        self.page_text_format = {1: []}
        self.page_text_box_widgets = {1: []}
        self.page_images = {1: []}
        self.page_image_widgets = {1: []}
        self.page_display_names = {}
        self.active_page_id = 1
        self.current_file = None
        self._rtf_tags_pending = None
        self.embedded_images_cache.clear()
        # Stesso fix di open_file: distrugge i vecchi editor prima del reset,
        # altrimenti save_current_page_state() dentro update_page_layout()
        # rimette il vecchio testo nel documento "nuovo".
        self.canvas_desktop.delete("text_page_win")
        for pid, frame in list(self.text_page_frames.items()):
            try:
                frame.destroy()
            except Exception:
                pass
        self.text_page_frames = {}
        self.text_page_editors = {}
        self.update_page_layout()
        self.root.title("Nuovo Documento - UniversalWriter")
        return "break"

    def save_file(self):
        if self.current_file:
            self.export_universal(self.current_file)
        else:
            self.save_as_file()
        return "break"


def _generate_tags_from_text(text, text_chars):
    """
    Genera la lista di tag da applicare al testo Tkinter.
    Trasforma i char indices in Tkinter indices (line.col).
    """
    tags = []
    
    if not text_chars:
        return tags
    
    # Traccia le "run" di formattazione (sequenze di caratteri con lo stesso stile)
    current_run = None
    current_run_start = 0
    
    for char_idx, (char, fmt) in enumerate(text_chars):
        run_key = (fmt["bold"], fmt["italic"], fmt["underline"])
        
        if current_run is None:
            # Inizio prima run
            current_run = run_key
            current_run_start = char_idx
        elif current_run != run_key:
            # Fine della run precedente
            if current_run != (False, False, False):
                # Genera tag per questa run
                bold, italic, underline = current_run
                if bold:
                    tags.append(("bold", current_run_start, char_idx))
                if italic:
                    tags.append(("italic", current_run_start, char_idx))
                if underline:
                    tags.append(("underline", current_run_start, char_idx))
            
            # Inizia nuova run
            current_run = run_key
            current_run_start = char_idx
    
    # Finisci l'ultima run
    if current_run and current_run != (False, False, False):
        bold, italic, underline = current_run
        if bold:
            tags.append(("bold", current_run_start, len(text_chars)))
        if italic:
            tags.append(("italic", current_run_start, len(text_chars)))
        if underline:
            tags.append(("underline", current_run_start, len(text_chars)))
    
    # Converti char indices a Tkinter indices
    tkinter_tags = []
    for tag_name, start_char, end_char in tags:
        start_idx = _char_index_to_tkinter(text, start_char)
        end_idx = _char_index_to_tkinter(text, end_char)
        tkinter_tags.append((tag_name, start_idx, end_idx))
    
    return tkinter_tags


def _char_index_to_tkinter(text, char_idx):
    """Converte indice char in indice Tkinter (line.col)."""
    line = 1
    col = 0
    for i in range(min(char_idx, len(text))):
        if text[i] == "\n":
            line += 1
            col = 0
        else:
            col += 1
    return f"{line}.{col}"


if __name__ == "__main__":
    root = tk.Tk()
    app = AdvancedTextEditor(root)
    root.mainloop()
