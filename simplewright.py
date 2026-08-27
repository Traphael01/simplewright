import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font
from tkinter.scrolledtext import ScrolledText
import os
import platform
import time
import subprocess
import io

# Gestione formati estesi con PyMuPDF e python-docx
import pymupdf
from docx import Document

# Prova a importare docx2pdf per la conversione nativa dei file Word
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

class AdvancedTextEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("Simplewright - Multi-Page Editor")
        self.root.geometry("1400x850")

        self.current_file = None
        self.is_dark = False

        # Stati di scrittura attiva (quando non c'è selezione)
        self.active_styles = {
            "bold": False,
            "italic": False,
            "underline": False,
            "alignment": "left"
        }

        # Stato dei colori di testo e evidenziatore
        self.current_color = "#000000"
        self.recent_colors = ["#ffffff", "#ffffff", "#ffffff", "#ffffff"]
        self.selected_slot_idx = 0
        self.current_bg_color = "#ffffff"

        # Gestione Popup
        self.current_popup = None
        self.last_popup_closed_time = 0
        self.active_popup_type = None

        # Dimensioni BASE dei fogli
        self.base_page_sizes = {
            "A4": (595, 842),
            "A3": (842, 1191),
            "A2": (1191, 1684)
        }
        self.current_format = "A4"
        self.current_orientation = "Verticale"
        self.zoom_factor = 1.0

        # Throttle/Debounce per lo Zoom fluido senza lag
        self._zoom_job = None

        # --- STRUTTURA MULTI-PAGINA ---
        self.pages_data = {1: ""}
        self.active_page_id = 1

        # Cache interna immagini
        self.embedded_images_cache = {}

        # --- STATO MODALITA' PDF/DOC (overlay tipo Adobe Acrobat) ---
        self.pdf_mode = False
        self.pdf_pages_info = {}
        self.pdf_render_dpi = 150
        self.pdf_add_mode = False
        self.pdf_photo_refs = {}
        self.pdf_entry_widgets = {}
        self.active_pdf_box = None

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

        self.setup_tags()
        self.setup_global_shortcuts()
        self.apply_theme()

        self.editor.bind("<KeyRelease>", self.on_text_key_release)

        self.update_pages_sidebar()
        self.update_page_layout()

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

        self.all_fonts = sorted(list(font.families()))
        self.font_family = tk.StringVar(value="Arial")
        self.font_cb = ttk.Combobox(self.toolbar, textvariable=self.font_family, values=self.all_fonts, width=15, state="readonly")
        self.font_cb.pack(side=tk.LEFT, padx=5)
        self.font_cb.bind("<<ComboboxSelected>>", self.change_font)

        self.font_size = tk.IntVar(value=12)
        self.size_cb = ttk.Combobox(self.toolbar, textvariable=self.font_size, values=list(range(8, 73, 2)), width=4, state="readonly")
        self.size_cb.pack(side=tk.LEFT, padx=5)
        self.size_cb.bind("<<ComboboxSelected>>", self.change_font)

        ttk.Separator(self.toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        self.btn_bold = ttk.Button(self.toolbar, text="B", width=3, command=self.toggle_bold)
        self.btn_bold.pack(side=tk.LEFT, padx=1)
        self.btn_italic = ttk.Button(self.toolbar, text="I", width=3, command=self.toggle_italic)
        self.btn_italic.pack(side=tk.LEFT, padx=1)
        self.btn_underline = ttk.Button(self.toolbar, text="U", width=3, command=self.toggle_underline)
        self.btn_underline.pack(side=tk.LEFT, padx=1)

        ttk.Separator(self.toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        self.btn_left = ttk.Button(self.toolbar, text="⌸ L", width=4, command=lambda: self.set_alignment("left"))
        self.btn_left.pack(side=tk.LEFT, padx=1)
        self.btn_center = ttk.Button(self.toolbar, text="⌸ C", width=4, command=lambda: self.set_alignment("center"))
        self.btn_center.pack(side=tk.LEFT, padx=1)
        self.btn_right = ttk.Button(self.toolbar, text="⌸ R", width=4, command=lambda: self.set_alignment("right"))
        self.btn_right.pack(side=tk.LEFT, padx=1)

        ttk.Separator(self.toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        self.btn_add_text = ttk.Button(self.toolbar, text="+ Casella di Testo", command=self.toggle_add_text_mode)
        self.btn_add_text.pack(side=tk.LEFT, padx=1)

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

        self.sidebar_canvas = tk.Canvas(self.right_sidebar, width=150, bd=0, highlightthickness=0)
        self.sidebar_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.sidebar_scroll = ttk.Scrollbar(self.right_sidebar, orient=tk.VERTICAL, command=self.sidebar_canvas.yview)
        self.sidebar_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.sidebar_canvas.configure(yscrollcommand=self.sidebar_scroll.set)

        self.pages_list_frame = ttk.Frame(self.sidebar_canvas)
        self.sidebar_canvas.create_window((0, 0), window=self.pages_list_frame, anchor="nw")
        self.pages_list_frame.bind("<Configure>", lambda e: self.sidebar_canvas.configure(scrollregion=self.sidebar_canvas.bbox("all")))

        self.canvas_desktop = tk.Canvas(self.lower_workspace, bg="#e0e0e0", bd=0, highlightthickness=0)
        self.canvas_desktop.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar_y = ttk.Scrollbar(self.lower_workspace, orient=tk.VERTICAL, command=self.canvas_desktop.yview)
        self.scrollbar_y.pack(side=tk.LEFT, fill=tk.Y)
        self.canvas_desktop.configure(yscrollcommand=self.scrollbar_y.set)

        self.page_frame = tk.Frame(self.canvas_desktop, bg="#ffffff", bd=1, relief="solid")
        self.page_window = self.canvas_desktop.create_window((50, 30), window=self.page_frame, anchor="nw")

        self.editor = ScrolledText(self.page_frame, wrap=tk.WORD, font=("Arial", 12), undo=True, maxundo=-1, bd=0, highlightthickness=0)
        self.editor.pack(fill=tk.BOTH, expand=True, side=tk.TOP)
        self.editor.focus_set()

        self.resize_grip = tk.Label(self.page_frame, text="◢", bg="#ffffff", fg="#888888", cursor="sizing")
        self.resize_grip.place(relx=1.0, rely=1.0, x=-2, y=-2, anchor="se")
        self.resize_grip.bind("<B1-Motion>", self.on_sheet_resize_motion)

        self.pdf_canvas = tk.Canvas(self.page_frame, bd=0, highlightthickness=0, bg="#ffffff")
        self.pdf_canvas.bind("<Button-1>", self.on_pdf_canvas_click)

        self.canvas_desktop.bind("<Configure>", lambda e: self.center_page_on_desktop())

        self.canvas_desktop.bind("<Button-4>", self.handle_mouse_wheel)
        self.canvas_desktop.bind("<Button-5>", self.handle_mouse_wheel)
        self.canvas_desktop.bind("<MouseWheel>", self.handle_mouse_wheel)

        self.editor.bind("<Button-4>", self.handle_mouse_wheel)
        self.editor.bind("<Button-5>", self.handle_mouse_wheel)
        self.editor.bind("<MouseWheel>", self.handle_mouse_wheel)

    def on_sheet_resize_motion(self, event):
        current_w = self.page_frame.winfo_width()
        current_h = self.page_frame.winfo_height()
        new_w = max(300, current_w + event.x)
        new_h = max(300, current_h + event.y)
        self.page_frame.config(width=new_w, height=new_h)
        self.page_frame.pack_propagate(False)
        self.canvas_desktop.config(scrollregion=(0, 0, new_w + 100, new_h + 100))
        self.center_page_on_desktop()

    def handle_mouse_wheel(self, event):
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

    def save_current_page_state(self):
        if not self.pdf_mode:
            self.pages_data[self.active_page_id] = self.editor.get("1.0", tk.END + "-1c")

    def switch_to_page(self, page_id):
        if self.pdf_mode:
            self.render_pdf_page(page_id)
            self.update_pages_sidebar()
            return
        self.save_current_page_state()
        self.active_page_id = page_id

        self.editor.delete("1.0", tk.END)
        content_to_insert = self.pages_data.get(page_id, "")
        self.editor.insert("1.0", content_to_insert)

        self.update_pages_sidebar()

    def add_new_page(self):
        if self.pdf_mode:
            messagebox.showinfo("Modalità Documento Visivo", "In questa modalità le pagine derivano dal documento caricato.")
            return
        self.save_current_page_state()
        new_id = max(self.pages_data.keys()) + 1 if self.pages_data else 1
        self.pages_data[new_id] = ""
        self.switch_to_page(new_id)

    def delete_page(self, page_id):
        if self.pdf_mode:
            return
        if len(self.pages_data) <= 1:
            messagebox.showwarning("Attenzione", "Impossibile eliminare l'unica pagina presente.")
            return
        del self.pages_data[page_id]
        sorted_content = [self.pages_data[k] for k in sorted(self.pages_data.keys())]
        self.pages_data = {i+1: content for i, content in enumerate(sorted_content)}

        if self.active_page_id not in self.pages_data:
            self.active_page_id = max(self.pages_data.keys())
        self.switch_to_page(self.active_page_id)

    def update_pages_sidebar(self):
        for widget in self.pages_list_frame.winfo_children():
            widget.destroy()
        keys = sorted(self.pdf_pages_info.keys()) if self.pdf_mode else sorted(self.pages_data.keys())
        for p_id in keys:
            p_frame = ttk.Frame(self.pages_list_frame, padding=2)
            p_frame.pack(fill=tk.X, pady=2)

            btn = ttk.Button(p_frame, text=f"Foglio {p_id}", width=12, command=lambda idx=p_id: self.switch_to_page(idx))
            btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

            if not self.pdf_mode:
                btn_del = tk.Button(p_frame, text="X", fg="red", bg="#fce8e6", bd=0, font=("Arial", 8, "bold"), command=lambda idx=p_id: self.delete_page(idx))
                btn_del.pack(side=tk.RIGHT, padx=2)

    def open_file(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("Tutti i formati supportati", "*.txt *.docx *.doc *.odt *.rtf *.md *.epub *.pdf *.ods"),
            ("Documenti PDF", "*.pdf"),
            ("Documenti Word", "*.docx *.doc"),
            ("Documenti OpenDocument", "*.odt *.ods"),
            ("E-book", "*.epub"),
            ("File di Testo / Markdown", "*.txt *.md *.rtf"),
            ("Tutti i file", "*.*")
        ])
        if not file_path:
            return "break"

        ext = os.path.splitext(file_path)[1].lower()

        self.pages_data.clear()
        self.embedded_images_cache.clear()
        self.pdf_pages_info.clear()
        self.pdf_photo_refs.clear()
        self.pdf_entry_widgets.clear()
        self.active_pdf_box = None

        try:
            # Riconoscimento intelligente e conversione automatica per i file complessi (Word, ODT)
            if ext in [".docx", ".doc", ".odt"]:
                temp_pdf_path = os.path.join(os.path.expanduser("~"), "_temp_doc_converted.pdf")
                converted = False

                if ext == ".docx" and HAS_DOCX2PDF:
                    try:
                        convert_docx_to_pdf(file_path, temp_pdf_path)
                        converted = True
                    except Exception:
                        pass

                if not converted:
                    try:
                        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(temp_pdf_path), file_path], check=True)
                        base_name = os.path.splitext(os.path.basename(file_path))[0]
                        generated_pdf = os.path.join(os.path.dirname(temp_pdf_path), base_name + ".pdf")
                        if os.path.exists(generated_pdf):
                            if generated_pdf != temp_pdf_path:
                                if os.path.exists(temp_pdf_path):
                                    os.remove(temp_pdf_path)
                                os.rename(generated_pdf, temp_pdf_path)
                            converted = True
                    except Exception:
                        pass

                if converted and os.path.exists(temp_pdf_path):
                    self.load_pdf(temp_pdf_path)
                    self.switch_to_pdf_view()
                    self.current_file = file_path
                    self.update_pages_sidebar()
                    self.render_pdf_page(1)
                    self.root.title(f"{file_path} - UniversalWriter [Layout Originale]")
                else:
                    # Fallback strutturale se non è possibile convertire via suite esterne
                    self.switch_to_text_view()
                    doc = Document(file_path)
                    page_idx = 1
                    current_page_text = ""
                    for p in doc.paragraphs:
                        current_page_text += p.text + "\n"
                        if len(current_page_text.split('\n')) > 35:
                            self.pages_data[page_idx] = current_page_text.strip()
                            page_idx += 1
                            current_page_text = ""
                    self.pages_data[page_idx] = current_page_text.strip()
                    self.current_file = file_path
                    self.switch_to_page(1)
                    self.update_page_layout()
                    self.root.title(f"{file_path} - UniversalWriter")

            elif ext in [".pdf", ".epub", ".ods"]:
                self.load_pdf(file_path)
                self.switch_to_pdf_view()
                self.current_file = file_path
                self.update_pages_sidebar()
                self.render_pdf_page(1)
                self.root.title(f"{file_path} - UniversalWriter [Layout Originale]")
            else:
                self.switch_to_text_view()
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    self.pages_data[1] = f.read()
                self.current_file = file_path
                self.switch_to_page(1)
                self.update_page_layout()
                self.root.title(f"{file_path} - UniversalWriter")
        except Exception as e:
            messagebox.showerror("Errore di Lettura", f"Errore nell'apertura del file:\n{str(e)}")
            self.new_file()
        return "break"

    def switch_to_pdf_view(self):
        self.pdf_mode = True
        self.pdf_add_mode = False
        self.editor.pack_forget()
        self.resize_grip.pack_forget()
        self.pdf_canvas.pack(fill=tk.BOTH, expand=True)
        self.btn_add_page.config(state="disabled")

    def switch_to_text_view(self):
        self.pdf_mode = False
        self.pdf_add_mode = False
        self.pdf_canvas.pack_forget()
        self.editor.pack(fill=tk.BOTH, expand=True, side=tk.TOP)
        self.resize_grip.place(relx=1.0, rely=1.0, x=-2, y=-2, anchor="se")
        self.btn_add_page.config(state="normal")

    def load_pdf(self, file_path):
        doc = pymupdf.open(file_path)
        for idx, page in enumerate(doc):
            page_id = idx + 1
            text_dict = page.get_text("dict")
            spans = []

            for block in text_dict.get("blocks", []):
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
                        })
                        rect = pymupdf.Rect(b)
                        page.add_redact_annot(rect, fill=(1, 1, 1))

            page.apply_redactions()
            pix = page.get_pixmap(dpi=self.pdf_render_dpi)
            img_bytes = pix.tobytes("png")

            self.pdf_pages_info[page_id] = {
                "spans": spans,
                "img_bytes": img_bytes,
                "width_pt": page.rect.width,
                "height_pt": page.rect.height,
                "overlays": [],
            }
        doc.close()
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

    def render_pdf_page(self, page_id):
        info = self.pdf_pages_info.get(page_id)
        if not info:
            return
        self.active_page_id = page_id
        self.pdf_canvas.delete("all")
        self.pdf_entry_widgets[page_id] = []
        self.active_pdf_box = None

        scale = self.pdf_render_dpi / 72.0
        img = Image.open(io.BytesIO(info["img_bytes"]))
        photo = ImageTk.PhotoImage(img)
        self.pdf_photo_refs[page_id] = photo

        canvas_w, canvas_h = img.width, img.height
        self.pdf_canvas.config(width=canvas_w, height=canvas_h, scrollregion=(0, 0, canvas_w, canvas_h))
        self.pdf_canvas.create_image(0, 0, anchor="nw", image=photo)

        self.page_frame.config(width=canvas_w, height=canvas_h)
        self.page_frame.pack_propagate(False)

        for span in info["spans"]:
            x0, y0, x1, y1 = span["bbox"]
            cx, cy = x0 * scale, y0 * scale

            pt_size = max(8, int(round(span["size"])))
            cw = max(40, span["width"] * scale * self.zoom_factor)
            ch = max(pt_size + 12, span["height"] * scale * self.zoom_factor)

            family, weight, slant = self._font_details_from_span(span)
            color_hex = self._int_color_to_hex(span["color"])

            span.setdefault("custom_font", family)
            span.setdefault("custom_size", pt_size)
            span.setdefault("custom_weight", weight)
            span.setdefault("custom_slant", slant)

            container, entry, grip = self._create_resizable_box(
                self.pdf_canvas, cx * self.zoom_factor, cy * self.zoom_factor, cw, ch,
                span["text"], span["custom_font"], span["custom_size"], span["custom_weight"], span["custom_slant"], color_hex, span
            )
            win_id = self.pdf_canvas.create_window(cx * self.zoom_factor, cy * self.zoom_factor, anchor="nw", window=container, width=cw, height=ch)

            self.pdf_entry_widgets[page_id].append((win_id, container, entry, span))

        for ov in info["overlays"]:
            self._create_overlay_entry(page_id, ov, scale)

        self.center_page_on_desktop()

    def _create_resizable_box(self, parent_canvas, x, y, w, h, text, family, pt_size, weight, slant, color_hex, data_obj):
        container = tk.Frame(parent_canvas, bg="#ffffff", bd=1, relief="solid", highlightthickness=1, highlightbackground="#b0c4de", highlightcolor="#3399ff", width=w, height=h)
        container.pack_propagate(False)

        box_font = font.Font(family=family, size=pt_size, weight=weight, slant=slant)

        entry = tk.Entry(
            container, bd=0, highlightthickness=0,
            bg="#ffffff", fg=color_hex,
            font=box_font,
        )
        entry.insert(0, text)
        entry.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=2, pady=2)

        entry.box_font = box_font
        entry.data_obj = data_obj

        entry.bind("<KeyRelease>", lambda e, d=data_obj: d.__setitem__("text", e.widget.get()))
        entry.bind("<FocusIn>", lambda e, en=entry: self.set_active_pdf_box(en))
        entry.bind("<Button-1>", lambda e, en=entry: self.set_active_pdf_box(en))

        grip = tk.Label(container, text="◢", bg="#ffffff", fg="#666666", cursor="sizing", font=("Arial", 7))
        grip.place(relx=1.0, rely=1.0, x=0, y=0, anchor="se")

        def start_resize(event):
            container._drag_start_x = event.x_root
            container._drag_start_y = event.y_root
            container._start_w = container.winfo_width()
            container._start_h = container.winfo_height()

        def do_resize(event):
            dx = event.x_root - container._drag_start_x
            dy = event.y_root - container._drag_start_y
            new_w = max(40, container._start_w + dx)
            new_h = max(20, container._start_h + dy)

            container.config(width=new_w, height=new_h)
            for item_id, c_win, _, _ in self.pdf_entry_widgets.get(self.active_page_id, []):
                if c_win == container:
                    parent_canvas.itemconfigure(item_id, width=new_w, height=new_h)
                    break

            scale = (self.pdf_render_dpi / 72.0) * self.zoom_factor
            data_obj["width"] = new_w / scale
            data_obj["height"] = new_h / scale

        grip.bind("<Button-1>", start_resize)
        grip.bind("<B1-Motion>", do_resize)

        return container, entry, grip

    def set_active_pdf_box(self, entry_widget):
        self.active_pdf_box = entry_widget
        try:
            f_fam = entry_widget.box_font.actual()["family"]
            f_size = entry_widget.box_font.actual()["size"]
            if f_size < 0: f_size = abs(f_size)
            self.font_family.set(f_fam)
            self.font_size.set(f_size)

            is_bold = (entry_widget.box_font.actual()["weight"] == "bold")
            is_italic = (entry_widget.box_font.actual()["slant"] == "italic")
            self.active_styles["bold"] = is_bold
            self.active_styles["italic"] = is_italic
            self.btn_bold.config(relief="sunken" if is_bold else "raised")
            self.btn_italic.config(relief="sunken" if is_italic else "raised")
        except Exception:
            pass

    def _create_overlay_entry(self, page_id, ov, scale):
        cx, cy = ov["x"] * scale * self.zoom_factor, ov["y"] * scale * self.zoom_factor
        pt_size = int(ov.get("size", 12.0))

        ov.setdefault("custom_font", "Arial")
        ov.setdefault("custom_size", pt_size)
        ov.setdefault("custom_weight", "normal")
        ov.setdefault("custom_slant", "roman")

        container, entry, grip = self._create_resizable_box(
            self.pdf_canvas, cx, cy, 180, pt_size + 14,
            ov["text"], ov["custom_font"], ov["custom_size"], ov["custom_weight"], ov["custom_slant"], "#000000", ov
        )
        win_id = self.pdf_canvas.create_window(cx, cy, anchor="nw", window=container, width=180, height=pt_size + 14)
        self.pdf_entry_widgets[page_id].append((win_id, container, entry, ov))

    def toggle_add_text_mode(self):
        if not self.pdf_mode:
            messagebox.showinfo("Modalità Modifica", "Apri prima un file PDF o Word per poter aggiungere caselle di testo.")
            return
        self.pdf_add_mode = not self.pdf_add_mode
        self.pdf_canvas.config(cursor="crosshair" if self.pdf_add_mode else "")

    def on_pdf_canvas_click(self, event):
        if not self.pdf_mode or not self.pdf_add_mode:
            return
        scale = (self.pdf_render_dpi / 72.0) * self.zoom_factor
        x_pt = self.pdf_canvas.canvasx(event.x) / scale
        y_pt = self.pdf_canvas.canvasy(event.y) / scale

        info = self.pdf_pages_info[self.active_page_id]
        ov = {
            "x": x_pt, "y": y_pt, "text": "", "size": float(self.font_size.get()),
            "color": 0x000000, "width": 150, "height": 30,
            "custom_font": self.font_family.get(),
            "custom_size": int(self.font_size.get()),
            "custom_weight": "bold" if self.active_styles["bold"] else "normal",
            "custom_slant": "italic" if self.active_styles["italic"] else "roman"
        }
        info["overlays"].append(ov)
        self._create_overlay_entry(self.active_page_id, ov, scale)
        self.toggle_add_text_mode()

    def save_as_file(self):
        file_path = filedialog.asksaveasfilename(
            filetypes=[
                ("Tutti i formati supportati", "*.txt *.docx *.doc *.odt *.rtf *.md *.epub *.pdf *.ods"),
                ("Documenti Word", "*.docx"),
                ("Documenti PDF", "*.pdf"),
                ("OpenDocument Text", "*.odt"),
                ("OpenDocument Spreadsheet", "*.ods"),
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
            if ext == ".pdf":
                self._export_to_pdf_universal(file_path)
            elif ext in [".docx", ".doc", ".odt"]:
                self._export_to_docx_universal(file_path)
            elif ext in [".txt", ".md", ".rtf", ".epub", ".ods"]:
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
                page_text = self.pages_data[p_id]
                for line in page_text.split("\n"):
                    if line.strip() == "":
                        story.append(Spacer(1, 12))
                    else:
                        clean_line = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        story.append(Paragraph(clean_line, custom_style))
                if p_id != max(self.pages_data.keys()):
                    story.append(PageBreak())
            doc.build(story)

    def _export_to_docx_universal(self, file_path):
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
            for p_id in sorted(self.pages_data.keys()):
                doc.add_paragraph(self.pages_data[p_id])
                if p_id != max(self.pages_data.keys()):
                    doc.add_page_break()
        doc.save(file_path)

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
        m_left = int(self.scale_m_left.get())
        m_right = int(self.scale_m_right.get())
        m_top = int(self.scale_m_top.get())
        m_bottom = int(self.scale_m_bottom.get())
        self.editor.pack_configure(padx=(m_left, m_right), pady=(m_top, m_bottom))

    def update_page_layout(self):
        base_w, base_h = self.base_page_sizes[self.current_format]
        w = int(base_w * self.zoom_factor)
        h = int(base_h * self.zoom_factor)
        if self.current_orientation == "Orizzontale":
            w, h = h, w

        self.page_frame.config(width=w, height=h)
        self.page_frame.pack_propagate(False)

        current_f_size = int(self.font_size.get() * self.zoom_factor)
        self.editor.config(font=(self.font_family.get(), max(8, current_f_size)))

        self.canvas_desktop.config(scrollregion=(0, 0, w + 100, h + 100))
        self.center_page_on_desktop()
        self.on_margin_change()
        if self.pdf_mode:
            self.render_pdf_page(self.active_page_id)

    def center_page_on_desktop(self):
        canvas_width = self.canvas_desktop.winfo_width()
        page_width = self.page_frame.winfo_width()
        new_x = max(20, (canvas_width - page_width) // 2)
        self.canvas_desktop.coords(self.page_window, new_x, 30)

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
            self.editor.config(bg="#ffffff", fg="#000000", insertbackground="black")
            self.page_frame.config(bg="#ffffff")
        else:
            self.canvas_desktop.config(bg="#e0e0e0")
            self.editor.config(bg="#ffffff", fg="#000000", insertbackground="black")
            self.page_frame.config(bg="#ffffff")

    def toggle_dark_mode(self):
        self.is_dark = not self.is_dark
        self.apply_theme()
        self.save_config()

    def setup_tags(self):
        self.editor.tag_configure("bold", font=("Arial", 12, "bold"))
        self.editor.tag_configure("italic", font=("Arial", 12, "italic"))
        self.editor.tag_configure("underline", underline=True)
        self.editor.tag_configure("left", justify=tk.LEFT)
        self.editor.tag_configure("center", justify=tk.CENTER)
        self.editor.tag_configure("right", justify=tk.RIGHT)

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
        if self.pdf_mode and self.active_pdf_box:
            try:
                self.active_pdf_box.config(fg=color)
                self.active_pdf_box.data_obj["color"] = int(color.replace("#", ""), 16)
            except Exception:
                pass
            return

        try:
            start, end = self.editor.index("sel.first"), self.editor.index("sel.last")
            tag_name = f"color_{color.replace('#', '')}"
            self.editor.tag_configure(tag_name, foreground=color)
            self.editor.tag_add(tag_name, start, end)
        except tk.TclError: self.editor.config(fg=color)

    def set_text_background_color(self, color):
        if not color: return
        self.current_bg_color = color
        try: start, end = self.editor.index("sel.first"), self.editor.index("sel.last")
        except tk.TclError: return
        tag_name = f"bg_{color.replace('#', '')}"
        self.editor.tag_configure(tag_name, background=color)
        self.editor.tag_add(tag_name, start, end)

    def change_font(self, event=None):
        if self.pdf_mode and self.active_pdf_box:
            try:
                fam = self.font_family.get()
                sz = int(self.font_size.get())
                self.active_pdf_box.box_font.config(family=fam, size=sz)
                self.active_pdf_box.data_obj["custom_font"] = fam
                self.active_pdf_box.data_obj["custom_size"] = sz
            except Exception:
                pass
            return
        self.update_page_layout()

    def toggle_style(self, tag_name):
        if self.pdf_mode and self.active_pdf_box:
            try:
                current_w = self.active_pdf_box.box_font.actual()["weight"]
                current_s = self.active_pdf_box.box_font.actual()["slant"]

                if tag_name == "bold":
                    new_w = "normal" if current_w == "bold" else "bold"
                    self.active_pdf_box.box_font.config(weight=new_w)
                    self.active_pdf_box.data_obj["custom_weight"] = new_w
                    self.active_styles["bold"] = (new_w == "bold")
                    self.btn_bold.config(relief="sunken" if new_w == "bold" else "raised")
                elif tag_name == "italic":
                    new_s = "roman" if current_s == "italic" else "italic"
                    self.active_pdf_box.box_font.config(slant=new_s)
                    self.active_pdf_box.data_obj["custom_slant"] = new_s
                    self.active_styles["italic"] = (new_s == "italic")
                    self.btn_italic.config(relief="sunken" if new_s == "italic" else "raised")
            except Exception as e:
                print("Error toggling PDF box style:", e)
            return

        try:
            start, end = self.editor.index("sel.first"), self.editor.index("sel.last")
            if tag_name in self.editor.tag_names("sel.first"):
                self.editor.tag_remove(tag_name, start, end)
            else:
                self.editor.tag_add(tag_name, start, end)
        except tk.TclError:
            self.active_styles[tag_name] = not self.active_styles[tag_name]
            btn_map = {"bold": self.btn_bold, "italic": self.btn_italic, "underline": self.btn_underline}
            if tag_name in btn_map:
                relief = "sunken" if self.active_styles[tag_name] else "raised"
                btn_map[tag_name].config(relief=relief)

    def on_text_key_release(self, event):
        if self.pdf_mode:
            return
        try:
            self.editor.index("sel.first")
        except tk.TclError:
            for style in ["bold", "italic", "underline"]:
                if self.active_styles[style]:
                    end_idx = self.editor.index("insert")
                    start_idx = self.editor.index("insert - 1c")
                    self.editor.tag_add(style, start_idx, end_idx)

    def toggle_bold(self): self.toggle_style("bold")
    def toggle_italic(self): self.toggle_style("italic")
    def toggle_underline(self): self.toggle_style("underline")

    def set_alignment(self, align):
        try:
            start, end = self.editor.index("sel.first"), self.editor.index("sel.last")
            for a in ["left", "center", "right"]:
                self.editor.tag_remove(a, start, end)
            self.editor.tag_add(align, start, end)
        except tk.TclError:
            self.active_styles["alignment"] = align
            try:
                line_start = self.editor.index("insert linestart")
                line_end = self.editor.index("insert lineend")
                for a in ["left", "center", "right"]:
                    self.editor.tag_remove(a, line_start, line_end)
                self.editor.tag_add(align, line_start, line_end)
            except Exception:
                pass

    def select_all(self):
        if self.pdf_mode and self.active_pdf_box:
            self.active_pdf_box.select_range(0, tk.END)
            return "break"
        self.editor.tag_add("sel", "1.0", tk.END)
        return "break"

    def trigger_undo(self):
        try: self.editor.edit_undo()
        except tk.TclError: pass
        return "break"

    def trigger_redo(self):
        try: self.editor.edit_redo()
        except tk.TclError: pass
        return "break"

    def new_file(self):
        self.switch_to_text_view()
        self.editor.delete("1.0", tk.END)
        self.pages_data = {1: ""}
        self.active_page_id = 1
        self.current_file = None
        self.embedded_images_cache.clear()
        self.update_pages_sidebar()
        self.root.title("Nuovo Documento - UniversalWriter")
        return "break"

    def save_file(self):
        if self.current_file:
            self.export_universal(self.current_file)
        else:
            self.save_as_file()
        return "break"

if __name__ == "__main__":
    root = tk.Tk()
    app = AdvancedTextEditor(root)
    root.mainloop()