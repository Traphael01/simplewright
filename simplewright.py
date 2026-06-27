import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font
from tkinter.scrolledtext import ScrolledText
from tkinter.colorchooser import askcolor
import os
import platform
import time
import subprocess

# Librerie per la gestione dei formati estesi
from docx import Document
from pypdf import PdfReader

# Libreria per esportare in PDF
from reportlab.lib.pagesizes import A4, A3, A2
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Librerie per OCR e immagini
import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageDraw, ImageFont, ImageTk

class AdvancedTextEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("UniversalWriter - Multi-Page Editor")
        self.root.geometry("1400x850")
        
        self.current_file = None
        self.is_dark = False  
        
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
        
        # --- STRUTTURA MULTI-PAGINA ---
        self.pages_data = {1: ""}
        self.active_page_id = 1
        
        # Cache interna per evitare il Garbage Collector sulle immagini Tkinter
        self.embedded_images_cache = {}
        
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
        
        self.update_pages_sidebar()
        self.update_page_layout()    

    def get_config_path(self):
        system_os = platform.system().lower()
        home_dir = os.path.expanduser("~")
        if "windows" in system_os:
            appdata_local_low = os.path.join(home_dir, "AppData", "LocalLow", "simplewright")
            os.makedirs(appdata_local_low, exist_ok=True)
            return os.path.join(appdata_local_low, "config.txt")
        else:
            linux_config_dir = os.path.join(home_dir, ".simplewright")
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
        except Exception as e:
            print(f"Impossibile salvare la configurazione: {e}")

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
        except: pil_font = ImageFont.load_default()
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
        file_menu.add_command(label="Salva con nome...", command=self.save_as_file)
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
        self.root.bind("<Control-p>", lambda e: self.print_document())
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

        # --- SEZIONE DESTRA ---
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
        
        # --- MARGINI UNITI VICINI ---
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

        # --- BARRA LATERALE DESTRA ---
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

        # La Scrivania Centrale
        self.canvas_desktop = tk.Canvas(self.lower_workspace, bg="#e0e0e0", bd=0, highlightthickness=0)
        self.canvas_desktop.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.scrollbar_y = ttk.Scrollbar(self.lower_workspace, orient=tk.VERTICAL, command=self.canvas_desktop.yview)
        self.scrollbar_y.pack(side=tk.LEFT, fill=tk.Y)
        self.canvas_desktop.configure(yscrollcommand=self.scrollbar_y.set)

        self.page_frame = tk.Frame(self.canvas_desktop, bg="#ffffff", bd=1, relief="solid")
        self.page_window = self.canvas_desktop.create_window((50, 30), window=self.page_frame, anchor="nw")
        
        self.editor = ScrolledText(self.page_frame, wrap=tk.WORD, font=("Arial", 12), undo=True, maxundo=-1, bd=0, highlightthickness=0)
        self.editor.pack(fill=tk.BOTH, expand=True)
        self.editor.focus_set()
        
        self.canvas_desktop.bind("<Configure>", lambda e: self.center_page_on_desktop())

    def save_current_page_state(self):
        self.pages_data[self.active_page_id] = self.editor.get("1.0", tk.END + "-1c")

    def switch_to_page(self, page_id):
        self.save_current_page_state()
        self.active_page_id = page_id
        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", self.pages_data[page_id])
        
        if page_id in self.embedded_images_cache:
            for tk_img in self.embedded_images_cache[page_id]:
                self.editor.image_create(tk.END, image=tk_img)
                self.editor.insert(tk.END, "\n")
                
        self.update_pages_sidebar()

    def add_new_page(self):
        self.save_current_page_state()
        new_id = max(self.pages_data.keys()) + 1 if self.pages_data else 1
        self.pages_data[new_id] = ""
        self.switch_to_page(new_id)

    def delete_page(self, page_id):
        if len(self.pages_data) <= 1:
            messagebox.showwarning("Attenzione", "Impossibile eliminare l'unica pagina presente.")
            return
        del self.pages_data[page_id]
        if page_id in self.embedded_images_cache:
            del self.embedded_images_cache[page_id]
            
        sorted_content = [self.pages_data[k] for k in sorted(self.pages_data.keys())]
        self.pages_data = {i+1: content for i, content in enumerate(sorted_content)}
        
        if self.active_page_id not in self.pages_data:
            self.active_page_id = max(self.pages_data.keys())
        self.switch_to_page(self.active_page_id)

    def update_pages_sidebar(self):
        for widget in self.pages_list_frame.winfo_children():
            widget.destroy()
        for p_id in sorted(self.pages_data.keys()):
            p_frame = ttk.Frame(self.pages_list_frame, padding=2)
            p_frame.pack(fill=tk.X, pady=2)
            
            btn = ttk.Button(p_frame, text=f"Foglio {p_id}", width=12, command=lambda idx=p_id: self.switch_to_page(idx))
            btn.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            btn_del = tk.Button(p_frame, text="X", fg="red", bg="#fce8e6", bd=0, font=("Arial", 8, "bold"), command=lambda idx=p_id: self.delete_page(idx))
            btn_del.pack(side=tk.RIGHT, padx=2)

    def open_file(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("Tutti i formati supportati", "*.txt *.docx *.pdf *.odt"),
            ("Documenti PDF", "*.pdf"), ("Documenti Word", "*.docx"),
            ("File di Testo", "*.txt"), ("Tutti i file", "*.*")
        ])
        if not file_path: return "break"
        
        ext = os.path.splitext(file_path)[1].lower()
        self.pages_data.clear()
        self.embedded_images_cache.clear()
        
        try:
            if ext == ".pdf":
                reader = PdfReader(file_path)
                num_pages = len(reader.pages)
                
                if num_pages == 0:
                    self.pages_data[1] = ""
                else:
                    images = []
                    try:
                        images = convert_from_path(file_path, dpi=100)
                    except:
                        pass

                    for idx in range(num_pages):
                        page_id = idx + 1
                        page = reader.pages[idx]
                        text = page.extract_text() or ""
                        
                        if not text.strip() and images and idx < len(images):
                            try:
                                text = pytesseract.image_to_string(images[idx], lang="ita")
                            except:
                                text = "[Pagina Grafica / Scansione non editabile]"
                        
                        self.pages_data[page_id] = text
                        
                        if images and idx < len(images):
                            pil_img = images[idx]
                            pil_img.thumbnail((500, 700), Image.Resampling.LANCZOS)
                            tk_img = ImageTk.PhotoImage(pil_img)
                            self.embedded_images_cache[page_id] = [tk_img]
            
            elif ext == ".docx":
                doc = Document(file_path)
                full_text = "\n".join([p.text for p in doc.paragraphs])
                self.pages_data[1] = full_text
                
            else:
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

    def print_document(self):
        self.save_current_page_state()
        temp_pdf = os.path.join(os.path.expanduser("~"), "UniversalWriter_PrintJob.pdf")
        try:
            pdf_sizes = {"A4": A4, "A3": A3, "A2": A2}
            chosen_size = pdf_sizes.get(self.current_format, A4)
            if self.current_orientation == "Orizzontale":
                from reportlab.lib.pagesizes import landscape
                chosen_size = landscape(chosen_size)
                
            doc = SimpleDocTemplate(temp_pdf, pagesize=chosen_size)
            styles = getSampleStyleSheet()
            custom_style = ParagraphStyle('Print_Style', parent=styles['Normal'], fontName='Helvetica', fontSize=self.font_size.get(), leading=self.font_size.get() + 4)
            
            story = []
            for p_id in sorted(self.pages_data.keys()):
                page_text = self.pages_data[p_id]
                for line in page_text.split("\n"):
                    if line.strip() == "": story.append(Spacer(1, 12))
                    else:
                        clean_line = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        story.append(Paragraph(clean_line, custom_style))
                if p_id != max(self.pages_data.keys()):
                    story.append(PageBreak())
            doc.build(story)
            
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
        if self.current_orientation == "Orizzontale": w, h = h, w
        self.page_frame.config(width=w, height=h)
        self.page_frame.pack_propagate(False)
        self.canvas_desktop.config(scrollregion=(0, 0, w + 100, h + 100))
        self.center_page_on_desktop()
        self.on_margin_change()

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
            # Sfondo del desktop scuro, ma il foglio e l'editor rimangono bianchi con testo nero
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
        f, s = self.font_family.get(), self.font_size.get()
        self.editor.config(font=(f, s))

    def toggle_style(self, tag_name):
        try:
            start, end = self.editor.index("sel.first"), self.editor.index("sel.last")
            if tag_name in self.editor.tag_names("sel.first"): self.editor.tag_remove(tag_name, start, end)
            else: self.editor.tag_add(tag_name, start, end)
        except tk.TclError: pass

    def toggle_bold(self): self.toggle_style("bold")
    def toggle_italic(self): self.toggle_style("italic")
    def toggle_underline(self): self.toggle_style("underline")

    def set_alignment(self, align):
        try:
            start, end = self.editor.index("sel.first"), self.editor.index("sel.last")
            for a in ["left", "center", "right"]: self.editor.tag_remove(a, start, end)
            self.editor.tag_add(align, start, end)
        except tk.TclError: pass

    def select_all(self):
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
        self.editor.delete("1.0", tk.END)
        self.pages_data = {1: ""}
        self.active_page_id = 1
        self.current_file = None
        self.embedded_images_cache.clear()
        self.update_pages_sidebar()
        self.root.title("Nuovo Documento - UniversalWriter")
        return "break"

    def save_file(self):
        if self.current_file: self.execute_save(self.current_file)
        else: self.save_as_file()
        return "break"

    def save_as_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf")
        if file_path: self.execute_save(file_path)
        return "break"

    def execute_save(self, file_path):
        self.save_current_page_state()
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(self.pages_data[self.active_page_id])
            self.current_file = file_path
            messagebox.showinfo("Salvataggio", "Salvataggio completato!")
        except Exception as e:
            messagebox.showerror("Errore", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = AdvancedTextEditor(root)
    root.mainloop()