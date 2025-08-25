import os
import sys
import math
import json
import shutil
import zipfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinter import font as tkfont
from PIL import Image, ImageTk
from pdf2image import convert_from_path
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib.utils import ImageReader

# ---------- helper resource path (PyInstaller safe) ----------
def resource_path(rel_path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, rel_path)
    return os.path.join(os.path.abspath("."), rel_path)

# =========================
# PDF CROP-PARAMETERS (jouw exacte waarden)
# =========================
PDF_COLS = 5
PDF_PHOTO_W = 236
PDF_PHOTO_H = 236
PDF_MARGIN_LEFT = 140
PDF_MARGIN_TOP = 319
PDF_H_SPACING = 40
PDF_V_SPACING = 88
PDF_DPI = 200  # hogere dpi = scherpere crop

# =========================
# UI/Render instellingen
# =========================
SEAT_MIN = 60
SEAT_MAX = 130
CAPTION_GAP = 8
INNER_PAD_X = 8
INNER_PAD_TOP = 8
INNER_PAD_BOTTOM = 12
SEAT_SPACING = 8
ROW_SPACING = 28
BANK_SPACING = 24

PAGE_MARGIN_LR = 28
PAGE_MARGIN_TOP = 56
PAGE_MARGIN_BOTTOM = 24

FONT_MAX = 12
FONT_MIN = 7

# =========================
# Layouts definitie (incl. default Eigen opstelling)
# =========================
LAYOUTS = {
    "Klas 1 — 5 rijen × 3 banken × 2 stoelen": {
        "regular": True, "rows": 5, "banks": 3, "seats": 2, "orientation": "portrait"
    },
    "Klas 2 — 4 rijen × 4 banken × 2 stoelen": {
        "regular": True, "rows": 4, "banks": 4, "seats": 2, "orientation": "landscape"
    },
    "Labo T117 — 4 rijen × 2 banken × 4 stoelen": {
        "regular": True, "rows": 4, "banks": 2, "seats": 4, "orientation": "landscape"
    },
    "Fysica T121 — top 4 banken gecentreerd, dan 3×3 banken (3 stoelen)": {
        "regular": False,
        "pattern": [[4], [3,3,3], [3,3,3], [3,3,3]],
        "orientation": "landscape",
        "center_first_row": True
    },
    "Eigen opstelling": {
        "regular": True, "rows": 4, "banks": 3, "seats": 2, "orientation": "portrait"
    }
}

def parse_pattern_text(raw: str):
    if raw is None:
        raise ValueError("Leeg patroon")
    s = raw.strip()
    if not s:
        raise ValueError("Leeg patroon")
    s = s.replace("], [", ";").replace("],[", ";").replace("][", ";")
    s = s.replace("[", "").replace("]", "")
    s = s.replace("\n", ";")
    parts = [p.strip() for p in s.split(";") if p.strip()]
    pattern = []
    for p in parts:
        nums = [x.strip() for x in p.split(",") if x.strip()]
        if not nums:
            raise ValueError("Lege rij in patroon")
        row = []
        for n in nums:
            if not n.isdigit():
                raise ValueError(f"Niet-numerieke waarde in patroon: '{n}'")
            v = int(n)
            if v <= 0:
                raise ValueError("Alle aantallen moeten > 0 zijn")
            row.append(v)
        pattern.append(row)
    return pattern

def safe_filename(s):
    keep = "-_.() abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    return "".join(c for c in s if c in keep).replace(" ", "_")

class SeatPlanner:
    def __init__(self, root):
        self.root = root
        self.root.title("Klasopstelling")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Force Helvetica overal
        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(family="Helvetica", size=10)
        try:
            tkfont.nametofont("TkTextFont").configure(family="Helvetica", size=10)
        except Exception:
            pass
        self.root.option_add("*Font", ("Helvetica", 10))

        # Data containers
        self.students = []   # list of dicts: name,pil,tk,slot,img_id,text_id,font_size,source,pdf_index,img_filename
        self.base_slots = []     # logical geometry used for export
        self.base_bank_rects = []
        self.slots = []          # visual (scaled) geometry
        self.bank_rects = []
        self.page_size = A4
        self.seat_size = 100
        self.zoom_level = 1.0
        self.drag = {"student": None, "offset": (0,0)}

        # icons
        self.load_icons()

        # ---------- Top row: buttons ----------
        top_buttons = tk.Frame(root)
        top_buttons.pack(side=tk.TOP, fill=tk.X, padx=6, pady=(6,4))

        btn_font = ("Helvetica", 10)
        btn_font_bold = ("Helvetica", 10, "bold")

        tk.Button(top_buttons, text="  Foto's uit map", image=self.ic_camera, compound="left",
                  command=self.load_from_folder, bg="#D7EEF9", fg="black", bd=1, relief="raised",
                  padx=6, pady=4, font=btn_font).pack(side=tk.LEFT, padx=4)

        tk.Button(top_buttons, text="  Foto's uit PDF", image=self.ic_camera, compound="left",
                  command=self.load_from_pdf_and_names, bg="#D7EEF9", fg="black", bd=1, relief="raised",
                  padx=6, pady=4, font=btn_font).pack(side=tk.LEFT, padx=4)

        tk.Button(top_buttons, text="  Shuffle", image=self.ic_herh, compound="left",
                  command=self.shuffle_students, bg="#FDE5C6", fg="black", bd=1, relief="raised",
                  padx=6, pady=4, font=btn_font).pack(side=tk.LEFT, padx=4)

        tk.Button(top_buttons, text="  Opslaan opstelling", image=self.ic_opslaan, compound="left",
                  command=self.save_seating, bg="#DFF3DF", fg="black", bd=1, relief="raised",
                  padx=6, pady=4, font=btn_font).pack(side=tk.LEFT, padx=4)

        tk.Button(top_buttons, text="  Open opstelling", image=self.ic_open, compound="left",
                  command=self.load_seating, bg="#DFF3DF", fg="black", bd=1, relief="raised",
                  padx=6, pady=4, font=btn_font).pack(side=tk.LEFT, padx=4)

        tk.Button(top_buttons, text="  Export PDF", image=self.ic_outbox, compound="left",
                  command=self.export_pdf, bg="#FFF7CC", fg="black", bd=1, relief="raised",
                  padx=6, pady=4, font=btn_font_bold).pack(side=tk.LEFT, padx=4)

        # Reset (clear board) button — rechts van Export PDF
        tk.Button(top_buttons, text="  Reset", image=self.ic_reset, compound="left",
                  command=self.reset_board, bg="#F8D7DA", fg="black", bd=1, relief="raised",
                  padx=6, pady=4, font=btn_font).pack(side=tk.LEFT, padx=4)

        # ---------- Under buttons: inputs + dropdown + zoom (moved here) ----------
        inputs_row = tk.Frame(root)
        inputs_row.pack(side=tk.TOP, fill=tk.X, padx=8, pady=(4,8))

        tk.Label(inputs_row, text="Klas:").pack(side=tk.LEFT)
        self.var_class = tk.StringVar(value="klas")
        ent_class = tk.Entry(inputs_row, textvariable=self.var_class, width=22)
        ent_class.pack(side=tk.LEFT, padx=(4, 12))
        self.var_class.trace_add("write", lambda *_: self.update_title())

        tk.Label(inputs_row, text="Lokaal:").pack(side=tk.LEFT)
        self.var_room = tk.StringVar(value="lokaal")
        ent_room = tk.Entry(inputs_row, textvariable=self.var_room, width=12)
        ent_room.pack(side=tk.LEFT, padx=(4, 12))
        self.var_room.trace_add("write", lambda *_: self.update_title())

        tk.Label(inputs_row, text="Opstelling:").pack(side=tk.LEFT, padx=(12,4))
        self.var_layout = tk.StringVar(value=list(LAYOUTS.keys())[0])
        self.op_layout_menu = ttk.OptionMenu(inputs_row, self.var_layout, self.var_layout.get(), *LAYOUTS.keys(),
                                             command=lambda *_: self.set_layout())
        self.op_layout_menu.pack(side=tk.LEFT)

        ttk.Button(inputs_row, text="Eigen opstelling", command=self.custom_layout_popup).pack(side=tk.LEFT, padx=8)

        # Zoom controls (moved to second row)
        zoom_frame = tk.Frame(inputs_row)
        zoom_frame.pack(side=tk.RIGHT, padx=4)
        tk.Button(zoom_frame, text="−", width=3, command=lambda: self.zoom(0.9)).pack(side=tk.LEFT, padx=2)
        tk.Button(zoom_frame, text="100%", width=5, command=self.reset_zoom).pack(side=tk.LEFT, padx=2)
        tk.Button(zoom_frame, text="+", width=3, command=lambda: self.zoom(1.1)).pack(side=tk.LEFT, padx=2)

        # ---------- Scrollable Canvas ----------
        viewport_frame = tk.Frame(root)
        viewport_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=8)

        vbar = tk.Scrollbar(viewport_frame, orient=tk.VERTICAL)
        vbar.pack(side=tk.RIGHT, fill=tk.Y)
        hbar = tk.Scrollbar(viewport_frame, orient=tk.HORIZONTAL)
        hbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.canvas = tk.Canvas(viewport_frame, width=900, height=620, bg="white",
                                xscrollcommand=hbar.set, yscrollcommand=vbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        hbar.config(command=self.canvas.xview)
        vbar.config(command=self.canvas.yview)

        # enable mousewheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)      # Windows
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)        # older mac/linux
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

        # Context menu
        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label="Naam wijzigen", command=lambda: self.rename_selected())
        self.menu.add_command(label="Verwijder leerling", command=lambda: self.delete_selected())
        self.selected_student = None

        # bind clicks
        self.canvas.bind("<Button-3>", self.on_right_click)
        self.canvas.bind("<Double-Button-1>", self.on_double_click)

        # Init layout
        self.set_layout(initial=True)

    # ---------------- icons loader ----------------
    def load_icons(self):
        size = 24
        mapping = {
            "ic_tools": "icons/tools.png",
            "ic_camera": "icons/camera.png",
            "ic_herh": "icons/herh.png",
            "ic_opslaan": "icons/opslaan.png",
            "ic_open": "icons/open.png",
            "ic_outbox": "icons/outbox.png",
            "ic_reset": "icons/reset.png"
        }
        for attr, rel in mapping.items():
            p = resource_path(rel)
            imgtk = None
            try:
                im = Image.open(p).convert("RGBA")
                im = im.resize((size, size), Image.LANCZOS)
                imgtk = ImageTk.PhotoImage(im)
            except Exception:
                im = Image.new("RGBA", (size, size), (220,220,220,0))
                imgtk = ImageTk.PhotoImage(im)
            setattr(self, attr, imgtk)

    # ---------------- Title & close ----------------
    def on_close(self):
        if self.students:
            resp = messagebox.askyesnocancel("Bevestig afsluiten",
                                             "Er staan nog foto's op het bord. Wilt u opslaan vóór afsluiten?\n\nJa = opslaan en afsluiten\nNee = afsluiten zonder opslaan\nAnnuleer = terug")
            if resp is None:
                return
            if resp is True:
                self.save_seating()
        # either no students or user chose to continue
        if messagebox.askyesno("Bevestig afsluiten", "Ben je zeker dat je wil afsluiten?"):
            self.root.destroy()

    def update_title(self):
        try:
            self.canvas.delete("title")
        except Exception:
            pass
        W,_ = self.page_size
        # draw title with zoom applied visually
        self.canvas.create_text((W/2)*self.zoom_level, 24*self.zoom_level,
                                text=f"Klas {self.var_class.get()} — Lokaal {self.var_room.get()}",
                                font=("Helvetica", int(16*self.zoom_level), "bold"), tags=("title",))

    # ---------------- Custom layout popup ----------------
    def custom_layout_popup(self):
        top = tk.Toplevel(self.root)
        top.title("Eigen opstelling")
        top.grab_set()

        existing = LAYOUTS.get("Eigen opstelling", None) or {"regular": True, "rows":4, "banks":3, "seats":2, "orientation":"portrait"}

        var_regular = tk.BooleanVar(value=existing.get("regular", True))
        frame_type = tk.Frame(top)
        frame_type.pack(fill="x", padx=8, pady=(6,0))
        tk.Radiobutton(frame_type, text="Regulier (rijen × banken × stoelen)", variable=var_regular, value=True).pack(anchor="w")
        tk.Radiobutton(frame_type, text="Onregelmatig (patroon per rij)", variable=var_regular, value=False).pack(anchor="w")

        frm_regular = tk.Frame(top)
        frm_regular.pack(fill="x", padx=8, pady=4)
        tk.Label(frm_regular, text="Rijen:").grid(row=0,column=0,sticky="e")
        ent_rows = tk.Entry(frm_regular, width=6); ent_rows.grid(row=0,column=1,padx=6)
        tk.Label(frm_regular, text="Banken per rij:").grid(row=0,column=2,sticky="e")
        ent_banks = tk.Entry(frm_regular, width=6); ent_banks.grid(row=0,column=3,padx=6)
        tk.Label(frm_regular, text="Zitplaatsen per bank:").grid(row=0,column=4,sticky="e")
        ent_seats = tk.Entry(frm_regular, width=6); ent_seats.grid(row=0,column=5,padx=6)

        tk.Label(top, text="Onregelmatig patroon (bv: [4],[3,3,3],[3,3,3]) of 4;3,3,3;3,3,3)").pack(anchor="w", padx=8, pady=(6,0))
        txt = tk.Text(top, width=50, height=5)
        txt.pack(padx=8, pady=4)

        tk.Label(top, text="Oriëntatie:").pack(anchor="w", padx=8, pady=(4,0))
        var_orient = tk.StringVar(value=existing.get("orientation","portrait"))
        tk.Radiobutton(top, text="Staand", variable=var_orient, value="portrait").pack(anchor="w", padx=8)
        tk.Radiobutton(top, text="Liggend", variable=var_orient, value="landscape").pack(anchor="w", padx=8)

        if existing:
            if existing.get("regular", True):
                ent_rows.insert(0, str(existing.get("rows",4)))
                ent_banks.insert(0, str(existing.get("banks",3)))
                ent_seats.insert(0, str(existing.get("seats",2)))
            else:
                txt.delete("1.0","end")
                txt.insert("1.0", ",".join(["[" + ",".join(map(str, row)) + "]" for row in existing.get("pattern", [[4],[3,3,3]])]))

        def on_ok():
            try:
                orientation = var_orient.get()
                if var_regular.get():
                    rows = int(ent_rows.get())
                    banks = int(ent_banks.get())
                    seats = int(ent_seats.get())
                    if rows<=0 or banks<=0 or seats<=0:
                        raise ValueError("Negatieve of nul waarden niet toegestaan")
                    LAYOUTS["Eigen opstelling"] = {"regular": True, "rows": rows, "banks": banks, "seats": seats, "orientation": orientation}
                else:
                    raw = txt.get("1.0","end").strip()
                    pattern = parse_pattern_text(raw)
                    LAYOUTS["Eigen opstelling"] = {"regular": False, "pattern": pattern, "orientation": orientation, "center_first_row": True}
                self.var_layout.set("Eigen opstelling")
                self.set_layout()
                top.destroy()
            except Exception as e:
                messagebox.showerror("Fout", f"Ongeldige invoer:\n{e}")

        btns = tk.Frame(top)
        btns.pack(pady=6)
        ttk.Button(btns, text="OK", command=on_ok).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="Annuleer", command=top.destroy).pack(side=tk.LEFT, padx=6)
        top.wait_window()

    # ---------------- Loading images ----------------
    def load_from_folder(self):
        folder = filedialog.askdirectory(title="Kies map met foto's (jpg/png)")
        if not folder:
            return
        files = [f for f in os.listdir(folder) if f.lower().endswith((".jpg",".jpeg",".png"))]
        files.sort()
        if not files:
            messagebox.showwarning("Geen foto's", "Geen jpg/png gevonden in de gekozen map.")
            return
        names = self.prompt_names_list(default_list=[os.path.splitext(f)[0] for f in files])
        for i,f in enumerate(files):
            path = os.path.join(folder,f)
            try:
                im = Image.open(path).convert("RGB")
            except Exception:
                continue
            pil_sq = self.crop_square(im)
            name = names[i] if i < len(names) else os.path.splitext(f)[0]
            self.students.append({
                "name": name, "pil": pil_sq, "tk": None, "slot": None,
                "img_id": None, "text_id": None, "font_size": FONT_MAX,
                "source": path, "pdf_index": None, "img_filename": None
            })
        self.reflow_after_data_change()

    def load_from_pdf_and_names(self):
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")], title="Kies PDF met foto's")
        if not pdf_path:
            return
        try:
            pages = convert_from_path(pdf_path, dpi=PDF_DPI)
        except Exception as e:
            messagebox.showerror("PDF fout", f"Kon PDF niet lezen:\n{e}")
            return
        page = pages[0].convert("RGB")
        N = simpledialog.askinteger("Aantal leerlingen", "Hoeveel leerlingen staan op de PDF?", minvalue=1, maxvalue=200, parent=self.root)
        if not N:
            return
        names = self.prompt_names_list(count=N)
        for i in range(N):
            r = i // PDF_COLS
            c = i % PDF_COLS
            x1 = PDF_MARGIN_LEFT + c * (PDF_PHOTO_W + PDF_H_SPACING)
            y1 = PDF_MARGIN_TOP + r * (PDF_PHOTO_H + PDF_V_SPACING)
            x2 = x1 + PDF_PHOTO_W
            y2 = y1 + PDF_PHOTO_H
            crop = page.crop((x1, y1, x2, y2))
            pil_sq = self.crop_square(crop)
            name = names[i] if i < len(names) else f"leerling_{i+1}"
            self.students.append({
                "name": name, "pil": pil_sq, "tk": None, "slot": None,
                "img_id": None, "text_id": None, "font_size": FONT_MAX,
                "source": pdf_path, "pdf_index": i, "img_filename": None
            })
        self.reflow_after_data_change()

    def prompt_names_list(self, count=None, default_list=None):
        top = tk.Toplevel(self.root)
        top.title("Namen plakken (één per lijn)")
        top.grab_set()
        tk.Label(top, text="Plak hier de namen (één per regel). Leeg laten is ook oké.").pack(padx=8, pady=6)
        txt = tk.Text(top, width=40, height=12)
        txt.pack(padx=8, pady=6)
        if default_list:
            txt.insert("1.0", "\n".join(default_list))
        elif count:
            txt.insert("1.0", "\n".join([f"leerling_{i+1}" for i in range(count)]))
        result = {"names": []}
        def ok():
            content = txt.get("1.0","end").strip()
            result["names"] = [line.strip() for line in content.splitlines() if line.strip()]
            top.destroy()
        def cancel():
            result["names"] = []
            top.destroy()
        btns = tk.Frame(top); btns.pack(pady=6)
        ttk.Button(btns, text="OK", command=ok).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="Annuleren", command=cancel).pack(side=tk.LEFT, padx=6)
        top.wait_window()
        return result["names"] or []

    # ---------------- Helpers ----------------
    def crop_square(self, pil_img):
        w,h = pil_img.size
        side = min(w,h)
        L = (w - side)//2
        T = (h - side)//2
        return pil_img.crop((L,T,L+side,T+side))

    def set_layout(self, initial=False):
        chosen = self.var_layout.get()
        if chosen not in LAYOUTS:
            chosen = next(iter(LAYOUTS.keys()))
            self.var_layout.set(chosen)
        cfg = LAYOUTS[chosen]
        orient = cfg.get("orientation","portrait")
        self.page_size = portrait(A4) if orient=="portrait" else landscape(A4)
        # recompute geometry and redraw (keeps current zoom_level)
        self.compute_geometry_and_draw_static()
        self.build_tk_thumbs()
        self.auto_assign_students()
        self.draw_students()

    def compute_geometry_and_draw_static(self):
        """
        Compute both base (export) geometry based on logical seat_size,
        and display geometry according to zoom_level.
        """
        # Clear canvas items
        self.canvas.delete("all")
        self.base_slots.clear()
        self.base_bank_rects.clear()
        self.slots.clear()
        self.bank_rects.clear()

        W, H = self.page_size
        cfg = LAYOUTS[self.var_layout.get()]
        regular = cfg.get("regular", True)

        if regular:
            rows = cfg["rows"]
            banks_per_row = [cfg["banks"]]*rows
            seats_lookup = lambda r,c: cfg["seats"]
        else:
            pattern = cfg["pattern"]
            banks_per_row = [len(row) for row in pattern]
            rows = len(pattern)
            def seats_lookup(r,c):
                return pattern[r][c]

        max_banks = max(banks_per_row) if banks_per_row else 0

        max_seats_in_widest_row = 0
        for r in range(rows):
            seats_list = [seats_lookup(r, c) for c in range(banks_per_row[r])]
            max_seats_in_widest_row = max(max_seats_in_widest_row, max(seats_list) if seats_list else 0)

        avail_w = W - PAGE_MARGIN_LR*2 - (max_banks-1)*BANK_SPACING
        seats_per_bank_for_width = cfg["seats"] if regular else (max_seats_in_widest_row or 1)
        seat_by_w = (avail_w / max_banks - 2*INNER_PAD_X - (seats_per_bank_for_width-1)*SEAT_SPACING) / max(seats_per_bank_for_width,1)

        font_est = 14
        avail_h = H - PAGE_MARGIN_TOP - PAGE_MARGIN_BOTTOM - (rows-1)*ROW_SPACING
        seat_by_h = avail_h/rows - (INNER_PAD_TOP + CAPTION_GAP + font_est + INNER_PAD_BOTTOM)

        # base logical seat_size used for export
        self.seat_size = int(max(SEAT_MIN, min(SEAT_MAX, seat_by_w, seat_by_h)))

        def bank_w_base(seats):
            return int(2*INNER_PAD_X + seats*self.seat_size + (seats-1)*SEAT_SPACING)
        bank_h_base = int(INNER_PAD_TOP + self.seat_size + CAPTION_GAP + font_est + INNER_PAD_BOTTOM)

        # Build base geometry
        y_base = PAGE_MARGIN_TOP + max(0, (H - PAGE_MARGIN_TOP - PAGE_MARGIN_BOTTOM - (rows*bank_h_base + (rows-1)*ROW_SPACING))//2)
        title_y = 24
        if y_base < title_y + 10:
            y_base = title_y + 16

        for r in range(rows):
            row_banks = banks_per_row[r]
            if regular:
                row_bank_widths = [bank_w_base(cfg["seats"]) for _ in range(row_banks)]
            else:
                row_bank_widths = [bank_w_base(seats_lookup(r, c)) for c in range(row_banks)]
            row_total_w = sum(row_bank_widths) + (row_banks-1)*BANK_SPACING
            x_base = PAGE_MARGIN_LR + (W - 2*PAGE_MARGIN_LR - row_total_w)//2
            for b in range(row_banks):
                seats = seats_lookup(r,b) if not regular else cfg["seats"]
                bw = row_bank_widths[b]
                x0b, y0b = x_base, y_base
                x1b, y1b = x0b + bw, y0b + bank_h_base
                self.base_bank_rects.append((x0b, y0b, x1b, y1b))
                sx = x0b + INNER_PAD_X
                sy = y0b + INNER_PAD_TOP
                for s in range(seats):
                    self.base_slots.append({
                        "x": sx, "y": sy, "w": self.seat_size, "h": self.seat_size,
                        "cx": sx + self.seat_size/2, "cy": sy + self.seat_size/2
                    })
                    sx += self.seat_size + SEAT_SPACING
                x_base += bw + BANK_SPACING
            y_base += bank_h_base + ROW_SPACING

        # Build visual geometry using zoom_level
        vs = max(4, int(self.seat_size * self.zoom_level))
        def bank_w_disp(seats):
            return int(2*INNER_PAD_X + seats*vs + (seats-1)*SEAT_SPACING)
        bank_h_disp = int(INNER_PAD_TOP + vs + CAPTION_GAP + font_est + INNER_PAD_BOTTOM)

        y_disp = PAGE_MARGIN_TOP*self.zoom_level + max(0, int((H*self.zoom_level - (PAGE_MARGIN_TOP*self.zoom_level + PAGE_MARGIN_BOTTOM*self.zoom_level) - (rows*bank_h_disp + (rows-1)*int(ROW_SPACING*self.zoom_level)))//2))
        if y_disp < title_y*self.zoom_level + 10:
            y_disp = title_y*self.zoom_level + 16

        for r in range(rows):
            row_banks = banks_per_row[r]
            if regular:
                row_bank_widths = [bank_w_disp(cfg["seats"]) for _ in range(row_banks)]
            else:
                row_bank_widths = [bank_w_disp(seats_lookup(r, c)) for c in range(row_banks)]
            row_total_w = sum(row_bank_widths) + (row_banks-1)*int(BANK_SPACING*self.zoom_level)
            x_disp = int(PAGE_MARGIN_LR*self.zoom_level + (W*self.zoom_level - 2*PAGE_MARGIN_LR*self.zoom_level - row_total_w)//2)
            for b in range(row_banks):
                seats = seats_lookup(r,b) if not regular else cfg["seats"]
                bw = row_bank_widths[b]
                x0, y0 = x_disp, int(y_disp)
                x1, y1 = x0 + bw, y0 + bank_h_disp
                self.canvas.create_rectangle(x0, y0, x1, y1, outline="black", tags=("static","bank"))
                self.bank_rects.append((x0, y0, x1, y1))
                sx = x0 + INNER_PAD_X
                sy = y0 + INNER_PAD_TOP
                for s in range(seats):
                    self.canvas.create_rectangle(sx, sy, sx+vs, sy+vs, outline="#999", dash=(2,2), tags=("static","seatbox"))
                    self.slots.append({
                        "x": sx, "y": sy, "w": vs, "h": vs,
                        "cx": sx + vs/2, "cy": sy + vs/2
                    })
                    sx += vs + int(SEAT_SPACING*self.zoom_level)
                x_disp += bw + int(BANK_SPACING*self.zoom_level)
            y_disp += bank_h_disp + int(ROW_SPACING*self.zoom_level)

        # title and scrollregion
        self.update_title()
        bbox = self.canvas.bbox("all")
        if bbox:
            self.canvas.config(scrollregion=bbox)
        else:
            self.canvas.config(scrollregion=(0,0,W*self.zoom_level,H*self.zoom_level))

    def build_tk_thumbs(self):
        vs = max(4, int(self.seat_size * self.zoom_level))
        for s in self.students:
            try:
                thumb = s["pil"].resize((vs, vs), Image.LANCZOS)
            except Exception:
                thumb = Image.new("RGB", (vs, vs), (240,240,240))
            s["tk"] = ImageTk.PhotoImage(thumb)

    def auto_assign_students(self):
        used = set(s["slot"] for s in self.students if s["slot"] is not None and isinstance(s["slot"], int) and s["slot"] < len(self.slots))
        free = [i for i in range(len(self.slots)) if i not in used]
        for s in self.students:
            if s["slot"] is None or not isinstance(s["slot"], int) or s["slot"] >= len(self.slots):
                if free:
                    s["slot"] = free.pop(0)
                else:
                    s["slot"] = None

    def draw_students(self):
        self.canvas.delete("student")
        self.canvas.delete("photo")
        for s in self.students:
            if s["slot"] is None or not isinstance(s["slot"], int) or s["slot"] >= len(self.slots):
                continue
            slot = self.slots[s["slot"]]
            x, y = slot["x"], slot["y"]
            s["img_id"] = self.canvas.create_image(x, y, image=s["tk"], anchor="nw", tags=("photo","student"))
            font_size = self.fit_font_size(s["name"], max_width=int(slot["w"]*0.95))
            s["font_size_display"] = font_size
            s["text_id"] = self.canvas.create_text(x + slot["w"]/2, y + slot["h"] + CAPTION_GAP, text=s["name"],
                                                   font=("Helvetica", font_size, "bold"), anchor="n", tags=("student","label"))
            # bindings
            self.canvas.tag_bind(s["img_id"], "<Button-1>", self.on_drag_start)
            self.canvas.tag_bind(s["img_id"], "<B1-Motion>", self.on_drag_move)
            self.canvas.tag_bind(s["img_id"], "<ButtonRelease-1>", self.on_drag_end)
            self.canvas.tag_bind(s["img_id"], "<Double-Button-1>", self.on_double_click)
            self.canvas.tag_bind(s["text_id"], "<Double-Button-1>", self.on_double_click)

    def fit_font_size(self, text, max_width):
        size = FONT_MAX
        test_font = tkfont.Font(family="Helvetica", size=size, weight="bold")
        while size > FONT_MIN and test_font.measure(text) > max_width:
            size -= 1
            test_font.configure(size=size)
        return size

    def reflow_after_data_change(self):
        self.build_tk_thumbs()
        self.auto_assign_students()
        self.draw_students()

    # ---------------- Drag & Drop ----------------
    def find_student_by_img(self, item_id):
        for s in self.students:
            if s.get("img_id") == item_id:
                return s
        return None

    def on_drag_start(self, event):
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        item = self.canvas.find_withtag("current")
        if not item: return
        img_id = item[0]
        st = self.find_student_by_img(img_id)
        if not st: return
        self.drag["student"] = st
        bbox = self.canvas.bbox(img_id)
        self.drag["offset"] = (cx - bbox[0], cy - bbox[1])
        self.canvas.tag_raise(st["img_id"])
        if st.get("text_id"): self.canvas.tag_raise(st["text_id"])

    def on_drag_move(self, event):
        st = self.drag["student"]
        if not st: return
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        dx, dy = self.drag["offset"]
        nx, ny = cx - dx, cy - dy
        self.canvas.coords(st["img_id"], nx, ny)
        if st.get("text_id"):
            w = self.slots[st["slot"]]["w"] if (st["slot"] is not None and st["slot"] < len(self.slots)) else int(self.seat_size * self.zoom_level)
            self.canvas.coords(st["text_id"], nx + w/2, ny + w + CAPTION_GAP)

    def on_drag_end(self, event):
        st = self.drag["student"]
        if not st: return
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        nearest, bestd2 = None, 1e18
        for i, slot in enumerate(self.slots):
            dx = cx - slot["cx"]
            dy = cy - slot["cy"]
            d2 = dx*dx + dy*dy
            if d2 < bestd2:
                bestd2 = d2
                nearest = i
        snap = max(100, int(self.seat_size * 1.2 * self.zoom_level))
        if bestd2 > snap*snap:
            self.refresh_positions()
            self.drag["student"] = None
            return

        origin = st["slot"]
        target = nearest
        if target is None or target >= len(self.slots):
            self.refresh_positions()
            self.drag["student"] = None
            return

        other = None
        for s2 in self.students:
            if s2 is not st and s2["slot"] == target:
                other = s2
                break

        if other is None:
            st["slot"] = target
        else:
            st["slot"], other["slot"] = other["slot"], st["slot"]

        self.refresh_positions()
        self.drag["student"] = None

    def refresh_positions(self):
        for s in self.students:
            if s.get("slot") is None or not isinstance(s["slot"], int) or s["slot"] >= len(self.slots): continue
            slot = self.slots[s["slot"]]
            x, y = slot["x"], slot["y"]
            self.canvas.coords(s["img_id"], x, y)
            if s.get("text_id"):
                self.canvas.coords(s["text_id"], x + slot["w"]/2, y + slot["h"] + CAPTION_GAP)

    def shuffle_students(self):
        import random
        random.shuffle(self.students)
        for i, s in enumerate(self.students):
            s["slot"] = i if i < len(self.slots) else None
        self.draw_students()

    # ---------------- Contextmenu & edit name / delete ----------------
    def hit_student(self, cx, cy):
        items = self.canvas.find_overlapping(cx-1, cy-1, cx+1, cy+1)
        for it in reversed(items):
            for s in self.students:
                if it == s.get("img_id") or it == s.get("text_id"):
                    return s
        return None

    def on_right_click(self, event):
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        self.selected_student = self.hit_student(cx, cy)
        if self.selected_student:
            self.menu.tk_popup(event.x_root, event.y_root)

    def on_double_click(self, event):
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        s = self.hit_student(cx, cy)
        if s:
            self.edit_name_dialog(s)

    def rename_selected(self):
        if self.selected_student:
            self.edit_name_dialog(self.selected_student)
            self.selected_student = None

    def edit_name_dialog(self, student):
        top = tk.Toplevel(self.root)
        top.title("Naam wijzigen")
        top.grab_set()
        tk.Label(top, text="Nieuwe naam:").pack(padx=8, pady=(8,4))
        var = tk.StringVar(value=student["name"])
        ent = tk.Entry(top, textvariable=var, width=30); ent.pack(padx=8, pady=6); ent.focus_set()
        def ok():
            student["name"] = var.get().strip() or student["name"]
            new_size = self.fit_font_size(student["name"], max_width=int(self.seat_size * self.zoom_level * 0.95))
            student["font_size"] = new_size
            if student.get("text_id"):
                self.canvas.itemconfig(student["text_id"], text=student["name"], font=("Helvetica", new_size, "bold"))
            top.destroy()
        ttk.Button(top, text="OK", command=ok).pack(pady=(4,10))
        top.bind("<Return>", lambda e: ok())

    def delete_selected(self):
        if not self.selected_student:
            return
        s = self.selected_student
        if s.get("img_id"): self.canvas.delete(s["img_id"])
        if s.get("text_id"): self.canvas.delete(s["text_id"])
        self.students.remove(s)
        self.selected_student = None
        self.reflow_after_data_change()

    # ---------------- Save / Load seating (with images zipped) ----------------
    def save_seating(self):
        if not self.students:
            messagebox.showwarning("Opslaan", "Geen leerlingen om op te slaan.")
            return
        fpath = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON","*.json")], title="Bewaar opstelling")
        if not fpath:
            return

        base, _ = os.path.splitext(fpath)
        assets_dir = base + "_assets"
        zip_path = base + "_assets.zip"
        try:
            if os.path.isdir(assets_dir):
                shutil.rmtree(assets_dir)
            os.makedirs(assets_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Fout", f"Kan assets-map niet aanmaken:\n{e}")
            return

        students_meta = []
        for i, s in enumerate(self.students):
            fname = f"{i}_{safe_filename(s['name'])}.png"
            save_path = os.path.join(assets_dir, fname)
            try:
                s["pil"].save(save_path, format="PNG")
            except Exception:
                Image.new("RGB", (self.seat_size, self.seat_size), (240,240,240)).save(save_path, format="PNG")
            students_meta.append({
                "name": s["name"],
                "slot": s["slot"],
                "source": s.get("source"),
                "pdf_index": s.get("pdf_index"),
                "font_size": s.get("font_size", FONT_MAX),
                "img_filename": fname
            })

        data = {
            "class": self.var_class.get(),
            "room": self.var_room.get(),
            "layout": self.var_layout.get(),
            "custom_layout": LAYOUTS.get("Eigen opstelling"),
            "students": students_meta
        }

        # write JSON and create ZIP of assets, then remove the assets_dir
        try:
            with open(fpath, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            # create zip
            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for fname in os.listdir(assets_dir):
                    zf.write(os.path.join(assets_dir, fname), arcname=fname)
            # remove the temporary assets_dir
            shutil.rmtree(assets_dir)
            messagebox.showinfo("Opslaan", f"Opstelling opgeslagen in:\n{fpath}\n(en assets gecomprimeerd in {os.path.basename(zip_path)})")
        except Exception as e:
            messagebox.showerror("Fout", f"Kon niet opslaan:\n{e}")

    def load_seating(self):
        # if there are existing students, ask user whether to save before opening
        if self.students:
            resp = messagebox.askyesnocancel("Open opstelling",
                                             "Er staan foto's op het bord. Wilt u opslaan vóór openen?\n\nJa = opslaan en openen\nNee = openen zonder opslaan\nAnnuleer = stoppen")
            if resp is None:
                return
            if resp is True:
                self.save_seating()
        # proceed: ask user for JSON file
        fpath = filedialog.askopenfilename(filetypes=[("JSON","*.json")], title="Open opstelling")
        if not fpath:
            return
        base, _ = os.path.splitext(fpath)
        assets_dir = base + "_assets"
        zip_path = base + "_assets.zip"
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Fout", f"Kon bestand niet lezen:\n{e}")
            return

        # extract zip if present
        try:
            if os.path.isfile(zip_path):
                if os.path.isdir(assets_dir):
                    shutil.rmtree(assets_dir)
                os.makedirs(assets_dir, exist_ok=True)
                with zipfile.ZipFile(zip_path, "r") as zf:
                    zf.extractall(assets_dir)
        except Exception as e:
            messagebox.showwarning("Assets", f"Fout bij uitpakken assets: {e}")

        # restore custom layout if present
        custom = data.get("custom_layout")
        if custom and isinstance(custom, dict) and "regular" in custom:
            LAYOUTS["Eigen opstelling"] = custom

        if "class" in data: self.var_class.set(data["class"])
        if "room" in data: self.var_room.set(data["room"])
        if "layout" in data:
            layout_name = data["layout"]
            if layout_name not in LAYOUTS:
                layout_name = list(LAYOUTS.keys())[0]
            self.var_layout.set(layout_name)

        saved_students = data.get("students", [])

        # CLEAR current state completely (user chose to open a file)
        self.students = []
        self.canvas.delete("all")
        self.base_slots.clear()
        self.base_bank_rects.clear()
        self.slots.clear()
        self.bank_rects.clear()

        # Rebuild base layout first to know slots
        self.set_layout()

        new_students = []
        for i, meta in enumerate(saved_students):
            img_pil = None
            imgfile = meta.get("img_filename")
            if imgfile:
                p = os.path.join(assets_dir, imgfile)
                if os.path.isfile(p):
                    try:
                        img_pil = Image.open(p).convert("RGB")
                    except Exception:
                        img_pil = None
            # fallback: try to recover from source (image file or pdf)
            if img_pil is None and meta.get("source"):
                src = meta["source"]
                if os.path.isfile(src) and src.lower().endswith(".pdf") and isinstance(meta.get("pdf_index"), int):
                    try:
                        pages = convert_from_path(src, dpi=PDF_DPI)
                        page = pages[0].convert("RGB")
                        iidx = meta["pdf_index"]
                        r = iidx // PDF_COLS
                        c = iidx % PDF_COLS
                        x1 = PDF_MARGIN_LEFT + c * (PDF_PHOTO_W + PDF_H_SPACING)
                        y1 = PDF_MARGIN_TOP + r * (PDF_PHOTO_H + PDF_V_SPACING)
                        x2 = x1 + PDF_PHOTO_W
                        y2 = y1 + PDF_PHOTO_H
                        crop = page.crop((x1,y1,x2,y2))
                        img_pil = self.crop_square(crop)
                    except Exception:
                        img_pil = None
                elif os.path.isfile(src):
                    try:
                        img_pil = Image.open(src).convert("RGB")
                        img_pil = self.crop_square(img_pil)
                    except Exception:
                        img_pil = None
            if img_pil is None:
                img_pil = Image.new("RGB", (self.seat_size, self.seat_size), (240,240,240))
            name = meta.get("name", f"leerling_{i+1}")
            new_students.append({
                "name": name, "pil": img_pil, "tk": None, "slot": meta.get("slot"),
                "img_id": None, "text_id": None, "font_size": meta.get("font_size", FONT_MAX),
                "source": meta.get("source"), "pdf_index": meta.get("pdf_index"), "img_filename": meta.get("img_filename")
            })
        # set students and redraw
        self.students = new_students
        self.reflow_after_data_change()

    # ---------------- Reset board ----------------
    def reset_board(self):
        if not self.students:
            # nothing to do, but confirm anyway (as requested)
            if not messagebox.askyesno("Reset bord", "Het bord is leeg. Wil je het toch resetten?"):
                return
        else:
            if not messagebox.askyesno("Reset bord", "Ben je zeker dat je wil resetten? Dit verwijdert alle foto's op het bord."):
                return
        # clear students, keep layout
        for s in list(self.students):
            if s.get('img_id'): self.canvas.delete(s.get('img_id'))  # keep safe deletion
        self.students = []
        self.reflow_after_data_change()
        # ensure banks/slots remain visible: recompute layout
        self.set_layout()

    # ---------------- Export PDF ----------------
    def export_pdf(self):
        if not self.students:
            messagebox.showwarning("Geen leerlingen", "Er zijn geen leerlingen om te exporteren.")
            return
        fpath = filedialog.asksaveasfilename(defaultextension=".pdf",
                                             filetypes=[("PDF", "*.pdf")],
                                             title="Bewaar als PDF",
                                             initialfile=f"{self.var_class.get()}_{self.var_room.get()}.pdf")
        if not fpath:
            return

        cfg = LAYOUTS[self.var_layout.get()]
        page = portrait(A4) if cfg.get("orientation","portrait")=="portrait" else landscape(A4)
        W,H = page
        c = pdfcanvas.Canvas(fpath, pagesize=page)

        # Titel
        c.setFont("Helvetica-Bold", 20)
        title = f"Klas {self.var_class.get()} — Lokaal {self.var_room.get()}"
        c.drawCentredString(W/2, H-36, title)

        # Banken (use base bank rects)
        c.setLineWidth(1)
        for (x0,y0,x1,y1) in self.base_bank_rects:
            c.rect(x0, H - y1, x1-x0, y1-y0, stroke=1, fill=0)

        # Stoel placeholders (dotted)
        c.setDash(2,2)
        for slot in self.base_slots:
            x,y,w,h = slot["x"], slot["y"], slot["w"], slot["h"]
            c.rect(x, H-(y+h), w, h, stroke=1, fill=0)
        c.setDash()

        oversample = 2
        for s in self.students:
            if s["slot"] is None or not isinstance(s["slot"], int) or s["slot"] >= len(self.base_slots): continue
            slot = self.base_slots[s["slot"]]
            x, y = slot["x"], slot["y"]
            draw_w = draw_h = slot["w"]
            thumb = s["pil"].resize((max(1, draw_w*oversample), max(1, draw_h*oversample)), Image.LANCZOS)
            ir = ImageReader(thumb)
            c.drawImage(ir, x, H-(y+draw_h), width=draw_w, height=draw_h, preserveAspectRatio=True, mask='auto')
            ui_font_size = int(s.get("font_size", FONT_MAX))
            ui_font_size = max(FONT_MIN, min(FONT_MAX, ui_font_size))
            c.setFont("Helvetica-Bold", ui_font_size)
            c.drawCentredString(x + draw_w/2, H-(y+draw_h+CAPTION_GAP+12), s["name"])

        c.showPage()
        c.save()
        messagebox.showinfo("Export", f"PDF opgeslagen:\n{fpath}")

    # ---------------- Zoom helpers ----------------
    def zoom(self, factor):
        new_z = self.zoom_level * factor
        new_z = max(0.5, min(2.0, new_z))
        self.zoom_level = new_z
        self.compute_geometry_and_draw_static()
        self.build_tk_thumbs()
        self.draw_students()

    def reset_zoom(self):
        self.zoom_level = 1.0
        self.compute_geometry_and_draw_static()
        self.build_tk_thumbs()
        self.draw_students()

    # ---------------- Mousewheel ----------------
    def _on_mousewheel(self, event):
        if hasattr(event, "delta") and event.delta:
            delta = -1 * int(event.delta/120)
        elif getattr(event, "num", None) == 4:
            delta = -1
        elif getattr(event, "num", None) == 5:
            delta = 1
        else:
            delta = 0
        self.canvas.yview_scroll(delta, "units")

# ---------------- Run ----------------
if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.state("zoomed")
    except Exception:
        pass
    app = SeatPlanner(root)
    root.mainloop()
