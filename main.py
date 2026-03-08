# -*- coding: utf-8 -*-
"""
main.py  --  PC Health AI  (Redesigned UI)
  - Sidebar navigation replacing flat tab bar
  - 6 clean sections: Dashboard · Hardware · Network · Processes · Maintenance · Optimise
  - App name prominent in header
  - "Created By: CRiMiNAL" in footer
  - All original functionality preserved
"""

import sys, os, json, threading, tkinter as tk
from tkinter import messagebox

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import customtkinter as ctk

from system_info      import SystemInfo
from update_manager   import UpdateManager
from diagnosis_engine import DiagnosisEngine

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
APP_DIR     = os.path.dirname(os.path.abspath(__file__))

# ── Midnight Cyber palette ────────────────────────────────────────────────────
C_BG      = "#070b14"
C_SIDEBAR = "#050810"
C_PANEL   = "#0b1020"
C_CARD    = "#0f1628"
C_CARD2   = "#151e33"
C_BORDER  = "#1c2847"
C_ACCENT  = "#00d4ff"
C_GREEN   = "#00f598"
C_RED     = "#ff3355"
C_ORANGE  = "#ff8c00"
C_YELLOW  = "#ffd700"
C_HDR     = "#05080f"
C_PURPLE  = "#bd69f8"
C_PINK    = "#f040b0"
C_TEAL    = "#00e5cc"
C_NAVSEL  = "#0a1830"

def F (s=14, b=True): return ctk.CTkFont(size=s, weight="bold" if b else "normal")
def FT(s=20):         return F(s)
def FH(s=15):         return F(s)
def FL(s=13):         return F(s)
def FV(s=13):         return F(s, False)
def FB(s=13):         return F(s)
def FM(s=12):         return F(s)

def _is_admin() -> bool:
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


# ══════════════════════════════════════════════════════════════════════════════
class PCHealthApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("PC Health AI")
        self.geometry("1440x920")
        self.minsize(1150, 720)
        self.configure(fg_color=C_BG)

        self.sys_info    = SystemInfo()
        self.upd_manager = UpdateManager()
        self.diag_engine = DiagnosisEngine()
        self.system_data: dict | None = None
        self._is_admin   = _is_admin()

        self._upd_vars:   dict[str, tk.BooleanVar] = {}
        self._upd_rows:   list[dict]                = []
        self._game_vars:  dict[str, tk.BooleanVar] = {}
        self._bloat_vars: dict[str, tk.BooleanVar] = {}
        self._last_gpu_usage: dict = {"gpu_load_pct": None, "vram_ctrl_pct": None, "source": "none"}
        self._proc_data: dict = {"by_cpu": [], "by_ram": []}

        self._active_nav      = None
        self._nav_buttons:    dict[str, ctk.CTkButton]  = {}
        self._content_frames: dict[str, ctk.CTkFrame]   = {}

        self._build_ui()
        self.after(300, lambda: threading.Thread(target=self._scan_system, daemon=True).start())

    # ── config ─────────────────────────────────────────────────────────────────
    def _load_cfg(self):
        if os.path.isfile(CONFIG_FILE):
            try:
                with open(CONFIG_FILE) as f: return json.load(f)
            except Exception: pass
        return {}

    def _save_cfg(self, data):
        d = self._load_cfg(); d.update(data)
        with open(CONFIG_FILE, "w") as f: json.dump(d, f, indent=2)

    # ══════════════════════════════════════════════════════════════════════════
    # ROOT UI SKELETON
    # ══════════════════════════════════════════════════════════════════════════
    def _build_ui(self):
        # row 0 = header, row 1 = main area (weight=1), row 2 = footer
        self.grid_columnconfigure(0, weight=0, minsize=230)   # sidebar
        self.grid_columnconfigure(1, weight=1)   # content
        self.grid_rowconfigure(1, weight=1)

        self._build_header()   # row 0, colspan 2
        self._build_sidebar()  # row 1, col 0
        self._build_content()  # row 1, col 1
        self._build_footer()   # row 2, colspan 2

        # Navigate to dashboard by default
        self._nav_select("dashboard")

    # ── HEADER ─────────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr = ctk.CTkFrame(self, height=72, corner_radius=0, fg_color=C_HDR)
        hdr.grid(row=0, column=0, columnspan=2, sticky="ew")
        # 3-column layout: left actions | center logo | right actions
        hdr.grid_columnconfigure(0, weight=1)
        hdr.grid_columnconfigure(1, weight=0)   # logo — natural size
        hdr.grid_columnconfigure(2, weight=1)
        hdr.grid_propagate(False)

        # ── Left: status label ────────────────────────────────────────────────
        self._status_lbl = ctk.CTkLabel(hdr, text="Initialising…",
                                        font=FM(12), text_color="gray45", anchor="e")
        self._status_lbl.grid(row=0, column=0, padx=(18, 20), sticky="e")

        # ── Center: App name ──────────────────────────────────────────────────
        logo = ctk.CTkFrame(hdr, fg_color="transparent")
        logo.grid(row=0, column=1, sticky="ns", pady=0)
        ctk.CTkFrame(logo, width=4, corner_radius=2, fg_color=C_ACCENT).pack(
            side="left", fill="y", pady=14, padx=(0, 14))
        txt = ctk.CTkFrame(logo, fg_color="transparent")
        txt.pack(side="left", fill="y")
        ctk.CTkLabel(txt, text="🖥  PC HEALTH AI",
                     font=ctk.CTkFont(size=23, weight="bold"),
                     text_color=C_ACCENT).pack(anchor="center", pady=(16, 0))
        ctk.CTkLabel(txt, text="SYSTEM MONITOR & OPTIMIZER",
                     font=ctk.CTkFont(size=9, weight="bold"),
                     text_color="gray30").pack(anchor="center")

        # ── Right: Action buttons ─────────────────────────────────────────────
        right = ctk.CTkFrame(hdr, fg_color="transparent")
        right.grid(row=0, column=2, padx=(20, 18), sticky="ns")
        right.grid_rowconfigure(0, weight=1)
        inner = ctk.CTkFrame(right, fg_color="transparent")
        inner.grid(row=0, column=0, sticky="e")
        right.grid_columnconfigure(0, weight=1)
        self._scan_btn = ctk.CTkButton(
            inner, text="🔍  Scan System", font=FB(13),
            command=lambda: threading.Thread(target=self._scan_system, daemon=True).start(),
            width=158, height=40, fg_color="#0a5abf", hover_color="#1269d3",
            corner_radius=10, border_width=1, border_color="#1a7fff")
        self._scan_btn.pack(side="left", padx=(0, 8))
        ctk.CTkButton(inner, text="⚙  Settings", font=FB(13),
                      command=self._open_settings,
                      width=118, height=40, fg_color=C_CARD2, hover_color=C_BORDER,
                      corner_radius=10, border_width=1, border_color=C_BORDER).pack(side="left")

        # Glowing accent line at bottom of header
        ctk.CTkFrame(self, height=2, corner_radius=0, fg_color=C_ACCENT).grid(
            row=0, column=0, columnspan=2, sticky="sew")

    # ── SIDEBAR ────────────────────────────────────────────────────────────────
    def _build_sidebar(self):
        sb = ctk.CTkFrame(self, width=230, corner_radius=0, fg_color=C_SIDEBAR)
        sb.grid(row=1, column=0, sticky="nsew")
        sb.grid_propagate(False)
        sb.grid_columnconfigure(0, weight=1)
        sb.grid_rowconfigure(99, weight=1)

        # Separator under header
        ctk.CTkFrame(sb, height=1, fg_color=C_BORDER, corner_radius=0).grid(
            row=0, column=0, sticky="ew")

        # Section label
        ctk.CTkLabel(sb, text="  NAVIGATION",
                     font=ctk.CTkFont(size=10, weight="bold"), text_color="gray35",
                     anchor="w").grid(row=1, column=0, padx=8, pady=(16, 8), sticky="ew")

        nav_items = [
            ("dashboard",   "🏠", "Dashboard",    "All system stats at a glance"),
            ("hardware",    "💻", "Hardware",     "CPU · GPU · RAM · Storage"),
            ("network",     "🌐", "Network",      "Adapters & DNS changer"),
            ("processes",   "📋", "Processes",    "Top CPU & RAM consumers"),
            ("maintenance", "🔧", "Maintenance",  "Updates · Diagnosis · Repair"),
            ("optimise",    "⚡", "Optimise",     "Gaming tweaks & debloat"),
        ]

        for i, (nid, icon, label, hint) in enumerate(nav_items):
            outer = ctk.CTkFrame(sb, fg_color="transparent", corner_radius=10)
            outer.grid(row=i + 2, column=0, padx=8, pady=3, sticky="ew")
            outer.grid_columnconfigure(0, weight=1)

            btn = ctk.CTkButton(
                outer,
                text=f"  {icon}   {label}",
                font=ctk.CTkFont(size=15, weight="bold"),
                anchor="w",
                height=50,
                corner_radius=10,
                fg_color="transparent",
                hover_color=C_NAVSEL,
                text_color="gray65",
                border_width=0,
                command=lambda n=nid: self._nav_select(n)
            )
            btn.grid(row=0, column=0, sticky="ew")
            self._nav_buttons[nid] = btn

            # Hint label
            ctk.CTkLabel(outer, text=f"    {hint}",
                         font=ctk.CTkFont(size=10), text_color="gray32",
                         anchor="w").grid(row=1, column=0, padx=4, pady=(0, 3), sticky="w")

        # Spacer row
        ctk.CTkFrame(sb, fg_color="transparent").grid(row=99, column=0, sticky="nsew")

        # Separator
        ctk.CTkFrame(sb, height=1, fg_color=C_BORDER).grid(
            row=100, column=0, sticky="ew", padx=10, pady=(0, 8))

        # Admin status badge
        if not self._is_admin:
            badge = ctk.CTkFrame(sb, fg_color="#120800", corner_radius=8,
                                 border_color="#3a1a00", border_width=1)
            badge.grid(row=101, column=0, padx=10, pady=(0, 8), sticky="ew")
            ctk.CTkLabel(badge, text="⚠  Not Admin",
                         font=ctk.CTkFont(size=12, weight="bold"),
                         text_color=C_ORANGE).pack(padx=10, pady=(6, 2), anchor="w")
            ctk.CTkLabel(badge, text="Some features need admin.\nRight-click → Run as admin.",
                         font=ctk.CTkFont(size=10), text_color="#804000",
                         wraplength=190, justify="left").pack(padx=10, pady=(0, 4), anchor="w")
            ctk.CTkButton(badge, text="🔑 Relaunch as Admin",
                          font=ctk.CTkFont(size=11, weight="bold"), height=30, corner_radius=6,
                          fg_color="#7a3500", hover_color="#9a4500",
                          command=self._relaunch_admin
                          ).pack(padx=8, pady=(2, 8), fill="x")
        else:
            ctk.CTkLabel(sb, text="✓  Running as Administrator",
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=C_GREEN).grid(row=101, column=0, padx=14, pady=(0, 12), sticky="w")

        self._sidebar = sb

    def _nav_select(self, nav_id: str):
        """Switch active navigation section."""
        # Deselect previous
        if self._active_nav and self._active_nav in self._nav_buttons:
            self._nav_buttons[self._active_nav].configure(
                fg_color="transparent", text_color="gray65")
        # Highlight new
        self._active_nav = nav_id
        if nav_id in self._nav_buttons:
            self._nav_buttons[nav_id].configure(fg_color=C_NAVSEL, text_color=C_ACCENT)
        # Show/hide content frames
        for fid, frame in self._content_frames.items():
            if fid == nav_id:
                frame.grid()
            else:
                frame.grid_remove()

    def _relaunch_admin(self):
        import ctypes
        try:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable,
                f'"{os.path.abspath(__file__)}"', None, 1)
            self.quit()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ── CONTENT AREA ───────────────────────────────────────────────────────────
    def _build_content(self):
        container = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        container.grid(row=1, column=1, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)
        self._container = container

        self._build_dashboard_section()
        self._build_hardware_section()
        self._build_network_section()
        self._build_process_section()
        self._build_maintenance_section()
        self._build_optimise_section()

        for f in self._content_frames.values():
            f.grid_remove()

    def _make_section(self, nid: str) -> ctk.CTkFrame:
        f = ctk.CTkFrame(self._container, fg_color="transparent", corner_radius=0)
        f.grid(row=0, column=0, sticky="nsew")
        f.grid_columnconfigure(0, weight=1)
        # NOTE: do NOT set row weights here — each section builder sets its own
        self._content_frames[nid] = f
        return f

    # ── FOOTER ─────────────────────────────────────────────────────────────────
    def _build_footer(self):
        foot = ctk.CTkFrame(self, height=28, corner_radius=0, fg_color=C_HDR)
        foot.grid(row=2, column=0, columnspan=2, sticky="ew")
        foot.grid_propagate(False)
        foot.grid_columnconfigure(1, weight=1)
        ctk.CTkFrame(foot, height=1, corner_radius=0, fg_color=C_BORDER).grid(
            row=0, column=0, columnspan=3, sticky="new")
        self._ov_refresh_ts = ctk.CTkLabel(foot, text="", font=FM(10), text_color="gray35")
        self._ov_refresh_ts.grid(row=0, column=0, padx=14, pady=4, sticky="w")
        ctk.CTkLabel(foot, text="Created By: CRiMiNAL",
                     font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=C_PURPLE).grid(row=0, column=2, padx=16, pady=4, sticky="e")

    # ══════════════════════════════════════════════════════════════════════════
    # SHARED WIDGET HELPERS
    # ══════════════════════════════════════════════════════════════════════════
    def _section_lbl(self, parent, row, text, color=C_ACCENT):
        hf = ctk.CTkFrame(parent, fg_color=C_CARD2, corner_radius=6)
        hf.grid(row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=(10, 0))
        hf.grid_columnconfigure(1, weight=1)
        ctk.CTkFrame(hf, width=3, corner_radius=1, fg_color=color).grid(
            row=0, column=0, sticky="ns", padx=(8, 6), pady=5)
        ctk.CTkLabel(hf, text=text, font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=color, anchor="w").grid(
            row=0, column=1, padx=(0, 10), pady=5, sticky="w")
        ctk.CTkFrame(parent, height=1, fg_color="transparent").grid(
            row=row + 1, column=0, columnspan=2, sticky="ew", padx=6)

    def _kv(self, parent, row, key, val, kcolor=C_ACCENT, vcolor="gray85"):
        ctk.CTkLabel(parent, text=f"{key}:", font=FL(13), text_color=kcolor, anchor="w").grid(
            row=row, column=0, padx=(18, 6), pady=3, sticky="w")
        ctk.CTkLabel(parent, text=str(val), font=FV(13), text_color=vcolor, anchor="w",
                     wraplength=500).grid(row=row, column=1, padx=(4, 14), pady=3, sticky="w")

    def _pbar(self, parent, row, label, pct, warn_at=70, crit_at=88):
        color = C_GREEN if pct < warn_at else C_ORANGE if pct < crit_at else C_RED
        ctk.CTkLabel(parent, text=f"{label}:", font=FL(13), text_color=C_ACCENT, anchor="w").grid(
            row=row, column=0, padx=(18, 6), pady=4, sticky="w")
        cont = ctk.CTkFrame(parent, fg_color="transparent")
        cont.grid(row=row, column=1, padx=(4, 14), pady=4, sticky="ew")
        cont.grid_columnconfigure(0, weight=1)
        bar = ctk.CTkProgressBar(cont, progress_color=color, height=16, corner_radius=8,
                                  fg_color="#0a1020")
        bar.grid(row=0, column=0, sticky="ew")
        bar.set(min(max(float(pct) / 100.0, 0.0), 1.0))
        ctk.CTkLabel(cont, text=f"{pct}%", font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=color, width=44).grid(row=0, column=1, padx=(8, 0))

    def _clear(self, f):
        for w in f.winfo_children(): w.destroy()

    def _set_status(self, text):
        self.after(0, lambda: self._status_lbl.configure(text=text))

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 1: DASHBOARD (was Overview)
    # ══════════════════════════════════════════════════════════════════════════
    def _build_dashboard_section(self):
        section = self._make_section("dashboard")
        section.grid_columnconfigure((0, 1, 2), weight=1)
        section.grid_rowconfigure(0, weight=0)   # header row — fixed height
        section.grid_rowconfigure(1, weight=1)   # scroll area — takes all space
        section.grid_rowconfigure(2, weight=0)   # DNS quick strip — fixed

        # Header row
        hdr = ctk.CTkFrame(section, fg_color="transparent")
        hdr.grid(row=0, column=0, columnspan=3, sticky="ew", padx=12, pady=(8, 4))
        hdr.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(hdr, text="System Dashboard",
                     font=ctk.CTkFont(size=19, weight="bold"), text_color="white").grid(
            row=0, column=0, padx=4)
        ctk.CTkLabel(hdr, text="Live system overview · click any card to see details",
                     font=ctk.CTkFont(size=11), text_color="gray38").grid(
            row=0, column=1, padx=10, sticky="w")
        ctk.CTkButton(hdr, text="🔄  Refresh", font=FB(12), width=128, height=34,
                      corner_radius=8, fg_color=C_CARD2, hover_color=C_BORDER,
                      command=lambda: threading.Thread(
                          target=self._manual_refresh_overview, daemon=True).start()
                      ).grid(row=0, column=2, padx=4)

        # Scrollable grid of cards
        scroll = ctk.CTkScrollableFrame(section, fg_color="transparent", corner_radius=0)
        scroll.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=4)
        scroll.grid_columnconfigure((0, 1, 2), weight=1)
        self._ov_scroll = scroll

        self._ov_cards   = {}
        self._ov_labels  = {}
        self._ov_bars    = {}
        self._ov_bar_lbl = {}
        self._ov_usage_bar = {}
        self._ov_usage_lbl = {}

        specs = [
            ("🔲  CPU",         "cpu",     C_ACCENT),
            ("🎮  GPU",         "gpu",     C_PURPLE),
            ("💾  RAM",         "ram",     C_TEAL),
            ("💿  Storage",     "storage", "#26c6a8"),
            ("🖥  Motherboard", "mb",      "#78909c"),
            ("🌐  Network",     "net",     C_GREEN),
            ("🪟  OS",          "os",      "#6c8ef7"),
            ("📈  Performance", "perf",    C_ORANGE),
            ("🔋  Power & BIOS","power",   C_PINK),
        ]

        for i, (title, key, accent) in enumerate(specs):
            card = ctk.CTkFrame(scroll, corner_radius=14,
                                fg_color=C_CARD, border_color=accent, border_width=1)
            card.grid(row=i // 3, column=i % 3, padx=7, pady=7, sticky="nsew")
            card.grid_columnconfigure(0, weight=1)

            # Accent top strip
            ctk.CTkFrame(card, height=5, corner_radius=0, fg_color=accent).grid(
                row=0, column=0, sticky="ew")
            # Card title bar
            h2 = ctk.CTkFrame(card, fg_color=C_CARD2, corner_radius=0)
            h2.grid(row=1, column=0, sticky="ew")
            h2.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(h2, text=title,
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=accent, anchor="w").grid(
                row=0, column=0, padx=14, pady=8, sticky="w")
            ctk.CTkFrame(card, height=1, fg_color=C_BORDER).grid(
                row=2, column=0, sticky="ew", padx=10)

            lbl = ctk.CTkLabel(card, text="Scanning…",
                               font=ctk.CTkFont(family="Consolas", size=12, weight="bold"),
                               text_color="gray45", justify="left", wraplength=300, anchor="nw")
            lbl.grid(row=3, column=0, padx=14, pady=(10, 4), sticky="w")

            bar_row = ctk.CTkFrame(card, fg_color="transparent")
            bar_row.grid(row=4, column=0, sticky="ew", padx=14, pady=(0, 12))
            bar_row.grid_columnconfigure(0, weight=1)
            bar = ctk.CTkProgressBar(bar_row, height=10, corner_radius=5,
                                      progress_color=accent, fg_color="#0a1020")
            bar.grid(row=0, column=0, sticky="ew")
            bar.set(0)
            bar_lbl = ctk.CTkLabel(bar_row, text="",
                                    font=ctk.CTkFont(size=11, weight="bold"),
                                    text_color=accent, width=44, anchor="e")
            bar_lbl.grid(row=0, column=1, padx=(8, 0))

            self._ov_cards[key]   = card
            self._ov_labels[key]  = lbl
            self._ov_bars[key]    = bar
            self._ov_bar_lbl[key] = bar_lbl

        # Storage: inner scrollable list
        stor_card = self._ov_cards["storage"]
        self._ov_labels["storage"].grid_remove()
        self._ov_bars["storage"].master.grid_configure(row=5, pady=(0, 10))
        stor_card.grid_rowconfigure(3, weight=1)
        self._ov_stor_inner = ctk.CTkScrollableFrame(
            stor_card, fg_color="transparent", height=190, corner_radius=0)
        self._ov_stor_inner.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=6, pady=(4, 0))
        self._ov_stor_inner.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self._ov_stor_inner, text="Scanning…",
                     font=FM(11), text_color="gray45").grid(row=0, column=0, pady=20)

        # CPU & GPU already get the bold ubadge below — remove their thin bar_row
        for _k in ("cpu", "gpu"):
            self._ov_bars[_k].master.grid_remove()

        # CPU & GPU live usage badges
        for key in ("cpu", "gpu"):
            accent   = C_ACCENT if key == "cpu" else C_PURPLE
            icon_txt = "CPU %" if key == "cpu" else "GPU %"
            c_card   = self._ov_cards[key]
            ubadge = ctk.CTkFrame(c_card, fg_color=C_CARD2, corner_radius=8,
                                  border_color=C_BORDER, border_width=1)
            ubadge.grid(row=5, column=0, sticky="ew", padx=10, pady=(2, 10))
            ubadge.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(ubadge, text=icon_txt,
                         font=ctk.CTkFont(family="Consolas", size=11, weight="bold"),
                         text_color=accent, width=44, anchor="w").grid(
                row=0, column=0, padx=(10, 6), pady=8)
            ubar = ctk.CTkProgressBar(ubadge, height=14, corner_radius=7,
                                       progress_color=accent, fg_color="#0a1020")
            ubar.grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=8)
            ubar.set(0)
            ulbl = ctk.CTkLabel(ubadge, text="--",
                                font=ctk.CTkFont(size=13, weight="bold"),
                                text_color=accent, width=50, anchor="e")
            ulbl.grid(row=0, column=2, padx=(2, 10), pady=8)
            self._ov_usage_bar[key] = ubar
            self._ov_usage_lbl[key] = ulbl

        # ── DNS Quick Strip (row 2 of dashboard) ─────────────────────────────
        dns_strip = ctk.CTkFrame(section, fg_color=C_CARD, corner_radius=10,
                                 border_color=C_GREEN, border_width=1)
        dns_strip.grid(row=2, column=0, columnspan=3, sticky="ew", padx=8, pady=(2, 8))
        dns_strip.grid_columnconfigure(3, weight=1)

        ctk.CTkFrame(dns_strip, width=4, corner_radius=2, fg_color=C_GREEN).grid(
            row=0, column=0, sticky="ns", padx=(10, 8), pady=10)
        ctk.CTkLabel(dns_strip, text="🌐  Quick DNS",
                     font=ctk.CTkFont(size=13, weight="bold"), text_color=C_GREEN).grid(
            row=0, column=1, padx=(0, 12), pady=10)

        # Use same _dns_var — lazily init in case Network section hasn't built yet
        if not hasattr(self, "_dns_var"):
            self._dns_var = tk.StringVar(value=list(self._DNS_PRESETS.keys())[0])
        ctk.CTkOptionMenu(dns_strip, variable=self._dns_var,
                          values=list(self._DNS_PRESETS.keys()),
                          font=FM(12), fg_color=C_CARD2, button_color=C_BORDER,
                          button_hover_color="#2a3560", dropdown_fg_color=C_CARD2,
                          width=260).grid(row=0, column=2, padx=(0, 10), pady=10)

        # Dashboard-specific status label (DNS actions update both status labels)
        self._dns_status_lbl_dash = ctk.CTkLabel(dns_strip, text="",
                                                   font=ctk.CTkFont(size=11, weight="bold"),
                                                   text_color="gray50", anchor="w")
        self._dns_status_lbl_dash.grid(row=0, column=3, padx=(0, 8), pady=10, sticky="w")

        # Buttons sub-frame using pack for clean horizontal layout
        dns_btns = ctk.CTkFrame(dns_strip, fg_color="transparent")
        dns_btns.grid(row=0, column=4, padx=(0, 10), pady=10)
        for text, fg, hv, cmd in [
            ("✅ Apply",  "#0b5c1e", "#10762a",
             lambda: threading.Thread(target=self._apply_dns, daemon=True).start()),
            ("↩ Reset",  "#2a3550", "#3a4a70",
             lambda: threading.Thread(target=self._reset_dns, daemon=True).start()),
            ("📡 Ping",  "#0b3d91", "#1153b5",
             lambda: threading.Thread(target=self._ping_dns, daemon=True).start()),
        ]:
            ctk.CTkButton(dns_btns, text=text, font=FB(12), height=36, width=104,
                          fg_color=fg, hover_color=hv, corner_radius=8,
                          command=cmd).pack(side="left", padx=(0, 6))

    def _ov_update_usage_badge(self, key: str, pct):
        if key not in self._ov_usage_bar: return
        if pct is None:
            self._ov_usage_bar[key].set(0)
            self._ov_usage_lbl[key].configure(text="--", text_color="gray50")
            return
        accent = C_ACCENT if key == "cpu" else C_PURPLE
        color  = C_GREEN if pct < 50 else accent if pct < 75 else C_ORANGE if pct < 90 else C_RED
        self._ov_usage_bar[key].configure(progress_color=color)
        self._ov_usage_bar[key].set(min(float(pct) / 100.0, 1.0))
        self._ov_usage_lbl[key].configure(text=f"{pct:.0f}%", text_color=color)

    def _ov_update_storage(self, stor):
        for w in self._ov_stor_inner.winfo_children(): w.destroy()
        all_used_pcts = []; r = 0
        for disk in stor:
            if "error" in disk: continue
            parts = disk.get("partitions", []); size = disk.get("size_gb", "?")
            if not parts and (size == 0 or size == "N/A"): continue
            hf = ctk.CTkFrame(self._ov_stor_inner, fg_color=C_CARD2, corner_radius=8)
            hf.grid(row=r, column=0, sticky="ew", padx=4, pady=(6, 2))
            hf.grid_columnconfigure(0, weight=1); r += 1
            model_short = (disk.get("model", "?") or "?")[:38]
            ctk.CTkLabel(hf, text=f"💿  {model_short}",
                         font=ctk.CTkFont(size=12, weight="bold"),
                         text_color="#26a69a", anchor="w").grid(row=0, column=0, padx=10, pady=(6, 0), sticky="w")
            ctk.CTkLabel(hf, text=f"{size} GB   ·   {disk.get('interface', '?')}",
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color="gray45", anchor="w").grid(row=1, column=0, padx=10, pady=(0, 6), sticky="w")
            for p in parts:
                used_pct = float(p.get("percent", 0)); free_pct = max(0.0, 100.0 - used_pct)
                free_gb  = p.get("free_gb", "?"); total_gb = p.get("total_gb", "?")
                mnt = (p.get("mountpoint", "?") or "?").rstrip("\\")
                all_used_pcts.append(used_pct)
                bar_color = C_GREEN if used_pct < 70 else C_ORANGE if used_pct < 88 else C_RED
                pf = ctk.CTkFrame(self._ov_stor_inner, fg_color="#0e1222",
                                  corner_radius=7, border_color=C_BORDER, border_width=1)
                pf.grid(row=r, column=0, sticky="ew", padx=4, pady=2)
                pf.grid_columnconfigure(0, weight=1); r += 1
                top = ctk.CTkFrame(pf, fg_color="transparent")
                top.grid(row=0, column=0, sticky="ew", padx=10, pady=(7, 2))
                top.grid_columnconfigure(1, weight=1)
                ctk.CTkLabel(top, text=mnt,
                             font=ctk.CTkFont(family="Consolas", size=13, weight="bold"),
                             text_color="gray85", width=36, anchor="w").grid(row=0, column=0, padx=(0, 8))
                bar = ctk.CTkProgressBar(top, height=12, corner_radius=6,
                                          progress_color=bar_color, fg_color="#1a1f38")
                bar.grid(row=0, column=1, sticky="ew")
                bar.set(min(free_pct / 100.0, 1.0))
                ctk.CTkLabel(top, text=f"{int(used_pct)}% used",
                             font=ctk.CTkFont(size=11, weight="bold"),
                             text_color=bar_color, width=72, anchor="e").grid(row=0, column=2, padx=(8, 0))
                ctk.CTkLabel(pf, text=f"Free:  {free_gb} GB  /  {total_gb} GB",
                             font=ctk.CTkFont(size=11, weight="bold"),
                             text_color="gray40", anchor="w").grid(row=1, column=0, padx=10, pady=(0, 7), sticky="w")
        if not all_used_pcts:
            ctk.CTkLabel(self._ov_stor_inner, text="No drives detected",
                         font=FM(11), text_color="gray45").grid(row=0, column=0, pady=20)
            return
        mp = max(all_used_pcts); bc = C_GREEN if mp < 70 else C_ORANGE if mp < 88 else C_RED
        self._ov_bars["storage"].configure(progress_color=bc)
        self._ov_bars["storage"].set(min(mp / 100.0, 1.0))
        self._ov_bar_lbl["storage"].configure(text=f"{int(mp)}%", text_color=bc)

    def _ov_set(self, key, text, color="gray80", pct=None, bar_color=None):
        if key in self._ov_labels:
            self._ov_labels[key].configure(text=text, text_color=color)
        if pct is not None and key in self._ov_bars:
            bc = bar_color or (C_GREEN if pct < 70 else C_ORANGE if pct < 88 else C_RED)
            self._ov_bars[key].configure(progress_color=bc)
            self._ov_bars[key].set(min(max(float(pct) / 100.0, 0.0), 1.0))
            self._ov_bar_lbl[key].configure(text=f"{pct}%", text_color=bc)
        elif key in self._ov_bars and pct is None:
            self._ov_bars[key].set(0); self._ov_bar_lbl[key].configure(text="")

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 2: HARDWARE (CPU · GPU · RAM · Storage sub-tabs)
    # ══════════════════════════════════════════════════════════════════════════
    def _build_hardware_section(self):
        section = self._make_section("hardware")
        section.grid_columnconfigure(0, weight=1)
        section.grid_rowconfigure(0, weight=1)   # tabs take all space

        hw_tabs = ctk.CTkTabview(
            section, corner_radius=12, fg_color=C_PANEL,
            segmented_button_fg_color=C_CARD,
            segmented_button_selected_color="#0a4a9f",
            segmented_button_selected_hover_color="#0e5bbf",
            segmented_button_unselected_color=C_CARD,
            segmented_button_unselected_hover_color=C_CARD2)
        hw_tabs.grid(row=0, column=0, sticky="nsew", padx=10, pady=(8, 10))

        for name in ("🔲 CPU", "🎮 GPU", "💾 RAM", "💿 Storage"):
            hw_tabs.add(name)

        self._hw_tabs    = hw_tabs
        self._cpu_frame  = self._make_hw_scroll("🔲 CPU",     "CPU Details")
        self._gpu_frame  = self._make_hw_scroll("🎮 GPU",     "GPU Details")
        self._ram_frame  = self._make_hw_scroll("💾 RAM",     "RAM Details")
        self._stor_frame = self._make_hw_scroll("💿 Storage", "Storage Details")

    def _make_hw_scroll(self, tab_name, title):
        tab = self._hw_tabs.tab(tab_name)
        tab.configure(fg_color="transparent")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=0)   # title bar — fixed
        tab.grid_rowconfigure(1, weight=1)   # scroll area — fills remaining space

        # Compact section title bar
        title_bar = ctk.CTkFrame(tab, fg_color=C_CARD2, corner_radius=8, height=36)
        title_bar.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 4))
        title_bar.grid_propagate(False)
        title_bar.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(title_bar, text=title,
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=C_ACCENT, anchor="w").grid(
            row=0, column=0, padx=14, sticky="w")

        frame = ctk.CTkScrollableFrame(tab, fg_color="transparent",
                                        corner_radius=0, height=1)
        frame.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 4))
        frame.grid_columnconfigure(0, weight=0)
        frame.grid_columnconfigure(1, weight=1)
        return frame

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 3: NETWORK (adapters + DNS changer)
    # ══════════════════════════════════════════════════════════════════════════
    _DNS_PRESETS = {
        "Cloudflare  (1.1.1.1)":        ("1.1.1.1",        "1.0.0.1"),
        "Google  (8.8.8.8)":             ("8.8.8.8",        "8.8.4.4"),
        "Quad9  (9.9.9.9)":              ("9.9.9.9",        "149.112.112.112"),
        "OpenDNS  (208.67.222.222)":     ("208.67.222.222", "208.67.220.220"),
        "Cloudflare Gaming  (1.1.1.2)":  ("1.1.1.2",        "1.0.0.2"),
    }

    def _build_network_section(self):
        section = self._make_section("network")
        section.grid_columnconfigure(0, weight=1)
        section.grid_rowconfigure(0, weight=0)   # section title bar — fixed
        section.grid_rowconfigure(1, weight=1)   # adapter list — expands
        section.grid_rowconfigure(2, weight=0)   # dns card — fixed height

        # ── Section title bar (outside the scrollframe) ───────────────────────
        net_hdr = ctk.CTkFrame(section, fg_color=C_CARD2, corner_radius=8, height=36)
        net_hdr.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        net_hdr.grid_propagate(False)
        net_hdr.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(net_hdr, text="🌐  Network Adapters",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=C_ACCENT, anchor="w").grid(
            row=0, column=0, padx=14, sticky="w")

        # ── Adapter list (pure data, no title inside) ─────────────────────────
        self._net_frame = ctk.CTkScrollableFrame(section, fg_color="transparent",
                                                   corner_radius=0, height=1)
        self._net_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 4))
        self._net_frame.grid_columnconfigure(0, weight=0)
        self._net_frame.grid_columnconfigure(1, weight=1)

        # ── DNS Changer card ──────────────────────────────────────────────────
        dns_card = ctk.CTkFrame(section, fg_color=C_CARD, corner_radius=12,
                                border_color=C_GREEN, border_width=1)
        dns_card.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        dns_card.grid_columnconfigure(2, weight=1)

        ctk.CTkFrame(dns_card, height=4, corner_radius=0, fg_color=C_GREEN).grid(
            row=0, column=0, columnspan=5, sticky="ew")
        hdr_f = ctk.CTkFrame(dns_card, fg_color=C_CARD2, corner_radius=0)
        hdr_f.grid(row=1, column=0, columnspan=5, sticky="ew")
        hdr_f.grid_columnconfigure(1, weight=1)
        ctk.CTkFrame(hdr_f, width=4, corner_radius=2, fg_color=C_GREEN).grid(
            row=0, column=0, sticky="ns", padx=(10, 8), pady=8)
        ctk.CTkLabel(hdr_f, text="🌐  DNS Changer",
                     font=ctk.CTkFont(size=14, weight="bold"), text_color=C_GREEN).grid(
            row=0, column=1, padx=0, pady=8, sticky="w")

        if not hasattr(self, "_dns_var"):
            self._dns_var = tk.StringVar(value=list(self._DNS_PRESETS.keys())[0])
        ctk.CTkLabel(dns_card, text="DNS Server:", font=FL(13), text_color=C_ACCENT, anchor="w").grid(
            row=2, column=0, padx=(14, 6), pady=(10, 10), sticky="w")
        ctk.CTkOptionMenu(dns_card, variable=self._dns_var, values=list(self._DNS_PRESETS.keys()),
                          font=FM(12), fg_color=C_CARD2, button_color=C_BORDER,
                          button_hover_color="#2a3560", dropdown_fg_color=C_CARD2,
                          width=300).grid(row=2, column=1, padx=(0, 10), pady=(10, 10), sticky="w")
        self._dns_status_lbl = ctk.CTkLabel(dns_card, text="", font=FM(11),
                                             text_color="gray50", anchor="w")
        self._dns_status_lbl.grid(row=2, column=2, padx=(0, 14), pady=(10, 10), sticky="w")

        bf = ctk.CTkFrame(dns_card, fg_color="transparent")
        bf.grid(row=3, column=0, columnspan=5, sticky="w", padx=14, pady=(0, 12))
        for text, fg, hv, cmd in [
            ("✅  Apply DNS",    "#0b5c1e", "#10762a",
             lambda: threading.Thread(target=self._apply_dns, daemon=True).start()),
            ("↩  Reset to Auto", "#2a3550", "#3a4a70",
             lambda: threading.Thread(target=self._reset_dns, daemon=True).start()),
            ("📡  Ping Test",    "#0b3d91", "#1153b5",
             lambda: threading.Thread(target=self._ping_dns, daemon=True).start()),
        ]:
            ctk.CTkButton(bf, text=text, font=FB(12), height=40, width=162,
                          fg_color=fg, hover_color=hv, corner_radius=10, command=cmd).pack(
                side="left", padx=(0, 8))

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 4: PROCESSES
    # ══════════════════════════════════════════════════════════════════════════
    def _build_process_section(self):
        section = self._make_section("processes")
        section.grid_columnconfigure((0, 1), weight=1)
        section.grid_rowconfigure(0, weight=0)   # header fixed
        section.grid_rowconfigure(1, weight=1)   # tables expand

        hdr = ctk.CTkFrame(section, fg_color="transparent")
        hdr.grid(row=0, column=0, columnspan=2, sticky="ew", padx=12, pady=(10, 4))
        hdr.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(hdr, text="Process Monitor",
                     font=ctk.CTkFont(size=19, weight="bold"), text_color="white").grid(
            row=0, column=0, padx=4)
        self._proc_ts_lbl = ctk.CTkLabel(hdr, text="", font=FM(11), text_color="gray45")
        self._proc_ts_lbl.grid(row=0, column=1, sticky="w", padx=10)
        ctk.CTkButton(hdr, text="🔄  Refresh Now", font=FB(12), width=140, height=36,
                      fg_color=C_CARD2, hover_color=C_BORDER, corner_radius=10,
                      command=lambda: threading.Thread(
                          target=self._refresh_processes, daemon=True).start()
                      ).grid(row=0, column=2, padx=4)

        for col, (title, tcolor, attr) in enumerate([
            ("🔲  Top CPU Consumers", C_ACCENT, "_proc_cpu_frame"),
            ("💾  Top RAM Consumers", "#00bcd4", "_proc_ram_frame"),
        ]):
            panel = ctk.CTkFrame(section, fg_color=C_CARD, corner_radius=10,
                                  border_color=C_BORDER, border_width=1)
            panel.grid(row=1, column=col, sticky="nsew",
                       padx=(10 if col == 0 else 4, 4 if col == 0 else 10), pady=(0, 10))
            panel.grid_columnconfigure(0, weight=1)
            panel.grid_rowconfigure(0, weight=0)   # title — fixed
            panel.grid_rowconfigure(1, weight=1)   # scrollframe — fills panel
            panel.grid_propagate(True)

            # Panel title bar
            ptitle = ctk.CTkFrame(panel, fg_color=C_CARD2, corner_radius=8, height=36)
            ptitle.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 4))
            ptitle.grid_propagate(False)
            ptitle.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(ptitle, text=title, font=FH(13), text_color=tcolor, anchor="w").grid(
                row=0, column=0, padx=12, sticky="w")

            # height=1 surrenders size control to grid manager → fills panel properly
            sf = ctk.CTkScrollableFrame(panel, fg_color="transparent",
                                         corner_radius=0, height=1)
            sf.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))
            sf.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)
            setattr(self, attr, sf)

    def _render_proc_list(self, frame, procs, sort_key):
        for w in frame.winfo_children(): w.destroy()
        hdr_cols = [("PID", 48), ("Name", 0), ("CPU%", 56), ("RAM MB", 68), ("Status", 68)]
        hf = ctk.CTkFrame(frame, fg_color=C_CARD2, corner_radius=6)
        hf.grid(row=0, column=0, columnspan=5, sticky="ew", padx=2, pady=(0, 2))
        for c, (t, w) in enumerate(hdr_cols):
            ctk.CTkLabel(hf, text=t, font=FM(11), text_color=C_ACCENT,
                         anchor="w", width=w if w else 160).grid(
                row=0, column=c, padx=6, pady=4, sticky="w")
        for i, p in enumerate(procs):
            bg = C_CARD if i % 2 == 0 else C_CARD2
            row_f = ctk.CTkFrame(frame, fg_color=bg, corner_radius=4)
            row_f.grid(row=i + 1, column=0, columnspan=5, sticky="ew", padx=2, pady=1)
            cpu_val = p.get("cpu", 0); ram_val = p.get("ram_mb", 0)
            cpu_col = C_GREEN if cpu_val < 40 else C_ORANGE if cpu_val < 70 else C_RED
            ram_col = C_GREEN if ram_val < 1000 else C_ORANGE if ram_val < 3000 else C_RED
            data = [
                (str(p.get("pid", "?")), "gray55", 48),
                (p.get("name", "?")[:24], "gray85", 0),
                (f"{cpu_val:.1f}%", cpu_col, 56),
                (f"{ram_val:.0f}", ram_col, 68),
                (p.get("status", "?"), "gray55", 68),
            ]
            for c, (val, color, width) in enumerate(data):
                ctk.CTkLabel(row_f, text=val, font=FM(11), text_color=color,
                             anchor="w", width=width if width else 160).grid(
                    row=0, column=c, padx=6, pady=3, sticky="w")

    def _refresh_processes(self):
        try:
            import os as _os
            data = self.sys_info.get_top_processes(20)   # fetch more so we have enough after filtering
            own_pid = _os.getpid()
            _SKIP_NAMES = {"system idle process", "system idle", "idle"}
            def _clean(lst):
                out = []
                for p in lst:
                    if p.get("pid") == 0: continue          # System Idle Process
                    if p.get("pid") == own_pid: continue    # this app (python.exe)
                    if (p.get("name") or "").lower() in _SKIP_NAMES: continue
                    out.append(p)
                    if len(out) >= 12: break
                return out
            cleaned = {"by_cpu": _clean(data["by_cpu"]), "by_ram": _clean(data["by_ram"])}
            self._proc_data = cleaned
            import datetime as _dt; ts = _dt.datetime.now().strftime("%H:%M:%S")
            self.after(0, lambda d=cleaned, t=ts: (
                self._render_proc_list(self._proc_cpu_frame, d["by_cpu"], "cpu"),
                self._render_proc_list(self._proc_ram_frame, d["by_ram"], "ram_mb"),
                self._proc_ts_lbl.configure(text=f"Updated: {t}", text_color="gray45")))
        except Exception:
            pass

    def _auto_refresh_processes(self):
        threading.Thread(target=self._refresh_processes, daemon=True).start()
        self.after(3000, self._auto_refresh_processes)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 5: MAINTENANCE (Updates · Diagnosis · System Repair)
    # ══════════════════════════════════════════════════════════════════════════
    def _build_maintenance_section(self):
        section = self._make_section("maintenance")
        section.grid_columnconfigure(0, weight=1)
        section.grid_rowconfigure(0, weight=1)   # tabs take all space

        m_tabs = ctk.CTkTabview(
            section, corner_radius=12, fg_color=C_PANEL,
            segmented_button_fg_color=C_CARD,
            segmented_button_selected_color="#0a4a9f",
            segmented_button_selected_hover_color="#0e5bbf",
            segmented_button_unselected_color=C_CARD,
            segmented_button_unselected_hover_color=C_CARD2)
        m_tabs.grid(row=0, column=0, sticky="nsew", padx=10, pady=(8, 10))
        for name in ("🔄 Updates", "🔬 Diagnosis", "🛡 System Repair"):
            m_tabs.add(name)
        self._maint_tabs = m_tabs

        self._build_updates_tab()
        self._build_diagnosis_tab()
        self._build_repair_tab()

    def _switch_to_updates(self):
        """Switch to maintenance > updates (used from diagnosis action)."""
        self._nav_select("maintenance")
        self.after(100, lambda: self._maint_tabs.set("🔄 Updates"))

    # ── UPDATES TAB ──────────────────────────────────────────────────────────
    def _build_updates_tab(self):
        tab = self._maint_tabs.tab("🔄 Updates")
        tab.configure(fg_color="transparent")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(tab, text="Updates & Drivers", font=FT(17)).grid(
            row=0, column=0, pady=(6, 6))

        ar = ctk.CTkFrame(tab, fg_color=C_CARD, corner_radius=10,
                          border_color=C_BORDER, border_width=1)
        ar.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 6))
        ar.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)
        for col, (text, cmd, fg, hv) in enumerate([
            ("🔍  Scan Updates",    self._check_all_updates,   "#0d47a1", "#1565c0"),
            ("📦  Update Selected", self._install_selected,    "#1b5e20", "#2e7d32"),
            ("📦  Update All",      self._install_all_updates, "#33691e", "#558b2f"),
            ("🎮  AMD Driver",      self._update_amd,          "#5b11a0", "#7b1fa2"),
            ("🪟  Windows Update",  self._open_win_upd,        "#0d47a1", "#1565c0"),
        ]):
            ctk.CTkButton(ar, text=text, font=FB(12), command=cmd,
                          height=46, fg_color=fg, hover_color=hv, corner_radius=10).grid(
                row=0, column=col, padx=5, pady=10, sticky="ew")

        body = ctk.CTkFrame(tab, fg_color="transparent")
        body.grid(row=2, column=0, sticky="nsew", padx=6)
        body.grid_columnconfigure(0, weight=3); body.grid_columnconfigure(1, weight=2)
        body.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(body, fg_color=C_CARD, corner_radius=10,
                            border_color=C_BORDER, border_width=1)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        left.grid_columnconfigure(0, weight=1); left.grid_rowconfigure(1, weight=1)
        lhdr = ctk.CTkFrame(left, fg_color=C_CARD2, corner_radius=8)
        lhdr.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        lhdr.grid_columnconfigure(2, weight=1)
        self._upd_selall_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(lhdr, text="", variable=self._upd_selall_var,
                        width=24, command=self._toggle_all_updates).grid(
            row=0, column=0, padx=8, pady=6)
        for col, (text, width) in enumerate([("Application", 220), ("Current", 100), ("Available", 100)], 1):
            ctk.CTkLabel(lhdr, text=text, font=FL(12), text_color=C_ACCENT,
                         width=width, anchor="w").grid(row=0, column=col, padx=6, pady=6, sticky="w")
        self._upd_scroll = ctk.CTkScrollableFrame(left, fg_color="transparent", corner_radius=0)
        self._upd_scroll.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))
        self._upd_scroll.grid_columnconfigure(2, weight=1)
        self._upd_status_lbl = ctk.CTkLabel(left,
                                             text="Click 'Scan Updates' to check for available updates.",
                                             font=FM(11), text_color="gray50")
        self._upd_status_lbl.grid(row=2, column=0, pady=4)

        right = ctk.CTkFrame(body, fg_color=C_CARD, corner_radius=10,
                             border_color=C_BORDER, border_width=1)
        right.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        right.grid_columnconfigure(0, weight=1); right.grid_rowconfigure(1, weight=1)
        rh = ctk.CTkFrame(right, fg_color=C_CARD2, corner_radius=8)
        rh.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 0))
        rh.grid_columnconfigure(1, weight=1)
        ctk.CTkFrame(rh, width=3, corner_radius=1, fg_color=C_ACCENT).grid(
            row=0, column=0, sticky="ns", padx=(8, 6), pady=6)
        ctk.CTkLabel(rh, text="📋  Activity Log", font=FL(13), text_color=C_ACCENT).grid(
            row=0, column=1, padx=4, sticky="w", pady=6)
        ctk.CTkButton(rh, text="Clear", font=FM(11), width=60, height=26,
                      corner_radius=6, fg_color=C_BORDER, hover_color="#2a3560",
                      command=self._clear_upd_log).grid(row=0, column=2, padx=(4, 8), pady=6)
        self._upd_box = ctk.CTkTextbox(right,
                                        font=ctk.CTkFont(family="Courier New", size=11, weight="bold"),
                                        fg_color="#040810", text_color="#90caf9",
                                        corner_radius=8, border_color=C_BORDER, border_width=1)
        self._upd_box.grid(row=1, column=0, sticky="nsew", padx=6, pady=(4, 6))
        for tag, col in [("success", C_GREEN), ("warn", C_ORANGE), ("error", C_RED),
                         ("info", C_ACCENT), ("dim", "gray50")]:
            self._upd_box._textbox.tag_configure(tag, foreground=col)
        self._upd_log("Update Centre ready — click 'Scan Updates' to begin.", "info")

    # ── DIAGNOSIS TAB ─────────────────────────────────────────────────────────
    _DIAG_STATUS = {"ok": (C_GREEN, "OK"), "info": (C_ACCENT, "i"),
                    "warning": (C_ORANGE, "!"), "critical": (C_RED, "X")}

    def _build_diagnosis_tab(self):
        tab = self._maint_tabs.tab("🔬 Diagnosis")
        tab.configure(fg_color="transparent")
        tab.grid_columnconfigure(0, weight=1); tab.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(tab, text="PC Health Diagnosis", font=FT(17)).grid(row=0, column=0, pady=(6, 4))

        ctrl = ctk.CTkFrame(tab, fg_color=C_CARD, corner_radius=10,
                            border_color=C_BORDER, border_width=1)
        ctrl.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 6))
        ctrl.grid_columnconfigure(2, weight=1)
        self._diag_run_btn = ctk.CTkButton(ctrl, text="🔬  Run Diagnosis", font=FB(13),
                                            width=185, height=46, corner_radius=10,
                                            fg_color="#0a5abf", hover_color="#1269d3",
                                            command=self._start_diagnosis)
        self._diag_run_btn.grid(row=0, column=0, padx=10, pady=10)
        self._diag_prog = ctk.CTkProgressBar(ctrl, width=240, height=14, progress_color=C_ACCENT)
        self._diag_prog.grid(row=0, column=1, padx=(0, 10), pady=10)
        self._diag_prog.set(0)
        self._diag_status_lbl = ctk.CTkLabel(ctrl,
                                              text="Click 'Run Diagnosis' to scan your PC for issues.",
                                              font=FM(12), text_color="gray50", anchor="w")
        self._diag_status_lbl.grid(row=0, column=2, padx=4, sticky="w")
        self._diag_chip_frame = ctk.CTkFrame(ctrl, fg_color="transparent")
        self._diag_chip_frame.grid(row=0, column=3, padx=12, pady=10)

        body = ctk.CTkFrame(tab, fg_color="transparent")
        body.grid(row=2, column=0, sticky="nsew", padx=6, pady=(0, 6))
        body.grid_columnconfigure(0, weight=3); body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)
        self._diag_scroll = ctk.CTkScrollableFrame(body, fg_color="transparent", corner_radius=0)
        self._diag_scroll.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self._diag_scroll.grid_columnconfigure(0, weight=1, minsize=400)
        ctk.CTkLabel(self._diag_scroll, text="Run a diagnosis to check your system for issues.",
                     font=FL(13), text_color="gray45").grid(row=0, column=0, pady=60)

        log_panel = ctk.CTkFrame(body, fg_color=C_CARD, corner_radius=10,
                                  border_color=C_BORDER, border_width=1)
        log_panel.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        log_panel.grid_columnconfigure(0, weight=1); log_panel.grid_rowconfigure(1, weight=1)
        lh = ctk.CTkFrame(log_panel, fg_color=C_CARD2, corner_radius=8)
        lh.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 0))
        lh.grid_columnconfigure(1, weight=1)
        ctk.CTkFrame(lh, width=3, corner_radius=1, fg_color=C_ACCENT).grid(
            row=0, column=0, sticky="ns", padx=(8, 6), pady=6)
        ctk.CTkLabel(lh, text="📋  Diagnosis Log", font=FL(13), text_color=C_ACCENT).grid(
            row=0, column=1, sticky="w", pady=6)
        ctk.CTkButton(lh, text="Clear", font=FM(11), width=56, height=26,
                      corner_radius=6, fg_color=C_BORDER, hover_color="#2a3560",
                      command=self._clear_diag_log).grid(row=0, column=2, padx=(4, 8), pady=6)
        self._diag_log_box = ctk.CTkTextbox(log_panel,
                                             font=ctk.CTkFont(family="Courier New", size=11, weight="bold"),
                                             fg_color="#040810", text_color="#90caf9",
                                             corner_radius=8, border_color=C_BORDER, border_width=1)
        self._diag_log_box.grid(row=1, column=0, sticky="nsew", padx=6, pady=(4, 6))
        for tag, col in [("success", C_GREEN), ("warn", C_ORANGE), ("error", C_RED),
                         ("info", C_ACCENT), ("dim", "gray50")]:
            self._diag_log_box._textbox.tag_configure(tag, foreground=col)
        self._diag_log("Diagnosis ready — click 'Run Diagnosis' to begin.", "info")

    # ── REPAIR TAB ────────────────────────────────────────────────────────────
    def _build_repair_tab(self):
        tab = self._maint_tabs.tab("🛡 System Repair")
        tab.configure(fg_color="transparent")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(4, weight=1)
        ctk.CTkLabel(tab, text="System Repair Tools", font=FT(17)).grid(
            row=0, column=0, pady=(6, 6))

        warn = ctk.CTkFrame(tab, fg_color="#0e0900", corner_radius=10,
                            border_color="#3a2800", border_width=1)
        warn.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 8))
        warn.grid_columnconfigure(1, weight=1)
        ctk.CTkFrame(warn, width=4, corner_radius=2, fg_color=C_YELLOW).grid(
            row=0, column=0, sticky="ns", padx=(10, 8), pady=10)
        ctk.CTkLabel(warn,
                     text="Administrator rights required.  SFC repairs corrupted system files.  "
                          "DISM RestoreHealth re-downloads clean Windows files.  "
                          "Full Repair runs the correct sequence automatically.",
                     font=FM(12), text_color=C_YELLOW, wraplength=1100, justify="left", anchor="w"
                     ).grid(row=0, column=1, padx=(0, 14), pady=10, sticky="w")

        bg = ctk.CTkFrame(tab, fg_color=C_CARD, corner_radius=12,
                          border_color=C_BORDER, border_width=1)
        bg.grid(row=2, column=0, sticky="ew", padx=6, pady=(0, 5))
        bg.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)
        for col, (text, fg, hv, cmd) in enumerate([
            ("🛠  SFC Scan",    "#0b3d91", "#1153b5",
             lambda: threading.Thread(target=self._run_sfc, daemon=True).start()),
            ("🔍  DISM Scan",   "#3d0b7a", "#561494",
             lambda: threading.Thread(target=self._run_dism_scan, daemon=True).start()),
            ("🔧  DISM Restore","#7a0b3d", "#991452",
             lambda: threading.Thread(target=self._run_dism_restore, daemon=True).start()),
            ("⚡  Full Repair", "#0b5c1e", "#10762a",
             lambda: threading.Thread(target=self._run_full_repair, daemon=True).start()),
            ("🗑  Clean Temp",  "#8a3500", "#a84200",
             lambda: threading.Thread(target=self._clean_temp, daemon=True).start()),
        ]):
            ctk.CTkButton(bg, text=text, font=FB(12), command=cmd,
                          height=46, fg_color=fg, hover_color=hv, corner_radius=10).grid(
                row=0, column=col, padx=6, pady=10, sticky="ew")

        bg2 = ctk.CTkFrame(tab, fg_color=C_CARD, corner_radius=12,
                           border_color=C_BORDER, border_width=1)
        bg2.grid(row=3, column=0, sticky="ew", padx=6, pady=(0, 8))
        bg2.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(bg2, text="🧹  WU Cleanup  (removes old update packages)",
                      font=FB(12), height=46, corner_radius=10,
                      fg_color="#151d7a", hover_color="#1f2aa0",
                      command=lambda: threading.Thread(target=self._run_wu_cleanup, daemon=True).start()
                      ).grid(row=0, column=0, padx=(10, 5), pady=10, sticky="ew")
        ctk.CTkButton(bg2, text="🧠  RAM Cleaner  (flush standby working sets)",
                      font=FB(12), height=46, corner_radius=10,
                      fg_color="#005962", hover_color="#007a86",
                      command=lambda: threading.Thread(target=self._clean_ram, daemon=True).start()
                      ).grid(row=0, column=1, padx=(5, 10), pady=10, sticky="ew")

        log_wrap = ctk.CTkFrame(tab, fg_color=C_CARD2, corner_radius=10,
                                border_color=C_BORDER, border_width=1)
        log_wrap.grid(row=4, column=0, sticky="nsew", padx=6, pady=(0, 6))
        log_wrap.grid_columnconfigure(0, weight=1); log_wrap.grid_rowconfigure(2, weight=1)
        log_hdr = ctk.CTkFrame(log_wrap, fg_color=C_CARD2, corner_radius=0)
        log_hdr.grid(row=0, column=0, sticky="ew")
        log_hdr.grid_columnconfigure(1, weight=1)
        ctk.CTkFrame(log_hdr, width=3, corner_radius=1, fg_color=C_ACCENT).grid(
            row=0, column=0, sticky="ns", padx=(10, 7), pady=6)
        ctk.CTkLabel(log_hdr, text="📟  Live Output",
                     font=ctk.CTkFont(size=12, weight="bold"), text_color=C_ACCENT).grid(
            row=0, column=1, sticky="w", pady=6)
        ctk.CTkFrame(log_wrap, height=1, fg_color=C_BORDER, corner_radius=0).grid(
            row=1, column=0, sticky="ew")
        self._repair_box = ctk.CTkTextbox(log_wrap,
                                           font=ctk.CTkFont(family="Courier New", size=12, weight="bold"),
                                           fg_color="#040810", text_color="#90caf9",
                                           corner_radius=0, border_width=0)
        self._repair_box.grid(row=2, column=0, sticky="nsew")
        self._repair_log("System Repair ready — choose an action above.\n")

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 6: OPTIMISE (Gaming tweaks + Debloat)
    # ══════════════════════════════════════════════════════════════════════════
    GAMING_TWEAKS = [
        ("power_plan",  "Power Plan → High Performance",           'powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'),
        ("hags",        "Enable HAGS (Hardware GPU Scheduling)",    r'reg add "HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers" /v HwSchMode /t REG_DWORD /d 2 /f'),
        ("gamebar",     "Disable Xbox Game Bar",                    r'reg add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\GameDVR" /v AppCaptureEnabled /t REG_DWORD /d 0 /f'),
        ("gamedvr",     "Disable Game DVR recording",               r'reg add "HKCU\System\GameConfigStore" /v GameDVR_Enabled /t REG_DWORD /d 0 /f'),
        ("nagle",       "Disable Nagle Algorithm (lower latency)",  r'reg add "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces" /v TcpAckFrequency /t REG_DWORD /d 1 /f'),
        ("sysmain",     "Disable SysMain / Superfetch",             'sc config SysMain start= disabled & net stop SysMain'),
        ("wsearch",     "Disable Windows Search Indexer",           'sc config WSearch start= disabled & net stop WSearch'),
        ("gpu_prio",    "GPU & CPU Priority → High for Games",      r'reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games" /v "GPU Priority" /t REG_DWORD /d 8 /f'),
        ("fso",         "Disable Fullscreen Optimisations",         r'reg add "HKCU\System\GameConfigStore" /v GameDVR_FSEBehaviorMode /t REG_DWORD /d 2 /f'),
        ("visual_perf", "Visual Effects → Best Performance",        r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" /v VisualFXSetting /t REG_DWORD /d 2 /f'),
        ("mouse_accel", "Disable Mouse Acceleration",               r'reg add "HKCU\Control Panel\Mouse" /v MouseSpeed /t REG_SZ /d 0 /f'),
    ]
    DEBLOAT_ITEMS = [
        ("cortana",      "Cortana",                    "Microsoft.549981C3F5F10"),
        ("copilot",      "Microsoft Copilot",           "Microsoft.Copilot"),
        ("copilot_btn",  "Copilot taskbar button",      None),
        ("recall",       "Windows Recall (AI)",         None),
        ("xbox_app",     "Xbox App",                    "Microsoft.XboxApp"),
        ("xbox_overlay", "Xbox Game Overlay",           "Microsoft.XboxGameOverlay"),
        ("bing_news",    "Bing News",                   "Microsoft.BingNews"),
        ("bing_weather", "Bing Weather",                "Microsoft.BingWeather"),
        ("teams",        "Teams (personal)",            "MicrosoftTeams"),
        ("chat_icon",    "Chat/Teams taskbar icon",     None),
        ("widgets",      "Widgets on taskbar",          None),
        ("mail",         "Mail and Calendar",           "microsoft.windowscommunicationsapps"),
        ("people",       "People app",                  "Microsoft.People"),
        ("movies",       "Movies & TV",                 "Microsoft.ZuneVideo"),
        ("groove",       "Groove Music",                "Microsoft.ZuneMusic"),
        ("solitaire",    "Solitaire Collection",        "Microsoft.MicrosoftSolitaireCollection"),
        ("tips",         "Tips app",                    "Microsoft.Getstarted"),
        ("feedback",     "Feedback Hub",                "Microsoft.WindowsFeedbackHub"),
        ("clipchamp",    "Clipchamp",                   "Clipchamp.Clipchamp"),
        ("ads",          "Personalised Ads",            None),
        ("telemetry",    "Telemetry (DiagTrack)",       None),
        ("onedrive_run", "OneDrive auto-start",         None),
        ("start_ads",    "Start Menu suggestions/ads",  None),
    ]

    def _build_optimise_section(self):
        section = self._make_section("optimise")
        section.grid_columnconfigure((0, 1), weight=1)
        section.grid_rowconfigure(0, weight=0)   # title fixed
        section.grid_rowconfigure(1, weight=1)   # panels expand
        section.grid_rowconfigure(2, weight=0)   # log box fixed
        ctk.CTkLabel(section, text="Windows Optimisation", font=FT(17)).grid(
            row=0, column=0, columnspan=2, pady=(10, 4))

        def _panel(col, title, tc, items, var_dict, btn_text, btn_color, btn_hov, cmd):
            outer = ctk.CTkFrame(section, fg_color=C_CARD, corner_radius=12,
                                  border_color=C_BORDER, border_width=1)
            outer.grid(row=1, column=col, sticky="nsew",
                       padx=(8 if col == 0 else 4, 4 if col == 0 else 8), pady=(0, 4))
            outer.grid_columnconfigure(0, weight=1); outer.grid_rowconfigure(1, weight=1)
            h = ctk.CTkFrame(outer, fg_color=C_CARD2, corner_radius=8)
            h.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
            h.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(h, text=title, font=FH(14), text_color=tc).grid(
                row=0, column=0, padx=12, pady=8, sticky="w")
            all_v = tk.BooleanVar(value=True)
            def _tog(v=all_v, d=var_dict):
                for x in d.values(): x.set(v.get())
            ctk.CTkCheckBox(h, text="All", variable=all_v, font=FM(12), command=_tog).grid(
                row=0, column=1, padx=10, pady=8, sticky="e")
            sc = ctk.CTkScrollableFrame(outer, fg_color="transparent", corner_radius=0)
            sc.grid(row=1, column=0, sticky="nsew", padx=4, pady=0)
            sc.grid_columnconfigure(0, weight=1)
            for i, (key, label, *_) in enumerate(items):
                v = tk.BooleanVar(value=True); var_dict[key] = v
                bg = C_CARD if i % 2 == 0 else C_CARD2
                rf = ctk.CTkFrame(sc, fg_color=bg, corner_radius=6, height=34)
                rf.grid(row=i, column=0, sticky="ew", padx=2, pady=1)
                rf.grid_columnconfigure(0, weight=1)
                ctk.CTkCheckBox(rf, text=label, variable=v, font=FM(12), text_color="gray85").grid(
                    row=0, column=0, padx=10, pady=5, sticky="w")
            ctk.CTkButton(outer, text=btn_text, font=FB(13), height=46,
                          fg_color=btn_color, hover_color=btn_hov, corner_radius=10, command=cmd).grid(
                row=2, column=0, padx=10, pady=8, sticky="ew")

        _panel(0, "🎮  Gaming Optimisations", C_ACCENT, self.GAMING_TWEAKS, self._game_vars,
               "⚡  Apply Selected Gaming Tweaks", "#0b5c1e", "#10762a",
               lambda: threading.Thread(target=self._apply_gaming, daemon=True).start())
        _panel(1, "🗑  Remove Bloatware", C_RED, self.DEBLOAT_ITEMS, self._bloat_vars,
               "🗑  Remove Selected Items", "#7a0b0b", "#991010", self._confirm_debloat)

        self._opt_box = ctk.CTkTextbox(
            section, font=ctk.CTkFont(family="Courier New", size=11, weight="bold"),
            fg_color="#040810", text_color="#90caf9",
            corner_radius=0, border_color=C_BORDER, border_width=1, height=110)
        self._opt_box.grid(row=2, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 6))
        self._opt_log("Optimisation Centre ready — select items and apply.\n")

    # ══════════════════════════════════════════════════════════════════════════
    # SCAN & REFRESH
    # ══════════════════════════════════════════════════════════════════════════
    def _scan_system(self):
        self._set_status("🔍 Scanning hardware…")
        self.after(0, lambda: self._scan_btn.configure(state="disabled", text="Scanning…"))
        try:
            self.system_data = self.sys_info.get_all_info()
            self.after(0, self._refresh_all_tabs)
            self._set_status("Scan complete")
            if not getattr(self, "_ov_refresh_running", False):
                self._ov_refresh_running = True
                self.after(3000, self._auto_refresh_overview)
            if not getattr(self, "_proc_refresh_running", False):
                self._proc_refresh_running = True
                self.after(2000, self._auto_refresh_processes)
            if not getattr(self, "_gpu_usage_loop_running", False):
                self._gpu_usage_loop_running = True
                self.after(1500, self._gpu_usage_loop_tick)
        except Exception as exc:
            self._set_status(f"Error: {exc}")
        finally:
            self.after(0, lambda: self._scan_btn.configure(state="normal", text="🔍  Scan System"))

    def _auto_refresh_overview(self):
        def _bg():
            try:
                m = self.sys_info.get_live_metrics()
                if self.system_data:
                    cpu = self.system_data.get("cpu", {})
                    cpu["usage_percent"]  = m["cpu_usage"]
                    cpu["per_core_usage"] = m["cpu_per_core"]
                    if m["cpu_freq_mhz"]: cpu["current_speed_mhz"] = m["cpu_freq_mhz"]
                    ram = self.system_data.get("ram", {})
                    ram["usage_percent"] = m["ram_pct"]
                    ram["used_gb"]       = m["ram_used_gb"]
                    ram["available_gb"]  = m["ram_free_gb"]
                self.after(0, self._refresh_overview_partial)
                import datetime as _dt; ts = _dt.datetime.now().strftime("%H:%M:%S")
                pulse = "●" if getattr(self, "_live_pulse", True) else "○"
                self._live_pulse = not getattr(self, "_live_pulse", True)
                tc = C_GREEN if pulse == "●" else C_ACCENT
                self.after(0, lambda t=ts, p=pulse, c=tc:
                    self._ov_refresh_ts.configure(text=f"Live  {p}  {t}", text_color=c))
            except Exception:
                pass
        threading.Thread(target=_bg, daemon=True).start()
        self.after(3000, self._auto_refresh_overview)

    def _gpu_usage_loop_tick(self):
        def _bg():
            try:
                usage = self.sys_info.get_gpu_usage()
                self._last_gpu_usage = usage
                gpu_pct = usage.get("gpu_load_pct")
                self.after(0, lambda p=gpu_pct: self._ov_update_usage_badge("gpu", p))
                self.after(0, lambda p=gpu_pct: self._update_gpu_load_inplace(p))
            except Exception: pass
        threading.Thread(target=_bg, daemon=True).start()
        self.after(3000, self._gpu_usage_loop_tick)

    def _update_gpu_load_inplace(self, pct):
        if not hasattr(self, "_gpu_load_bar") or self._gpu_load_bar is None: return
        if pct is None: return
        try:
            color = C_GREEN if pct < 50 else C_ACCENT if pct < 75 else C_ORANGE if pct < 90 else C_RED
            self._gpu_load_bar.configure(progress_color=color)
            self._gpu_load_bar.set(min(float(pct) / 100.0, 1.0))
            if hasattr(self, "_gpu_load_lbl") and self._gpu_load_lbl:
                self._gpu_load_lbl.configure(text=f"{pct:.0f}%", text_color=color)
        except Exception: pass

    def _manual_refresh_overview(self):
        self.after(0, lambda: self._ov_refresh_ts.configure(text="Refreshing…", text_color=C_ACCENT))
        try:
            fresh = self.sys_info.get_all_info()
            if self.system_data: self.system_data.update(fresh)
            else: self.system_data = fresh
            self.after(0, self._refresh_overview)
            import datetime as _dt; ts = _dt.datetime.now().strftime("%H:%M:%S")
            self.after(0, lambda: self._ov_refresh_ts.configure(
                text=f"Last updated: {ts}", text_color="gray45"))
        except Exception as e:
            self.after(0, lambda: self._ov_refresh_ts.configure(text=f"Error: {e}", text_color=C_RED))

    def _refresh_all_tabs(self):
        if not self.system_data: return
        self._refresh_overview()
        self._refresh_cpu()
        self._refresh_gpu()
        self._refresh_ram()
        self._refresh_storage()
        self._refresh_network()

    def _refresh_overview_partial(self):
        if not self.system_data: return
        cpu = self.system_data.get("cpu", {}); ram = self.system_data.get("ram", {})
        if "error" not in cpu:
            u = cpu.get("usage_percent", 0)
            self._ov_set("cpu",
                cpu.get("name", "N/A") + "\n" +
                "Cores: " + str(cpu.get("cores", "?")) + "  |  Threads: " + str(cpu.get("threads", "?")) + "\n" +
                "Clock: " + str(cpu.get("current_speed_mhz", "?")) + " MHz  |  Max: " + str(cpu.get("max_speed_mhz", "?")) + " MHz\n" +
                "L2: " + str(cpu.get("l2_cache_kb", "?")) + " KB  |  L3: " + str(cpu.get("l3_cache_kb", "?")) + " KB",
                color=C_ACCENT, pct=u)
            self._ov_update_usage_badge("cpu", u)
        if "error" not in ram:
            up = ram.get("usage_percent", 0); mods = ram.get("modules", [])
            mtype = mods[0].get("memory_type", "?") if mods else "?"
            spd   = mods[0].get("speed_mhz", "?")   if mods else "?"
            mfr   = mods[0].get("manufacturer", "?") if mods else "?"
            ch    = ram.get("channel_info", ""); xmp = ram.get("xmp_expo", "")
            xs = ("\nXMP/EXPO: " + xmp.split("(")[0].strip() if "Enabled" in xmp
                  else "\nXMP/EXPO: Off" if "Disabled" in xmp else "")
            self._ov_set("ram",
                "Total: " + str(ram.get("total_gb", "?")) + " GB  |  Used: " + str(ram.get("used_gb", "?")) + " GB\n" +
                "Free: " + str(ram.get("available_gb", "?")) + " GB  |  " + ch + "\n" +
                "Type: " + mtype + "  |  " + str(spd) + " MHz  |  " + mfr + xs,
                color="#00bcd4", pct=up, bar_color="#00bcd4")
        cpu_u = cpu.get("usage_percent", 0) if "error" not in cpu else 0
        ram_u = ram.get("usage_percent", 0) if "error" not in ram else 0
        cores = cpu.get("per_core_usage", []) if "error" not in cpu else []
        peak  = max(cores) if cores else 0
        self._ov_set("perf",
            "CPU Load:  " + str(cpu_u) + "%\n" +
            "RAM Load:  " + str(ram_u) + "%\n" +
            "Peak Core: " + str(peak)  + "%\n" +
            "RAM Free:  " + str(ram.get("available_gb", "?")) + " GB",
            color=C_ORANGE, pct=cpu_u, bar_color=C_ORANGE)

    def _refresh_overview(self):
        if not self.system_data: return
        d   = self.system_data; cpu = d.get("cpu", {}); ram = d.get("ram", {})
        mb  = d.get("motherboard", {}); os_i = d.get("os", {})

        if "error" not in cpu:
            u = cpu.get("usage_percent", 0)
            self._ov_set("cpu",
                cpu.get("name", "N/A") + "\n" +
                "Cores: " + str(cpu.get("cores", "?")) + "  |  Threads: " + str(cpu.get("threads", "?")) + "\n" +
                "Clock: " + str(cpu.get("current_speed_mhz", "?")) + " MHz  |  Max: " + str(cpu.get("max_speed_mhz", "?")) + " MHz\n" +
                "L2: " + str(cpu.get("l2_cache_kb", "?")) + " KB  |  L3: " + str(cpu.get("l3_cache_kb", "?")) + " KB",
                color=C_ACCENT, pct=u)
            self._ov_update_usage_badge("cpu", u)
        else:
            self._ov_set("cpu", "Error reading CPU info", color=C_RED)

        gpus = d.get("gpu", []); g = gpus[0] if gpus else {}
        if gpus and "error" not in g:
            self._ov_set("gpu",
                str(g.get("name", "N/A")) + "\n" +
                "VRAM: " + str(g.get("vram_gb", "?")) + " GB\n" +
                "Resolution: " + str(g.get("current_resolution", "N/A")) +
                "  |  " + str(g.get("refresh_rate_hz", "?")) + " Hz\n" +
                "Driver: " + str(g.get("driver_version", "N/A")) + "\n" +
                "Driver Date: " + str(g.get("driver_date", "N/A")),
                color="#7c4dff", pct=0, bar_color="#7c4dff")
            self._ov_update_usage_badge("gpu", self._last_gpu_usage.get("gpu_load_pct"))
        else:
            self._ov_set("gpu", g.get("error", "No GPU detected") if g else "No GPU detected", color=C_RED)

        if "error" not in ram:
            up = ram.get("usage_percent", 0); mods = ram.get("modules", [])
            mtype = mods[0].get("memory_type", "?") if mods else "?"
            spd   = mods[0].get("speed_mhz", "?")   if mods else "?"
            mfr   = mods[0].get("manufacturer", "?") if mods else "?"
            ch    = ram.get("channel_info", ""); xmp = ram.get("xmp_expo", "")
            xs = ("\nXMP/EXPO: " + xmp.split("(")[0].strip() if "Enabled" in xmp
                  else "\nXMP/EXPO: Off" if "Disabled" in xmp else "")
            self._ov_set("ram",
                "Total: " + str(ram.get("total_gb", "?")) + " GB  |  Used: " + str(ram.get("used_gb", "?")) + " GB\n" +
                "Free: " + str(ram.get("available_gb", "?")) + " GB  |  " + ch + "\n" +
                "Type: " + mtype + "  |  " + str(spd) + " MHz  |  " + mfr + xs,
                color="#00bcd4", pct=up, bar_color="#00bcd4")
        else:
            self._ov_set("ram", "Error reading RAM info", color=C_RED)

        self._ov_update_storage(d.get("storage", []))

        if "error" not in mb:
            self._ov_set("mb",
                str(mb.get("manufacturer", "N/A")) + "  " + str(mb.get("product", "")) + "\n" +
                "Version: " + str(mb.get("version", "N/A")) + "\n" +
                "BIOS: " + str(mb.get("bios_version", "N/A")) + "\n" +
                "BIOS Date: " + str(mb.get("bios_date", "N/A")) + "\n" +
                "By: " + str(mb.get("bios_manufacturer", "N/A")), color="#78909c")
        else:
            self._ov_set("mb", "Error reading Motherboard info", color=C_RED)

        nets = d.get("network", []); up_nets = [n for n in nets if n.get("is_up") and "error" not in n]
        if up_nets:
            lines = ["Active adapters: " + str(len(up_nets))]
            for n in up_nets[:2]:
                lines.append("\n" + str(n.get("name", "?")))
                lines.append("  IP: "    + str(n.get("ipv4", "N/A")))
                lines.append("  Speed: " + str(n.get("speed_mbps", "?")) + " Mbps")
            self._ov_set("net", "\n".join(lines), color=C_GREEN, pct=100, bar_color=C_GREEN)
        else:
            self._ov_set("net", "No active adapters detected", color="gray50", pct=0)

        if "error" not in os_i:
            self._ov_set("os",
                str(os_i.get("name", "N/A")) + "\n" +
                "Build: " + str(os_i.get("build", "?")) + "  |  " + str(os_i.get("architecture", "?")) + "\n" +
                "Uptime: " + str(os_i.get("system_uptime", "N/A")) + "\n" +
                "User: " + str(os_i.get("registered_user", "N/A")) + "\n" +
                "PC: " + str(os_i.get("computer_name", "N/A")), color="#5c6bc0")
        else:
            self._ov_set("os", "Error reading OS info", color=C_RED)

        cpu_u = cpu.get("usage_percent", 0) if "error" not in cpu else 0
        ram_u = ram.get("usage_percent", 0) if "error" not in ram else 0
        cores = cpu.get("per_core_usage", []) if "error" not in cpu else []
        peak  = max(cores) if cores else 0
        self._ov_set("perf",
            "CPU Load:  " + str(cpu_u) + "%\n" +
            "RAM Load:  " + str(ram_u) + "%\n" +
            "Peak Core: " + str(peak)  + "%\n" +
            "RAM Free:  " + str(ram.get("available_gb", "?")) + " GB",
            color=C_ORANGE, pct=cpu_u, bar_color=C_ORANGE)

        bios_ver  = mb.get("bios_version", "N/A")  if "error" not in mb  else "N/A"
        bios_date = mb.get("bios_date", "N/A")      if "error" not in mb  else "N/A"
        os_build  = os_i.get("build", "N/A")        if "error" not in os_i else "N/A"
        last_boot = os_i.get("last_boot", "N/A")    if "error" not in os_i else "N/A"
        install   = os_i.get("install_date", "N/A") if "error" not in os_i else "N/A"
        self._ov_set("power",
            "BIOS: "      + str(bios_ver)  + "\n" +
            "BIOS Date: " + str(bios_date) + "\n" +
            "OS Build: "  + str(os_build)  + "\n" +
            "Last Boot: " + str(last_boot) + "\n" +
            "Install: "   + str(install), color="#ec407a")

    # ── Hardware detail refreshers ────────────────────────────────────────────
    def _refresh_cpu(self):
        self._clear(self._cpu_frame); self._cpu_frame.grid_columnconfigure((0, 1), weight=1)
        cpu = self.system_data.get("cpu", {})
        if "error" in cpu:
            ctk.CTkLabel(self._cpu_frame, text=f"Error: {cpu['error']}",
                         font=F(13), text_color=C_RED).grid(row=0, column=0); return
        fields = [
            ("Processor Name",  cpu.get("name")),
            ("Manufacturer",    cpu.get("manufacturer")),
            ("Physical Cores",  cpu.get("cores")),
            ("Logical Threads", cpu.get("threads")),
            ("Max Clock",       f"{cpu.get('max_speed_mhz','?')} MHz"),
            ("Current Clock",   f"{cpu.get('current_speed_mhz','?')} MHz"),
            ("Socket",          cpu.get("socket")),
            ("L2 Cache",        f"{cpu.get('l2_cache_kb','?')} KB"),
            ("L3 Cache",        f"{cpu.get('l3_cache_kb','?')} KB"),
            ("Architecture",    cpu.get("architecture")),
        ]
        for i, (l, v) in enumerate(fields): self._kv(self._cpu_frame, i, l, v)
        r = len(fields)
        self._pbar(self._cpu_frame, r, f"CPU Usage ({cpu.get('usage_percent','?')}%)",
                   cpu.get("usage_percent", 0)); r += 1
        per = cpu.get("per_core_usage", [])
        if per:
            self._section_lbl(self._cpu_frame, r, f"Per-Core Usage  ({len(per)} threads)"); r += 2
            for ci, u in enumerate(per): self._pbar(self._cpu_frame, r, f"  Core {ci}", u); r += 1

    def _refresh_gpu(self):
        self._clear(self._gpu_frame); self._gpu_frame.grid_columnconfigure((0, 1), weight=1)
        gpus = self.system_data.get("gpu", []); r = 0; self._gpu_load_bar = None; self._gpu_load_lbl = None
        for idx, gpu in enumerate(gpus):
            if "error" in gpu: continue
            if len(gpus) > 1: self._section_lbl(self._gpu_frame, r, f"GPU {idx+1}"); r += 2
            for l, v in [
                ("GPU Name",       gpu.get("name")),
                ("Video Memory",   f"{gpu.get('vram_gb','?')} GB"),
                ("Driver Version", gpu.get("driver_version")),
                ("Driver Date",    gpu.get("driver_date")),
                ("Resolution",     gpu.get("current_resolution")),
                ("Refresh Rate",   f"{gpu.get('refresh_rate_hz','?')} Hz"),
                ("Status",         gpu.get("status")),
            ]:
                self._kv(self._gpu_frame, r, l, v); r += 1
            gpu_load = self._last_gpu_usage.get("gpu_load_pct")
            vram_ctrl = self._last_gpu_usage.get("vram_ctrl_pct")
            self._section_lbl(self._gpu_frame, r, "GPU Usage"); r += 2
            pct_val = gpu_load if gpu_load is not None else 0
            color = C_GREEN if pct_val < 50 else C_ACCENT if pct_val < 75 else C_ORANGE if pct_val < 90 else C_RED
            ctk.CTkLabel(self._gpu_frame, text="GPU Core Load:", font=FL(13),
                         text_color=C_ACCENT, anchor="w").grid(
                row=r, column=0, padx=(16, 6), pady=3, sticky="w")
            load_cont = ctk.CTkFrame(self._gpu_frame, fg_color="transparent")
            load_cont.grid(row=r, column=1, padx=(4, 14), pady=3, sticky="ew")
            load_cont.grid_columnconfigure(0, weight=1)
            self._gpu_load_bar = ctk.CTkProgressBar(load_cont, height=14, progress_color=color)
            self._gpu_load_bar.grid(row=0, column=0, sticky="ew")
            self._gpu_load_bar.set(min(float(pct_val) / 100.0, 1.0))
            lbl_txt = f"{gpu_load:.0f}%" if gpu_load is not None else "N/A (LHM not running)"
            self._gpu_load_lbl = ctk.CTkLabel(load_cont, text=lbl_txt,
                                               font=FM(11), text_color=color if gpu_load is not None else "gray45",
                                               width=160, anchor="w")
            self._gpu_load_lbl.grid(row=0, column=1, padx=(8, 0))
            r += 1
            if vram_ctrl is not None:
                self._pbar(self._gpu_frame, r, "VRAM Controller", vram_ctrl, 70, 90); r += 1

    def _refresh_ram(self):
        self._clear(self._ram_frame); self._ram_frame.grid_columnconfigure((0, 1), weight=1)
        ram = self.system_data.get("ram", {})
        if "error" in ram:
            ctk.CTkLabel(self._ram_frame, text=f"Error: {ram['error']}",
                         font=F(13), text_color=C_RED).grid(row=0, column=0); return
        r = 0
        self._section_lbl(self._ram_frame, r, "Memory Summary"); r += 2
        for l, v in [("Total", f"{ram.get('total_gb','?')} GB"), ("Used", f"{ram.get('used_gb','?')} GB"),
                     ("Free", f"{ram.get('available_gb','?')} GB"), ("Channel Mode", ram.get("channel_info", "?"))]:
            self._kv(self._ram_frame, r, l, v); r += 1
        self._pbar(self._ram_frame, r, f"Usage ({ram.get('usage_percent','?')}%)",
                   ram.get("usage_percent", 0)); r += 1
        xmp = ram.get("xmp_expo", "N/A")
        xc = C_GREEN if "Enabled" in xmp else C_ORANGE if "Disabled" in xmp else "gray60"
        ctk.CTkLabel(self._ram_frame, text="XMP / EXPO:", font=FL(13), text_color=C_ACCENT, anchor="w").grid(
            row=r, column=0, padx=(16, 6), pady=2, sticky="w")
        ctk.CTkLabel(self._ram_frame, text=xmp, font=FV(13), text_color=xc, anchor="w", wraplength=500).grid(
            row=r, column=1, padx=(4, 14), pady=2, sticky="w"); r += 1
        for mi, mod in enumerate(ram.get("modules", [])):
            mtype = mod.get("memory_type", "?"); spd = mod.get("speed_mhz", "?"); mfr = mod.get("manufacturer", "?")
            self._section_lbl(self._ram_frame, r,
                               f"Slot {mi+1}  -  {mod.get('device_locator','?')}  ({mtype} @ {spd} MHz)"); r += 2
            for l, v in [("Capacity", f"{mod.get('capacity_gb','?')} GB"), ("Type", mtype),
                         ("Speed", f"{spd} MHz"), ("Manufacturer", mfr),
                         ("Part Number", mod.get("part_number", "N/A")), ("Bank", mod.get("bank_label", "?"))]:
                self._kv(self._ram_frame, r, l, v); r += 1

    def _refresh_storage(self):
        self._clear(self._stor_frame); self._stor_frame.grid_columnconfigure((0, 1), weight=1); r = 0
        for disk in self.system_data.get("storage", []):
            if "error" in disk: continue
            self._section_lbl(self._stor_frame, r, f"💿 {disk.get('model','N/A')}"); r += 2
            smart_txt, smart_col = self._get_smart_status(disk.get("model", ""), disk.get("serial", ""))
            ctk.CTkLabel(self._stor_frame, text="SMART Health:", font=FL(13), text_color=C_ACCENT, anchor="w").grid(
                row=r, column=0, padx=(16, 6), pady=2, sticky="w")
            ctk.CTkLabel(self._stor_frame, text=smart_txt, font=FV(13), text_color=smart_col, anchor="w").grid(
                row=r, column=1, padx=(4, 14), pady=2, sticky="w"); r += 1
            for l, v in [("Capacity", f"{disk.get('size_gb','?')} GB"), ("Interface", disk.get("interface")),
                         ("Media Type", disk.get("media_type")), ("Serial", disk.get("serial"))]:
                self._kv(self._stor_frame, r, l, v); r += 1
            for p in disk.get("partitions", []):
                ctk.CTkLabel(self._stor_frame, text=f"  Partition: {p.get('mountpoint','?')}",
                             font=FL(13), text_color="#a5d6a7").grid(
                    row=r, column=0, columnspan=2, padx=16, pady=(8, 2), sticky="w"); r += 1
                for l, v in [("File System", p.get("fstype")), ("Total", f"{p.get('total_gb','?')} GB"),
                             ("Used", f"{p.get('used_gb','?')} GB"), ("Free", f"{p.get('free_gb','?')} GB")]:
                    self._kv(self._stor_frame, r, l, v); r += 1
                self._pbar(self._stor_frame, r, f"Usage ({p.get('percent','?')}%)", p.get("percent", 0)); r += 1

    def _refresh_network(self):
        # Title is now outside the scrollframe — clear everything and rebuild from row 0
        for w in self._net_frame.winfo_children():
            w.destroy()
        r = 0
        for net in self.system_data.get("network", []):
            if "error" in net: continue
            color  = C_GREEN if net.get("is_up") else "gray50"
            status = "CONNECTED" if net.get("is_up") else "DISCONNECTED"
            ctk.CTkLabel(self._net_frame, text=f"  {net.get('name','?')}   {status}",
                         font=FL(13), text_color=color).grid(
                row=r, column=0, columnspan=2, padx=16, pady=(12, 2), sticky="w"); r += 1
            for l, v in [("IPv4", net.get("ipv4")), ("MAC", net.get("mac")),
                         ("Speed", f"{net.get('speed_mbps','?')} Mbps"),
                         ("Sent MB", net.get("bytes_sent")), ("Recv MB", net.get("bytes_recv"))]:
                self._kv(self._net_frame, r, l, v); r += 1

    # ══════════════════════════════════════════════════════════════════════════
    # UPDATE ACTIONS
    # ══════════════════════════════════════════════════════════════════════════
    def _clear_upd_log(self):
        self._upd_box.configure(state="normal"); self._upd_box.delete("1.0", "end")

    _SPINNER_CHARS = frozenset({"-", "\\", "|", "/", "⠻", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"})

    def _upd_log(self, text, tag=""):
        import datetime as _dt; clean = text.strip()
        if clean in self._SPINNER_CHARS or clean == "": return
        ts = _dt.datetime.now().strftime("%H:%M:%S")
        def _do():
            tb = self._upd_box._textbox
            tb.insert("end", f"[{ts}] ", "dim")
            tb.insert("end", clean + "\n", tag) if tag else tb.insert("end", clean + "\n")
            self._upd_box.see("end")
        self.after(0, _do)

    def _toggle_all_updates(self):
        s = self._upd_selall_var.get()
        for v in self._upd_vars.values(): v.set(s)

    def _populate_update_list(self, updates):
        for w in self._upd_scroll.winfo_children(): w.destroy()
        self._upd_vars.clear(); self._upd_rows = updates
        if not updates:
            ctk.CTkLabel(self._upd_scroll, text="All applications are up to date!",
                         font=FL(13), text_color=C_GREEN).grid(row=0, column=0, columnspan=4, pady=20)
            self.after(0, lambda: self._upd_status_lbl.configure(text="All up to date!"))
            return
        self.after(0, lambda: self._upd_status_lbl.configure(
            text=f"Found {len(updates)} update(s) — select and click 'Update Selected'"))
        for i, u in enumerate(updates):
            bg = C_CARD if i % 2 == 0 else C_CARD2
            rf = ctk.CTkFrame(self._upd_scroll, fg_color=bg, corner_radius=6, height=34)
            rf.grid(row=i, column=0, sticky="ew", padx=2, pady=1, columnspan=4)
            rf.grid_columnconfigure(2, weight=1)
            var = tk.BooleanVar(value=True); pkg_id = u.get("id", f"_idx_{i}")
            self._upd_vars[pkg_id] = var
            ctk.CTkCheckBox(rf, text="", variable=var, width=24).grid(row=0, column=0, padx=8, pady=4)
            name = u.get("name", "?"); name = name[:36] + "..." if len(name) > 38 else name
            ctk.CTkLabel(rf, text=name, font=FM(12), anchor="w", width=240).grid(
                row=0, column=1, padx=4, pady=4, sticky="w")
            cur = str(u.get("current_version", "?")); cur = cur[:13] + "..." if len(cur) > 14 else cur
            ctk.CTkLabel(rf, text=cur, font=FM(11), text_color="gray60", anchor="w", width=110).grid(
                row=0, column=2, padx=4, pady=4, sticky="w")
            avail = str(u.get("available_version", "?")); avail = avail[:13] + "..." if len(avail) > 14 else avail
            ctk.CTkLabel(rf, text=avail, font=FM(11), text_color=C_GREEN, anchor="w", width=110).grid(
                row=0, column=3, padx=4, pady=4, sticky="w")

    def _check_all_updates(self):
        threading.Thread(target=self.__check_updates_bg, daemon=True).start()

    def __check_updates_bg(self):
        self._upd_log("--- Scanning for app updates ---", "dim")
        self._upd_log("Scanning for app updates…", "info")
        self.after(0, lambda: self._upd_status_lbl.configure(text="Scanning…"))
        result = self.upd_manager.check_winget_updates()
        if isinstance(result, dict) and "error" in result:
            self._upd_log(f"Error: {result['error']}", "warn")
            self.after(0, lambda: self._populate_update_list([]))
        else:
            updates = result if isinstance(result, list) else []
            self.after(0, lambda: self._populate_update_list(updates))
            self._upd_log(f"Found {len(updates)} update(s) available." if updates
                          else "All apps are up to date.", "success")
        amd = self.upd_manager.check_amd_driver()
        if "name" in amd:
            self._upd_log(f"AMD Driver: {amd.get('name','?')}", "info")
            self._upd_log(f"  Current: {amd.get('current_driver','?')}  |  Date: {amd.get('driver_date','?')}")
        self._set_status("Update scan done")

    def _install_selected(self):
        selected = [uid for uid, var in self._upd_vars.items() if var.get()]
        if not selected: messagebox.showinfo("Nothing Selected", "Tick at least one app first."); return
        self._upd_log(f"Installing {len(selected)} selected update(s)…", "info")
        def _bg():
            for pkg_id in selected:
                name = next((u["name"] for u in self._upd_rows if u.get("id") == pkg_id), pkg_id)
                self._upd_log(f"  {name}")
                r = self.upd_manager.install_winget_update(pkg_id,
                    cb=lambda l, t="": self._upd_log(f"    {l}", t) if l else None)
                (self._upd_log(f"    Done: {name}", "success") if r.get("success")
                 else self._upd_log(f"    Failed: {name}", "warn"))
            self._upd_log("Selected updates complete.", "success"); self._set_status("Updates installed")
        threading.Thread(target=_bg, daemon=True).start()

    def _install_all_updates(self):
        self._upd_log("Installing ALL available updates…", "info")
        def _bg():
            self.upd_manager.install_all_winget_updates(
                cb=lambda l, t="": self._upd_log(f"  {l}", t) if l else None)
            self._upd_log("All updates complete.", "success"); self._set_status("All updates done")
        threading.Thread(target=_bg, daemon=True).start()

    def _update_amd(self):
        self._upd_log("Updating AMD Software Adrenalin…", "info")
        def _bg():
            r = self.upd_manager.install_amd_software(
                cb=lambda l, t="": self._upd_log(f"  {l}", t) if l else None)
            (self._upd_log("AMD updated. Restart to finish.", "success") if r.get("success")
             else self._upd_log("Check log. Manual: https://www.amd.com/en/support/download/drivers.html", "warn"))
        threading.Thread(target=_bg, daemon=True).start()

    def _open_win_upd(self):
        self.upd_manager.open_windows_update()
        self._upd_log("Windows Update opened.", "info")

    # ══════════════════════════════════════════════════════════════════════════
    # DIAGNOSIS ACTIONS
    # ══════════════════════════════════════════════════════════════════════════
    def _diag_log(self, text, tag=""):
        import datetime as _dt; clean = text.strip()
        if not clean: return
        ts = _dt.datetime.now().strftime("%H:%M:%S")
        def _do():
            tb = self._diag_log_box._textbox
            tb.insert("end", f"[{ts}] ", "dim")
            tb.insert("end", clean + "\n", tag) if tag else tb.insert("end", clean + "\n")
            self._diag_log_box.see("end")
        self.after(0, _do)

    def _clear_diag_log(self):
        self._diag_log_box.configure(state="normal"); self._diag_log_box.delete("1.0", "end")

    def _start_diagnosis(self):
        self._diag_run_btn.configure(state="disabled", text="Scanning…")
        self._diag_status_lbl.configure(text="Starting diagnosis…", text_color=C_ACCENT)
        self._diag_prog.set(0); self._clear_diag_log()
        self._diag_log("Starting PC Health Diagnosis…", "info")
        threading.Thread(target=self._diagnosis_bg, daemon=True).start()

    def _diagnosis_bg(self):
        def progress(frac, label):
            self.after(0, lambda f=frac, l=label: (
                self._diag_prog.set(f),
                self._diag_status_lbl.configure(text=l,
                    text_color="gray70" if frac < 1.0 else C_GREEN)))
            if label and label.strip():
                self._diag_log(label, "success" if frac >= 1.0 else "dim")
        findings = self.diag_engine.run_all(progress_cb=progress)
        self.after(0, lambda: self._diagnosis_show_results(findings))
        self.after(0, lambda: self._diag_run_btn.configure(state="normal", text="🔬  Run Diagnosis"))

    def _diagnosis_show_results(self, findings):
        for w in self._diag_scroll.winfo_children(): w.destroy()
        for w in self._diag_chip_frame.winfo_children(): w.destroy()
        counts = {"critical": 0, "warning": 0, "info": 0, "ok": 0}
        for f in findings: counts[f.get("status", "ok")] += 1
        for status, color, icon in [("critical", C_RED, "X"), ("warning", C_ORANGE, "!"),
                                     ("info", C_ACCENT, "i"), ("ok", C_GREEN, "OK")]:
            n = counts[status]
            if not n: continue
            chip = ctk.CTkFrame(self._diag_chip_frame, fg_color=C_CARD2, corner_radius=8)
            chip.pack(side="left", padx=3)
            ctk.CTkLabel(chip, text=f" {icon} {n} {status} ", font=FM(11), text_color=color).pack(padx=6, pady=3)
        self._diag_log(f"Scan complete — {counts['critical']} critical / {counts['warning']} warning / {counts['ok']} ok",
                       "success" if counts["critical"] == 0 else "warn")
        if not findings:
            ctk.CTkLabel(self._diag_scroll, text="No issues detected — your PC looks healthy!",
                         font=FL(14), text_color=C_GREEN).grid(row=0, column=0, pady=40)
            return
        for i, f in enumerate(findings): self._diag_card(self._diag_scroll, i, f)

    def _diag_card(self, parent, row, f):
        status = f.get("status", "ok")
        color, _ = self._DIAG_STATUS.get(status, ("gray60", "?"))
        card = ctk.CTkFrame(parent, corner_radius=10, fg_color=C_CARD,
                            border_color=C_BORDER, border_width=1)
        card.grid(row=row, column=0, sticky="ew", padx=4, pady=4)
        card.grid_columnconfigure(1, weight=1)
        ctk.CTkFrame(card, width=5, corner_radius=0, fg_color=color).grid(
            row=0, column=0, rowspan=3, sticky="ns", padx=(0, 10))
        hdr = ctk.CTkFrame(card, fg_color="transparent")
        hdr.grid(row=0, column=1, sticky="ew", padx=(0, 12), pady=(10, 2))
        hdr.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(hdr, text=f"{f.get('icon','•')}  {f.get('title','')}",
                     font=FH(14), text_color=color, anchor="w").grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(hdr, text=f.get("category", ""), font=FM(11),
                     text_color="gray50", anchor="e").grid(row=0, column=1, sticky="e", padx=(0, 4))
        detail = f.get("detail", "")
        if detail:
            dl = ctk.CTkLabel(card, text=detail, font=F(12, False),
                              text_color="gray75", anchor="w", wraplength=700, justify="left")
            dl.grid(row=1, column=1, sticky="ew", padx=(0, 16), pady=(2, 6))
            def _uw(event, lbl=dl): lbl.configure(wraplength=max(200, event.width - 80))
            card.bind("<Configure>", _uw)
        if status != "ok" and f.get("action_label"):
            bc = {"critical": ("#b71c1c", "#c62828"), "warning": ("#e65100", "#f4511e"),
                  "info": ("#1565c0", "#1976d2")}.get(status, ("#1565c0", "#1976d2"))
            ctk.CTkButton(card, text=f.get("action_label"), font=FB(12), height=34,
                          corner_radius=7, fg_color=bc[0], hover_color=bc[1],
                          command=lambda fnd=f: self._exec_diag_action(fnd)).grid(
                row=2, column=1, sticky="w", padx=(0, 12), pady=(0, 10))

    def _exec_diag_action(self, finding):
        import webbrowser, subprocess as _sp
        atype = finding.get("action_type"); aval = finding.get("action_value", "")
        self._diag_log(f"Action: {finding.get('action_label','')} — {finding.get('title','')}", "info")
        if atype == "url": webbrowser.open(aval)
        elif atype == "store":
            try: _sp.Popen(["cmd", "/c", f"start {aval}"], creationflags=0x08000000)
            except Exception: webbrowser.open(aval)
        elif atype == "command":
            try: _sp.Popen(["cmd", "/c", aval], creationflags=0x08000000)
            except Exception as e: messagebox.showerror("Error", str(e))
        elif atype == "winget":
            self._switch_to_updates()
            self._upd_log(f"Installing {aval}…", "info")
            def _bg():
                r = self.upd_manager.install_winget_update(aval,
                    cb=lambda l, t="": self._upd_log(f"  {l}", t) if l else None)
                tag = "success" if r.get("success") else "warn"
                self._upd_log(f"{'Done' if r.get('success') else 'Failed'}: {aval}", tag)
            threading.Thread(target=_bg, daemon=True).start()

    # ══════════════════════════════════════════════════════════════════════════
    # REPAIR ACTIONS
    # ══════════════════════════════════════════════════════════════════════════
    def _repair_log(self, text):
        def _do():
            self._repair_box.insert("end", text.rstrip("\n") + "\n")
            self._repair_box.see("end")
        self.after(0, _do)

    def _run_sfc(self):
        self._repair_log("\n" + "="*56 + "\n  SFC /scannow — may take 5-15 min…\n" + "="*56 + "\n")
        r = self.upd_manager.run_sfc_scan(cb=lambda l, *_: self._repair_log(f"  {l}") if l else None)
        self._repair_log(f"\n{'OK' if r.get('success') else 'Warning'}: SFC complete.\n")
        self._set_status("SFC done")

    def _run_dism_scan(self):
        self._repair_log("\n" + "="*56 + "\n  DISM /ScanHealth…\n" + "="*56 + "\n")
        r = self.upd_manager.run_dism_scan_health(cb=lambda l, *_: self._repair_log(f"  {l}") if l else None)
        self._repair_log(f"\n{'OK' if r.get('success') else 'Warning'}: DISM Scan complete.\n")

    def _run_dism_restore(self):
        self._repair_log("\n" + "="*56 + "\n  DISM /RestoreHealth — may take 15-45 min…\n" + "="*56 + "\n")
        r = self.upd_manager.run_dism_restore_health(cb=lambda l, *_: self._repair_log(f"  {l}") if l else None)
        self._repair_log(f"\n{'OK' if r.get('success') else 'Warning'}: DISM Restore complete.\n")

    def _run_full_repair(self):
        self._repair_log("\n" + "="*56 + "\n  FULL REPAIR: SFC → DISM → SFC\n" + "="*56)
        self._repair_log("\n  Step 1/3: SFC /scannow\n")
        self.upd_manager.run_sfc_scan(cb=lambda l, *_: self._repair_log(f"  {l}") if l else None)
        self._repair_log("\n  Step 2/3: DISM /RestoreHealth\n")
        self.upd_manager.run_dism_restore_health(cb=lambda l, *_: self._repair_log(f"  {l}") if l else None)
        self._repair_log("\n  Step 3/3: SFC /scannow (verify)\n")
        self.upd_manager.run_sfc_scan(cb=lambda l, *_: self._repair_log(f"  {l}") if l else None)
        self._repair_log("\n" + "="*56 + "\n  Full repair complete — restart your PC.\n" + "="*56 + "\n")
        self._set_status("Full repair done")

    def _clean_temp(self):
        self._repair_log("\n  Cleaning Temp, %TEMP%, Prefetch…\n")
        r = self.upd_manager.clean_temp_folders(cb=lambda l, *_: self._repair_log(l) if l else None)
        self._repair_log(f"\nCleaned {r.get('total_freed_mb',0)} MB\n")
        self._set_status(f"Cleaned {r.get('total_freed_mb',0)} MB")

    def _run_wu_cleanup(self):
        import subprocess
        self._repair_log("\n" + "="*58 + "\n  Windows Update Cleanup\n" +
                         "  DISM /Online /Cleanup-Image /StartComponentCleanup\n" +
                         "  This removes old WU packages — may take 5-20 min.\n" + "="*58 + "\n")
        self._set_status("WU Cleanup running…")
        try:
            proc = subprocess.Popen(
                ["dism", "/Online", "/Cleanup-Image", "/StartComponentCleanup", "/ResetBase"],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, creationflags=0x08000000)
            for line in iter(proc.stdout.readline, ""):
                l = line.strip()
                if l: self._repair_log(f"  {l}")
            proc.wait()
            if proc.returncode == 0:
                self._repair_log("\n  ✓ WU Cleanup complete — restart recommended.\n")
                self._set_status("WU Cleanup done — restart PC")
            else:
                self._repair_log(f"\n  Finished (rc={proc.returncode}). Run as Administrator if it failed.\n")
                self._set_status("WU Cleanup done")
        except Exception as exc:
            self._repair_log(f"\n  Error: {exc}\n"); self._set_status("WU Cleanup error")

    def _clean_ram(self):
        import ctypes, psutil as _ps
        before = _ps.virtual_memory()
        self._repair_log("\n" + "="*60 + "\n  RAM Cleaner / Memory Optimiser\n" + "="*60 + "\n")
        self._repair_log(f"  Before:  {round(before.used/(1024**3),2)} GB used  ({before.percent}%)  —  "
                         f"{round(before.available/(1024**3),2)} GB free\n")
        self._repair_log("  Flushing standby + process working sets…\n"); self._set_status("Cleaning RAM…")
        flushed = 0; errors = 0
        try:
            kernel32 = ctypes.windll.kernel32; psapi = ctypes.windll.psapi
            PROCESS_ALL_ACCESS = 0x1F0FFF
            for proc in _ps.process_iter(["pid", "name"]):
                try:
                    pid = proc.info["pid"]
                    if pid in (0, 4): continue
                    h = kernel32.OpenProcess(PROCESS_ALL_ACCESS, False, pid)
                    if h: psapi.EmptyWorkingSet(h); kernel32.CloseHandle(h); flushed += 1
                except Exception: errors += 1; continue
            try:
                MEM_EXTENDED_PARAMETER = ctypes.c_ulong(0)
                ctypes.windll.ntdll.NtSetSystemInformation(0x4C, ctypes.byref(MEM_EXTENDED_PARAMETER), 4)
            except Exception: pass
        except Exception as outer_exc:
            self._repair_log(f"  Error: {outer_exc}\n"); self._set_status("RAM Cleaner error"); return
        after = _ps.virtual_memory()
        freed_gb = round((before.used - after.used) / (1024**3), 2)
        freed_mb = round((before.used - after.used) / (1024**2))
        self._repair_log(f"  After:   {round(after.used/(1024**3),2)} GB used  ({after.percent}%)  —  "
                         f"{round(after.available/(1024**3),2)} GB free\n")
        self._repair_log(f"  Processes flushed: {flushed}  |  Skipped: {errors}\n")
        if freed_mb > 0:
            self._repair_log(f"  ✓ Freed ~{freed_gb} GB  ({freed_mb} MB)\n")
            self._set_status(f"RAM cleaned — {freed_gb} GB freed")
        else:
            self._repair_log("  ✓ Complete — Windows had already reclaimed standby memory.\n"
                             "  (Run as Administrator for deeper standby-list purge)\n")
            self._set_status("RAM clean — working sets flushed")

    # ══════════════════════════════════════════════════════════════════════════
    # OPTIMISE ACTIONS
    # ══════════════════════════════════════════════════════════════════════════
    def _opt_log(self, text):
        def _do():
            self._opt_box.insert("end", text.rstrip("\n") + "\n")
            self._opt_box.see("end")
        self.after(0, _do)

    def _apply_gaming(self):
        selected = [(k, l, c) for k, l, c in self.GAMING_TWEAKS
                    if self._game_vars.get(k, tk.BooleanVar()).get()]
        if not selected: messagebox.showinfo("Nothing Selected", "Tick at least one gaming tweak."); return
        self._opt_log(f"Applying {len(selected)} gaming tweak(s)…\n")
        import subprocess
        for key, label, cmd in selected:
            self._opt_log(f"  > {label}")
            try:
                r = subprocess.run(["cmd", "/c", cmd], capture_output=True, text=True,
                    creationflags=0x08000000, timeout=30)
                self._opt_log(f"    {'Done' if r.returncode == 0 else 'rc=' + str(r.returncode)}\n")
            except Exception as e: self._opt_log(f"    Error: {e}\n")
        self._opt_log("Gaming tweaks applied — restart to activate all.\n")
        self._set_status("Gaming tweaks applied")

    def _confirm_debloat(self):
        n = sum(1 for k, _, _ in self.DEBLOAT_ITEMS if self._bloat_vars.get(k, tk.BooleanVar()).get())
        if n == 0: messagebox.showinfo("Nothing Selected", "Tick at least one item."); return
        if messagebox.askyesno("Confirm Removal",
                f"Remove/disable {n} selected item(s)?\nA restart is recommended. Continue?"):
            threading.Thread(target=self._remove_selected_bloat, daemon=True).start()

    def _remove_selected_bloat(self):
        import subprocess
        CNW = 0x08000000
        def _ps(sc): subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-NonInteractive",
                                       "-Command", sc], capture_output=True, creationflags=CNW, timeout=60)
        def _reg(c): subprocess.run(["cmd", "/c", c], capture_output=True, creationflags=CNW, timeout=20)
        self._opt_log("Removing selected items…\n")
        for key, label, pkg in self.DEBLOAT_ITEMS:
            if not self._bloat_vars.get(key, tk.BooleanVar()).get(): continue
            self._opt_log(f"  > {label}")
            try:
                if pkg: _ps(f'Get-AppxPackage -AllUsers "*{pkg}*" | Remove-AppxPackage -ErrorAction SilentlyContinue')
                if key == "copilot_btn":
                    _reg(r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v ShowCopilotButton /t REG_DWORD /d 0 /f')
                elif key == "recall":
                    _reg(r'reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" /v DisableAIDataAnalysis /t REG_DWORD /d 1 /f')
                elif key == "chat_icon":
                    _reg(r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v TaskbarMn /t REG_DWORD /d 0 /f')
                elif key == "widgets":
                    _reg(r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v TaskbarDa /t REG_DWORD /d 0 /f')
                elif key == "ads":
                    _reg(r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" /v Enabled /t REG_DWORD /d 0 /f')
                elif key == "telemetry":
                    _reg(r'reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection" /v AllowTelemetry /t REG_DWORD /d 0 /f')
                    subprocess.run(["cmd", "/c", "sc config DiagTrack start= disabled & net stop DiagTrack"],
                        capture_output=True, creationflags=CNW, timeout=20)
                elif key == "onedrive_run":
                    _reg(r'reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v OneDrive /f')
                elif key == "start_ads":
                    _reg(r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v SystemPaneSuggestionsEnabled /t REG_DWORD /d 0 /f')
                self._opt_log("    Done\n")
            except Exception as e: self._opt_log(f"    Error: {e}\n")
        self._opt_log("Removal complete — restart your PC.\n"); self._set_status("Debloat done — restart PC")

    # ══════════════════════════════════════════════════════════════════════════
    # DNS ACTIONS
    # ══════════════════════════════════════════════════════════════════════════
    def _dns_status(self, msg, color):
        def _do(m=msg, c=color):
            # Update Network section status label (created in _build_network_section)
            if hasattr(self, "_dns_status_lbl"):
                self._dns_status_lbl.configure(text=m, text_color=c)
            # Update Dashboard DNS strip status label
            if hasattr(self, "_dns_status_lbl_dash"):
                self._dns_status_lbl_dash.configure(text=m, text_color=c)
        self.after(0, _do)

    def _get_active_adapters(self):
        import subprocess
        try:
            r = subprocess.run(["netsh", "interface", "show", "interface"],
                capture_output=True, text=True, creationflags=0x08000000, timeout=10)
            adapters = []
            for line in r.stdout.splitlines():
                line = line.strip()
                if "Connected" in line:
                    parts = line.split()
                    if len(parts) >= 4: adapters.append(" ".join(parts[3:]))
            return adapters
        except Exception: return []

    def _apply_dns(self):
        import subprocess; CNW = 0x08000000
        chosen = self._dns_var.get(); pri, sec = self._DNS_PRESETS.get(chosen, ("1.1.1.1", "1.0.0.1"))
        self._dns_status("Applying…", C_ACCENT)
        adapters = self._get_active_adapters()
        if not adapters: self._dns_status("No connected adapters found — run as Admin", C_ORANGE); return
        ok = 0
        for adapter in adapters:
            try:
                r1 = subprocess.run(["netsh", "interface", "ip", "set", "dns", adapter, "static", pri],
                    capture_output=True, text=True, creationflags=CNW, timeout=10)
                subprocess.run(["netsh", "interface", "ip", "add", "dns", adapter, sec, "index=2"],
                    capture_output=True, text=True, creationflags=CNW, timeout=10)
                if r1.returncode == 0: ok += 1
            except Exception: pass
        if ok:
            label = chosen.split("(")[0].strip()
            self._dns_status(f"✓ {label} applied to {ok} adapter(s)", C_GREEN)
        else:
            self._dns_status("Failed — run PC Health AI as Administrator", C_RED)

    def _reset_dns(self):
        import subprocess; CNW = 0x08000000
        self._dns_status("Resetting to Automatic…", C_ACCENT)
        ok = 0
        for adapter in self._get_active_adapters():
            try:
                r = subprocess.run(["netsh", "interface", "ip", "set", "dns", adapter, "dhcp"],
                    capture_output=True, text=True, creationflags=CNW, timeout=10)
                if r.returncode == 0: ok += 1
            except Exception: pass
        (self._dns_status(f"✓ Reset {ok} adapter(s) to Automatic (DHCP)", C_GREEN) if ok
         else self._dns_status("Failed — run as Administrator", C_RED))

    def _ping_dns(self):
        import subprocess
        chosen = self._dns_var.get(); pri, sec = self._DNS_PRESETS.get(chosen, ("1.1.1.1", "1.0.0.1"))
        self._dns_status("Pinging…", C_ACCENT); results = []
        for ip in (pri, sec):
            try:
                r = subprocess.run(["ping", "-n", "4", "-w", "1000", ip],
                    capture_output=True, text=True, creationflags=0x08000000, timeout=15)
                avg = None
                for line in r.stdout.splitlines():
                    if "Average" in line or "average" in line:
                        try: avg = int(line.split("=")[-1].strip().rstrip("ms").strip())
                        except Exception: pass
                results.append(f"{ip}: {avg}ms" if avg is not None else f"{ip}: timeout")
            except Exception: results.append(f"{ip}: error")
        self._dns_status("  |  ".join(results), C_ACCENT)

    # ══════════════════════════════════════════════════════════════════════════
    # SMART HELPER
    # ══════════════════════════════════════════════════════════════════════════
    def _get_smart_status(self, model: str, serial: str):
        import subprocess, json as _json; CNW = 0x08000000
        try:
            ps = ("Get-WmiObject -Namespace root\\wmi -Class MSStorageDriver_FailurePredictStatus "
                  "-ErrorAction SilentlyContinue | Select-Object InstanceName,PredictFailure "
                  "| ConvertTo-Json -ErrorAction SilentlyContinue")
            r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy",
                                 "Bypass", "-Command", ps],
                capture_output=True, text=True, creationflags=CNW, timeout=14)
            out = (r.stdout or "").strip()
            if out and out not in ("null", "", "[]"):
                data = _json.loads(out)
                if isinstance(data, dict): data = [data]
                for entry in (data or []):
                    if entry.get("PredictFailure") is True:
                        return ("⚠  Failure Predicted — Back Up Immediately!", C_RED)
                if data: return ("✓  Healthy  (SMART OK)", C_GREEN)
        except Exception: pass
        try:
            r = subprocess.run(["wmic", "diskdrive", "get", "model,status"],
                capture_output=True, text=True, creationflags=CNW, timeout=10)
            model_short = (model or "").strip().lower()[:20]
            for line in r.stdout.splitlines():
                line_l = line.lower()
                if model_short and model_short[:10] in line_l:
                    if "ok" in line_l: return ("✓  OK  (wmic)", C_GREEN)
                    elif "pred fail" in line_l: return ("⚠  Failure Predicted", C_RED)
                    status_word = line.strip().split()[-1] if line.strip() else ""
                    return (f"⚠  {status_word}", C_ORANGE) if status_word else ("? Unknown", C_ORANGE)
            oks = [l for l in r.stdout.splitlines() if "ok" in l.lower() and l.strip()]
            if oks: return ("✓  OK", C_GREEN)
        except Exception: pass
        return ("? Unknown — run as Administrator", C_ORANGE)

    # ══════════════════════════════════════════════════════════════════════════
    # SETTINGS DIALOG
    # ══════════════════════════════════════════════════════════════════════════
    def _open_settings(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Settings"); dlg.geometry("520x420"); dlg.transient(self)
        dlg.grab_set(); dlg.resizable(False, False); dlg.configure(fg_color=C_PANEL)

        # ── Header ────────────────────────────────────────────────────────────
        hdr = ctk.CTkFrame(dlg, fg_color=C_CARD, corner_radius=0)
        hdr.pack(fill="x")
        ctk.CTkFrame(hdr, height=4, corner_radius=0, fg_color=C_ACCENT).pack(fill="x")
        ctk.CTkLabel(hdr, text="⚙  Settings", font=FT(16), text_color=C_ACCENT).pack(pady=12)

        body = ctk.CTkFrame(dlg, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=12)

        # ── Admin status badge ────────────────────────────────────────────────
        admin_txt = ("✓  Running as Administrator" if self._is_admin
                     else "✗  NOT running as Administrator — some features require Admin rights")
        admin_col = C_GREEN if self._is_admin else C_ORANGE
        badge = ctk.CTkFrame(body, fg_color=C_CARD2, corner_radius=8,
                             border_color=admin_col, border_width=1)
        badge.pack(fill="x", pady=(0, 14))
        ctk.CTkLabel(badge, text=admin_txt, font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=admin_col, wraplength=440).pack(padx=16, pady=10)

        # ── Anthropic API Key section ──────────────────────────────────────────
        api_frame = ctk.CTkFrame(body, fg_color=C_CARD, corner_radius=10,
                                 border_color=C_BORDER, border_width=1)
        api_frame.pack(fill="x", pady=(0, 14))
        api_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkFrame(api_frame, height=3, corner_radius=0, fg_color=C_ACCENT).grid(
            row=0, column=0, columnspan=3, sticky="ew")
        ctk.CTkLabel(api_frame, text="🤖  Anthropic API Key",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=C_ACCENT, anchor="w").grid(
            row=1, column=0, columnspan=3, padx=14, pady=(10, 4), sticky="w")
        ctk.CTkLabel(api_frame,
                     text="Required for the AI assistant. Get yours free at console.anthropic.com",
                     font=ctk.CTkFont(size=11), text_color="gray45", anchor="w").grid(
            row=2, column=0, columnspan=3, padx=14, pady=(0, 8), sticky="w")

        # Load existing key
        cfg = self._load_cfg()
        existing_key = cfg.get("api_key", "")

        show_var = tk.BooleanVar(value=False)
        key_var  = tk.StringVar(value=existing_key)

        key_entry = ctk.CTkEntry(api_frame, textvariable=key_var,
                                  show="•", width=320,
                                  font=ctk.CTkFont(family="Consolas", size=12),
                                  fg_color=C_CARD2, border_color=C_BORDER,
                                  placeholder_text="sk-ant-api03-…")
        key_entry.grid(row=3, column=0, padx=14, pady=(0, 4), sticky="ew", columnspan=2)

        def _toggle_show():
            key_entry.configure(show="" if show_var.get() else "•")
        ctk.CTkCheckBox(api_frame, text="Show key", variable=show_var,
                        font=ctk.CTkFont(size=11), text_color="gray55",
                        command=_toggle_show).grid(row=3, column=2, padx=(4, 14), pady=(0, 4))

        save_lbl = ctk.CTkLabel(api_frame, text="", font=ctk.CTkFont(size=11, weight="bold"),
                                 text_color=C_GREEN, anchor="w")
        save_lbl.grid(row=4, column=0, columnspan=2, padx=14, pady=(0, 10), sticky="w")

        def _save_key():
            raw = key_var.get().strip()
            if raw and not raw.startswith("sk-ant-"):
                save_lbl.configure(text="⚠  Key should start with 'sk-ant-'", text_color=C_ORANGE)
                return
            self._save_cfg({"api_key": raw})
            save_lbl.configure(text="✓  Saved successfully!", text_color=C_GREEN)

        ctk.CTkButton(api_frame, text="💾  Save Key", font=FB(12), height=36,
                      fg_color="#0a5abf", hover_color="#1269d3",
                      corner_radius=8, command=_save_key).grid(
            row=4, column=2, padx=(4, 14), pady=(0, 10), sticky="e")

        # ── Close ─────────────────────────────────────────────────────────────
        ctk.CTkButton(body, text="Close", font=FB(13), command=dlg.destroy,
                      height=40, width=130, fg_color=C_CARD2, hover_color=C_BORDER,
                      corner_radius=10, border_width=1, border_color=C_BORDER).pack(pady=(0, 4))


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = PCHealthApp()
    app.mainloop()
