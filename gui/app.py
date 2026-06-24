"""OST -> PST Donusturucu - Tkinter masaustu arayuzu.

Her sey net, okunabilir ve Turkce. Donusum islemleri ayri bir is parcaciginda
calisir; arayuz hicbir zaman donmaz. Ilerleme ve gunluk mesajlari bir kuyruk
araciligi ile guvenli sekilde ana parcaciga aktarilir.

Ana akis:
    1. Kullanici ``.ost`` dosyasini ELLE secer.
    2. Dosyanin klasor/mesaj agaci OTOMATIK gosterilir; kullanici onay
       kutulariyla PST'ye gecirilecek klasor ve/veya mesajlari secer.
    3. Hedef olarak bir KLASOR secilir; cikti dosyasi (PST / EML / MBOX) o
       klasorde, kaynak OST adina gore otomatik olusturulur.

Cikan PST, OST'deki klasor yapisini BIREBIR korur (klasor adlari yan dosya ile
tasinir, kok sarmalayicilar duzlestirilir).
"""

from __future__ import annotations

import functools
import os
import queue
import re
import threading
import time
import tkinter as tk
from collections import defaultdict
from tkinter import filedialog, messagebox, ttk

from converter import orchestrator as core

APP_TITLE = "OST → PST Donusturucu"
PAD = 10
_INVALID = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# Onay kutusu sembolleri (secili / secili degil / kismi).
CHK_ON = "☑"
CHK_OFF = "☐"
CHK_PARTIAL = "▣"


def _safe_name(name: str, fallback: str = "cikti") -> str:
    name = _INVALID.sub("_", (name or "").strip()) or fallback
    return name[:80]


def _unique_path(path: str) -> str:
    """Dosya/klasor zaten varsa sonuna sayi ekleyerek benzersiz yol uretir."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    for i in range(1, 1000):
        cand = f"{base}_{i}{ext}"
        if not os.path.exists(cand):
            return cand
    return f"{base}_{int(time.time())}{ext}"


class App(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=PAD)
        self.master = master
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self._queue: "queue.Queue[tuple]" = queue.Queue()
        self._busy = False
        self._stores: list = []

        # --- OST secim agaci durumu ---
        self._current_ost: str = ""          # yuklu OST yolu
        self._nodes: dict = {}               # fid -> dugum (id/parent/name/count)
        self._children: dict = defaultdict(list)  # parent fid (veya None) -> [fid]
        self._folder_state: dict = {}        # fid -> bool (klasor isaretli mi)
        self._msg_override: dict = {}        # fid -> set(index) (mesaj bazli secim)
        self._loaded: set = set()            # mesajlari yuklenmis fid'ler
        self._msg_data: dict = {}            # fid -> {index: {subject,sender,date}}
        self._tree_token = 0                 # eski yuklemeleri gecersiz kil

        self._build_style()
        self._build_header()
        self._build_tabs()
        self._build_progress_and_log()
        self._build_statusbar()

        self.after(100, self._drain_queue)

    # ------------------------------------------------------------------ #
    # Gorunum
    # ------------------------------------------------------------------ #
    def _build_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        base = ("Segoe UI", 10)
        style.configure(".", font=base)
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Sub.TLabel", font=("Segoe UI", 10), foreground="#555")
        style.configure("Card.TLabelframe", padding=PAD)
        style.configure("Go.TButton", font=("Segoe UI", 11, "bold"), padding=8)
        style.configure("TButton", padding=5)
        style.configure("TEntry", padding=4)
        style.configure("Tree.Treeview", rowheight=22)

    def _build_header(self) -> None:
        head = ttk.Frame(self)
        head.grid(row=0, column=0, sticky="ew")
        head.columnconfigure(0, weight=1)
        ttk.Label(head, text=APP_TITLE, style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            head,
            text="OST dosyanizi secin; icerigi asagida gosterilir. Aktarmak "
            "istediginiz klasor/mesajlari isaretleyip Donustur'e basin.",
            style="Sub.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, PAD))

    def _build_tabs(self) -> None:
        nb = ttk.Notebook(self)
        nb.grid(row=1, column=0, sticky="nsew")
        self.rowconfigure(1, weight=1)
        # Ana akis once: OST dosyasini elle sec.
        self._build_tab_file(nb)
        self._build_tab_mailbox(nb)

    # ---- Sekme 1: OST dosyasi (elle sec) -> PST / EML / MBOX ----------- #
    def _build_tab_file(self, nb: ttk.Notebook) -> None:
        tab = ttk.Frame(nb, padding=PAD)
        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(6, weight=1)
        nb.add(tab, text="  OST Dosyasi → PST  ")

        ttk.Label(
            tab,
            text="Diskteki bir .ost dosyasini elle secin. PST cikti icin Outlook "
            "gerekir;\nOutlook yoksa EML veya MBOX olarak kurtarabilirsiniz.",
            justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, PAD))

        ttk.Label(tab, text="Kaynak OST dosyasi:").grid(row=1, column=0, sticky="w")
        self.var_src_ost = tk.StringVar()
        ttk.Entry(tab, textvariable=self.var_src_ost).grid(
            row=1, column=1, sticky="ew", padx=(PAD, 0)
        )
        ttk.Button(tab, text="Dosya sec...", command=self._pick_ost).grid(
            row=1, column=2, sticky="w", padx=(PAD, 0)
        )

        ttk.Label(tab, text="Cikti bicimi:").grid(
            row=2, column=0, sticky="w", pady=(PAD, 0)
        )
        self.var_fmt = tk.StringVar(value="pst")
        fmt_frame = ttk.Frame(tab)
        fmt_frame.grid(row=2, column=1, columnspan=2, sticky="w", pady=(PAD, 0))
        for text, val in [
            ("PST (Outlook gerekli)", "pst"),
            ("EML klasoru", "eml"),
            ("MBOX dosyasi", "mbox"),
        ]:
            ttk.Radiobutton(
                fmt_frame, text=text, value=val, variable=self.var_fmt,
                command=self._on_fmt_change,
            ).pack(side="left", padx=(0, PAD))

        ttk.Label(tab, text="Hedef klasor:").grid(
            row=3, column=0, sticky="w", pady=(PAD, 0)
        )
        self.var_dst_dir = tk.StringVar()
        ttk.Entry(tab, textvariable=self.var_dst_dir).grid(
            row=3, column=1, sticky="ew", padx=(PAD, 0), pady=(PAD, 0)
        )
        ttk.Button(tab, text="Klasor sec...", command=self._pick_dst_dir).grid(
            row=3, column=2, sticky="w", padx=(PAD, 0), pady=(PAD, 0)
        )

        ttk.Label(
            tab,
            text="Cikti dosyasi, kaynak OST adina gore bu klasorde otomatik "
            "olusturulur.",
            style="Sub.TLabel",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(4, 0))

        # ---- Icerik secim agaci ---- #
        sel_head = ttk.Frame(tab)
        sel_head.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(PAD, 2))
        sel_head.columnconfigure(0, weight=1)
        self.var_tree_info = tk.StringVar(
            value="OST secince icerik burada gosterilir (yalnizca mail iceren "
            "klasorler). Satira tiklayarak isaretleyin; klasoru acinca mesajlari "
            "tek tek de secebilirsiniz."
        )
        ttk.Label(sel_head, textvariable=self.var_tree_info,
                  style="Sub.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(sel_head, text="Tumunu sec",
                   command=lambda: self._select_all(True)).grid(
            row=0, column=1, sticky="e", padx=(PAD, 0))
        ttk.Button(sel_head, text="Temizle",
                   command=lambda: self._select_all(False)).grid(
            row=0, column=2, sticky="e", padx=(4, 0))
        ttk.Button(sel_head, text="Yenile",
                   command=self._load_ost_tree).grid(
            row=0, column=3, sticky="e", padx=(4, 0))

        tree_wrap = ttk.Frame(tab)
        tree_wrap.grid(row=6, column=0, columnspan=3, sticky="nsew")
        tree_wrap.columnconfigure(0, weight=1)
        tree_wrap.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_wrap, show="tree", selectmode="none",
                                 style="Tree.Treeview", height=10)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tsb = ttk.Scrollbar(tree_wrap, command=self.tree.yview)
        tsb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=tsb.set)
        self.tree.bind("<Button-1>", self._on_tree_click, add="+")
        self.tree.bind("<<TreeviewOpen>>", self._on_tree_open, add="+")

        self.btn_file_go = ttk.Button(
            tab, text="Donustur", style="Go.TButton", command=self._run_file
        )
        self.btn_file_go.grid(row=7, column=0, columnspan=3, sticky="e", pady=(PAD, 0))

    # ---- Sekme 2: Bagli posta kutusu -> PST  (Outlook hesabi varsa) ---- #
    def _build_tab_mailbox(self, nb: ttk.Notebook) -> None:
        tab = ttk.Frame(nb, padding=PAD)
        tab.columnconfigure(1, weight=1)
        nb.add(tab, text="  Posta Kutusu → PST (Outlook hesabi)  ")

        ttk.Label(
            tab,
            text="Outlook'ta TANIMLI bir hesabin OST onbellegini tam sadakatle "
            "PST'ye kopyalar.\nHesap listesi bossa, Outlook'ta yapilandirilmis "
            "hesap yok demektir; bu durumda 1. sekmeyi kullanin.",
            justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, PAD))

        ttk.Button(
            tab, text="Posta kutularini listele", command=self._load_stores
        ).grid(row=1, column=0, sticky="w")
        self.cmb_store = ttk.Combobox(tab, state="readonly", width=50)
        self.cmb_store.grid(row=1, column=1, columnspan=2, sticky="ew", padx=(PAD, 0))

        ttk.Label(tab, text="Hedef klasor:").grid(
            row=2, column=0, sticky="w", pady=(PAD, 0)
        )
        self.var_mbox_dir = tk.StringVar()
        ttk.Entry(tab, textvariable=self.var_mbox_dir).grid(
            row=2, column=1, sticky="ew", padx=(PAD, 0), pady=(PAD, 0)
        )
        ttk.Button(
            tab, text="Klasor sec...", command=lambda: self._pick_dir(self.var_mbox_dir)
        ).grid(row=2, column=2, sticky="w", padx=(PAD, 0), pady=(PAD, 0))

        self.btn_mbox_go = ttk.Button(
            tab, text="PST'ye Donustur", style="Go.TButton", command=self._run_mailbox
        )
        self.btn_mbox_go.grid(row=3, column=0, columnspan=3, sticky="e", pady=(PAD, 0))

    def _build_progress_and_log(self) -> None:
        frame = ttk.LabelFrame(self, text="Durum", style="Card.TLabelframe")
        frame.grid(row=2, column=0, sticky="nsew", pady=(PAD, 0))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)
        self.rowconfigure(2, weight=1)

        self.progress = ttk.Progressbar(frame, mode="determinate", maximum=1.0)
        self.progress.grid(row=0, column=0, sticky="ew")
        self.var_status = tk.StringVar(value="Hazir.")
        ttk.Label(frame, textvariable=self.var_status).grid(
            row=1, column=0, sticky="w", pady=(4, 4)
        )

        log_wrap = ttk.Frame(frame)
        log_wrap.grid(row=2, column=0, sticky="nsew")
        log_wrap.columnconfigure(0, weight=1)
        log_wrap.rowconfigure(0, weight=1)
        self.log = tk.Text(log_wrap, height=8, wrap="word", state="disabled",
                           font=("Consolas", 9), background="#1e1e1e",
                           foreground="#d4d4d4", insertbackground="#d4d4d4")
        self.log.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(log_wrap, command=self.log.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=sb.set)

    def _build_statusbar(self) -> None:
        eng = core.available_engines()
        ok, no = "✓", "✗"
        text = (
            f"Outlook (COM): {ok if eng['outlook'] else no}     "
            f"libpff (OST okuma): {ok if eng['libpff'] else no}"
        )
        bar = ttk.Frame(self)
        bar.grid(row=3, column=0, sticky="ew", pady=(PAD, 0))
        ttk.Label(bar, text=text, style="Sub.TLabel").grid(row=0, column=0, sticky="w")
        if not eng["outlook"] and not eng["libpff"]:
            self._log(
                "UYARI: Hicbir donusum motoru bulunamadi. "
                "Lutfen README'deki kurulum adimlarini izleyin."
            )
        self._on_fmt_change()

    # ------------------------------------------------------------------ #
    # Dosya / klasor secicileri
    # ------------------------------------------------------------------ #
    def _pick_ost(self) -> None:
        path = filedialog.askopenfilename(
            title="OST dosyasini secin",
            filetypes=[("Outlook OST", "*.ost"), ("Tum dosyalar", "*.*")],
        )
        if path:
            self.var_src_ost.set(path)
            self._load_ost_tree()

    def _pick_dir(self, var: tk.StringVar) -> None:
        path = filedialog.askdirectory(title="Hedef klasoru secin")
        if path:
            var.set(path)

    def _pick_dst_dir(self) -> None:
        self._pick_dir(self.var_dst_dir)

    def _on_fmt_change(self) -> None:
        eng = core.available_engines()
        fmt = self.var_fmt.get()
        if fmt == "pst" and not eng["outlook"]:
            self.var_status.set("Not: PST cikti icin Outlook gerekli; "
                                "yoksa EML/MBOX secin.")
        else:
            self.var_status.set("Hazir.")

    # ------------------------------------------------------------------ #
    # OST icerik agaci
    # ------------------------------------------------------------------ #
    def _clear_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self._nodes = {}
        self._children = defaultdict(list)
        self._folder_state = {}
        self._msg_override = {}
        self._loaded = set()
        self._msg_data = {}

    def _load_ost_tree(self) -> None:
        """Secili OST'nin klasor agacini (ayri parcacikta) yukler."""
        src = self.var_src_ost.get().strip()
        self._clear_tree()
        if not src or not os.path.exists(src):
            self.var_tree_info.set("Gecerli bir OST dosyasi secin.")
            return
        if not core.available_engines()["libpff"]:
            self.var_tree_info.set(
                "Icerik onizleme icin libpff gerekli ('pip install libpff-python'). "
                "Yine de tum dosyayi donusturebilirsiniz."
            )
            return
        self._current_ost = os.path.normpath(src)
        self.var_tree_info.set("Icerik okunuyor...")
        self._tree_token += 1
        token = self._tree_token

        def work() -> None:
            try:
                nodes = core.list_ost_tree(self._current_ost)
                self._queue.put(("tree", token, nodes))
            except Exception as exc:
                self._queue.put(("tree_err", token, str(exc)))

        threading.Thread(target=work, daemon=True).start()

    def _populate_tree(self, nodes: list) -> None:
        self._clear_tree()
        for n in nodes:
            fid = n["id"]
            self._nodes[fid] = n
            self._folder_state[fid] = True  # varsayilan: hepsi secili
            self._children[n["parent"]].append(fid)
            parent_iid = ("F:" + n["parent"]) if n["parent"] else ""
            try:
                self.tree.insert(parent_iid, "end", iid="F:" + fid,
                                 text=self._folder_text(fid), open=False)
            except tk.TclError:
                continue
            if n["count"] > 0:
                # Tembel yukleme icin gecici dugum.
                self.tree.insert("F:" + fid, "end", iid="D:" + fid,
                                 text="  (mesajlari gormek icin acin)")
        total_msgs = sum(n["count"] for n in nodes)
        self.var_tree_info.set(
            f"{len(nodes)} klasor, {total_msgs} mesaj. Aktarilacaklari isaretleyin "
            "(varsayilan: hepsi secili)."
        )

    # ---- onay kutusu / metin ---- #
    def _folder_glyph(self, fid: str) -> str:
        ov = self._msg_override.get(fid)
        count = self._nodes[fid]["count"]
        if ov is not None and 0 < len(ov) < count:
            return CHK_PARTIAL
        return CHK_ON if self._folder_state.get(fid) else CHK_OFF

    def _folder_text(self, fid: str) -> str:
        n = self._nodes[fid]
        suffix = f"   [{n['count']}]" if n["count"] else ""
        return f"{self._folder_glyph(fid)}  {n['name']}{suffix}"

    def _msg_included(self, fid: str, idx: int) -> bool:
        ov = self._msg_override.get(fid)
        if ov is None:
            return bool(self._folder_state.get(fid))
        return idx in ov

    def _msg_text(self, fid: str, idx: int) -> str:
        m = self._msg_data.get(fid, {}).get(idx, {})
        glyph = CHK_ON if self._msg_included(fid, idx) else CHK_OFF
        subject = m.get("subject", "(konusuz)")
        sender = m.get("sender", "")
        date = m.get("date", "")
        meta = "  —  ".join(x for x in (sender, date) if x)
        tail = f"   ({meta})" if meta else ""
        return f"{glyph}  {subject}{tail}"

    def _update_folder_row(self, fid: str) -> None:
        if self.tree.exists("F:" + fid):
            self.tree.item("F:" + fid, text=self._folder_text(fid))

    def _update_loaded_msgs(self, fid: str) -> None:
        if fid not in self._loaded:
            return
        for miid in self.tree.get_children("F:" + fid):
            if miid.startswith("M:"):
                idx = int(miid.rsplit("|", 1)[1])
                self.tree.item(miid, text=self._msg_text(fid, idx))

    def _norm_override(self, fid: str) -> None:
        """Mesaj secimi tamamen dolu/bos ise klasor durumuna sadelestirir."""
        ov = self._msg_override.get(fid)
        if ov is None:
            return
        count = self._nodes[fid]["count"]
        if len(ov) == 0:
            self._msg_override.pop(fid, None)
            self._folder_state[fid] = False
        elif len(ov) >= count:
            self._msg_override.pop(fid, None)
            self._folder_state[fid] = True

    def _set_folder_checked(self, fid: str, val: bool, cascade: bool = True) -> None:
        self._folder_state[fid] = val
        self._msg_override.pop(fid, None)
        self._update_folder_row(fid)
        self._update_loaded_msgs(fid)
        if cascade:
            for child in self._children.get(fid, []):
                self._set_folder_checked(child, val, True)

    def _toggle_msg(self, fid: str, idx: int) -> None:
        count = self._nodes[fid]["count"]
        ov = self._msg_override.get(fid)
        if ov is None:
            ov = set(range(count)) if self._folder_state.get(fid) else set()
        else:
            ov = set(ov)
        if idx in ov:
            ov.discard(idx)
        else:
            ov.add(idx)
        self._msg_override[fid] = ov
        self._norm_override(fid)
        self._update_folder_row(fid)
        self._update_loaded_msgs(fid)

    def _select_all(self, val: bool) -> None:
        for fid in self._nodes:
            self._folder_state[fid] = val
            self._msg_override.pop(fid, None)
            self._update_folder_row(fid)
            self._update_loaded_msgs(fid)

    def _on_tree_click(self, event) -> None:
        # Acma/kapama ucgenine tiklandiysa varsayilan davranisi birak.
        elem = self.tree.identify_element(event.x, event.y)
        if "indicator" in (elem or ""):
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if iid.startswith("F:"):
            fid = iid[2:]
            self._set_folder_checked(fid, not self._folder_state.get(fid, True))
        elif iid.startswith("M:"):
            rest = iid[2:]
            fid, sidx = rest.rsplit("|", 1)
            self._toggle_msg(fid, int(sidx))

    def _on_tree_open(self, event) -> None:
        iid = self.tree.focus()
        if not iid.startswith("F:"):
            return
        fid = iid[2:]
        if fid in self._loaded:
            return
        if not self.tree.exists("D:" + fid):
            return
        self.tree.item("D:" + fid, text="  (mesajlar yukleniyor...)")
        token = self._tree_token

        def work() -> None:
            try:
                msgs = core.list_ost_folder_messages(self._current_ost, fid)
                self._queue.put(("msgs", token, fid, msgs))
            except Exception as exc:
                self._queue.put(("msgs_err", token, fid, str(exc)))

        threading.Thread(target=work, daemon=True).start()

    def _on_msgs_loaded(self, fid: str, msgs: list) -> None:
        if self.tree.exists("D:" + fid):
            self.tree.delete("D:" + fid)
        self._msg_data[fid] = {m["index"]: m for m in msgs}
        for m in msgs:
            idx = m["index"]
            miid = "M:%s|%d" % (fid, idx)
            if self.tree.exists(miid):
                continue
            self.tree.insert("F:" + fid, "end", iid=miid,
                             text=self._msg_text(fid, idx))
        self._loaded.add(fid)

    def _gather_selection(self) -> dict:
        """Arayuz durumundan motor icin secim sozlugu uretir."""
        sel: dict = {}
        for fid in self._nodes:
            ov = self._msg_override.get(fid)
            if ov is not None:
                if ov:
                    sel[fid] = set(ov)
            elif self._folder_state.get(fid):
                sel[fid] = True
        return sel

    # ------------------------------------------------------------------ #
    # Posta kutularini listeleme
    # ------------------------------------------------------------------ #
    def _load_stores(self) -> None:
        if not core.available_engines()["outlook"]:
            messagebox.showwarning(
                APP_TITLE,
                "Outlook / pywin32 bulunamadi.\nBu sekme yalnizca Windows + "
                "Outlook ortaminda calisir. Lutfen 1. sekmeyi kullanin.",
            )
            return
        try:
            self._stores = core.list_outlook_stores()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Posta kutulari alinamadi:\n{exc}")
            return
        labels = [str(s) for s in self._stores]
        self.cmb_store["values"] = labels
        if labels:
            self.cmb_store.current(0)
            self._log(f"{len(labels)} posta kutusu bulundu.")
        else:
            self._log("Profilde yapilandirilmis posta kutusu bulunamadi.")
            messagebox.showinfo(
                APP_TITLE,
                "Outlook'ta yapilandirilmis hesap bulunamadi.\n\n"
                "OST dosyanizi dogrudan cevirmek icin 1. sekmeyi "
                "('OST Dosyasi → PST') kullanin.",
            )

    # ------------------------------------------------------------------ #
    # Calistirma
    # ------------------------------------------------------------------ #
    def _run_file(self) -> None:
        if self._busy:
            return
        src = self.var_src_ost.get().strip()
        out_dir = self.var_dst_dir.get().strip()
        if not src or not os.path.exists(src):
            messagebox.showwarning(APP_TITLE, "Gecerli bir OST dosyasi secin.")
            return
        if not out_dir or not os.path.isdir(out_dir):
            messagebox.showwarning(APP_TITLE, "Gecerli bir hedef klasor secin.")
            return

        src = os.path.normpath(src)
        out_dir = os.path.normpath(out_dir)

        # Secim: agac yuklendiyse oradan; yoksa None (tum dosya).
        selection = None
        if self._nodes:
            selection = self._gather_selection()
            if not selection:
                messagebox.showwarning(
                    APP_TITLE,
                    "Hic klasor/mesaj secilmedi.\nAktarmak istediginiz ogeleri "
                    "isaretleyin veya 'Tumunu sec' butonunu kullanin.",
                )
                return

        base = _safe_name(os.path.splitext(os.path.basename(src))[0], "donusum")
        fmt = self.var_fmt.get()
        if fmt == "pst":
            dst = _unique_path(os.path.join(out_dir, base + ".pst"))
            fn = core.file_to_pst
        elif fmt == "mbox":
            dst = _unique_path(os.path.join(out_dir, base + ".mbox"))
            fn = core.file_to_mbox
        else:
            dst = _unique_path(os.path.join(out_dir, base + "_eml"))
            fn = core.file_to_eml

        fn = functools.partial(fn, selection=selection)
        self._log(f"Cikti: {dst}")
        self._start(fn, src, dst)

    def _run_mailbox(self) -> None:
        if self._busy:
            return
        idx = self.cmb_store.current()
        if idx < 0 or idx >= len(self._stores):
            messagebox.showwarning(APP_TITLE, "Lutfen once bir posta kutusu secin.")
            return
        out_dir = self.var_mbox_dir.get().strip()
        if not out_dir or not os.path.isdir(out_dir):
            messagebox.showwarning(APP_TITLE, "Gecerli bir hedef klasor secin.")
            return
        out_dir = os.path.normpath(out_dir)
        store = self._stores[idx]
        base = _safe_name(getattr(store, "name", "posta_kutusu"), "posta_kutusu")
        dst = _unique_path(os.path.join(out_dir, base + ".pst"))
        self._log(f"Cikti: {dst}")
        self._start(core.mailbox_to_pst, store.store_id, dst)

    def _start(self, fn, *args) -> None:
        self._set_busy(True)
        self.progress["value"] = 0
        self.var_status.set("Calisiyor...")

        def progress(msg: str, frac: float) -> None:
            self._queue.put(("progress", msg, frac))

        def worker() -> None:
            try:
                result = fn(*args, progress=progress)
                self._queue.put(("done", result))
            except Exception as exc:  # core.ConversionError dahil
                self._queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------------------ #
    # Kuyruk -> arayuz
    # ------------------------------------------------------------------ #
    def _drain_queue(self) -> None:
        try:
            while True:
                item = self._queue.get_nowait()
                kind = item[0]
                if kind == "progress":
                    _, msg, frac = item
                    self._log(msg)
                    if frac is not None and frac >= 0:
                        self.progress.configure(mode="determinate")
                        self.progress["value"] = frac
                    self.var_status.set(msg)
                elif kind == "tree":
                    _, token, nodes = item
                    if token == self._tree_token:
                        self._populate_tree(nodes)
                elif kind == "tree_err":
                    _, token, msg = item
                    if token == self._tree_token:
                        self.var_tree_info.set("Icerik okunamadi: " + msg)
                elif kind == "msgs":
                    _, token, fid, msgs = item
                    if token == self._tree_token:
                        self._on_msgs_loaded(fid, msgs)
                elif kind == "msgs_err":
                    _, token, fid, msg = item
                    if token == self._tree_token and self.tree.exists("D:" + fid):
                        self.tree.item("D:" + fid,
                                       text="  (mesajlar okunamadi: %s)" % msg)
                elif kind == "done":
                    self._on_done(item[1])
                elif kind == "error":
                    self._on_error(item[1])
        except queue.Empty:
            pass
        self.after(100, self._drain_queue)

    def _on_done(self, result: str) -> None:
        self.progress["value"] = 1.0
        self.var_status.set("Tamamlandi.")
        self._set_busy(False)
        self._log(f"BASARILI -> {result}")
        messagebox.showinfo(
            APP_TITLE, f"Donusum tamamlandi.\n\nCikti:\n{result}"
        )

    def _on_error(self, msg: str) -> None:
        self.var_status.set("Hata.")
        self._set_busy(False)
        self._log("HATA: " + msg)
        messagebox.showerror(APP_TITLE, msg)

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.btn_mbox_go.configure(state=state)
        self.btn_file_go.configure(state=state)

    # ------------------------------------------------------------------ #
    # Gunluk
    # ------------------------------------------------------------------ #
    def _log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        # Gunluk kutusunu sinirla: cok uzun gunluk arayuzu yavaslatir.
        try:
            line_count = int(self.log.index("end-1c").split(".")[0])
            if line_count > 600:
                self.log.delete("1.0", f"{line_count - 600}.0")
        except Exception:
            pass
        self.log.see("end")
        self.log.configure(state="disabled")


def main() -> None:
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("820x760")
    root.minsize(720, 620)
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
