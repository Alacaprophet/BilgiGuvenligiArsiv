"""OST -> PST Donusturucu - Tkinter masaustu arayuzu.

Her sey net, okunabilir ve Turkce. Donusum islemleri ayri bir is parcaciginda
calisir; arayuz hicbir zaman donmaz. Ilerleme ve gunluk mesajlari bir kuyruk
araciligi ile guvenli sekilde ana parcaciga aktarilir.

Ana akis: kullanici .ost dosyasini ELLE secer, hedef olarak yalnizca bir
KLASOR secer; cikti dosyasi (PST / EML / MBOX) o klasorde otomatik olusturulur.
"""

from __future__ import annotations

import os
import queue
import re
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from converter import orchestrator as core

APP_TITLE = "OST → PST Donusturucu"
PAD = 10
_INVALID = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


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

    def _build_header(self) -> None:
        head = ttk.Frame(self)
        head.grid(row=0, column=0, sticky="ew")
        head.columnconfigure(0, weight=1)
        ttk.Label(head, text=APP_TITLE, style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            head,
            text="OST dosyanizi secin, hedef klasoru secin; PST dosyasi otomatik "
            "olusturulur.",
            style="Sub.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, PAD))

    def _build_tabs(self) -> None:
        nb = ttk.Notebook(self)
        nb.grid(row=1, column=0, sticky="nsew")
        self.rowconfigure(1, weight=0)
        # Ana akis once: OST dosyasini elle sec.
        self._build_tab_file(nb)
        self._build_tab_mailbox(nb)

    # ---- Sekme 1: OST dosyasi (elle sec) -> PST / EML / MBOX ----------- #
    def _build_tab_file(self, nb: ttk.Notebook) -> None:
        tab = ttk.Frame(nb, padding=PAD)
        tab.columnconfigure(1, weight=1)
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

        self.btn_file_go = ttk.Button(
            tab, text="Donustur", style="Go.TButton", command=self._run_file
        )
        self.btn_file_go.grid(row=5, column=0, columnspan=3, sticky="e", pady=(PAD, 0))

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
        self.log = tk.Text(log_wrap, height=10, wrap="word", state="disabled",
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
        self.log.see("end")
        self.log.configure(state="disabled")


def main() -> None:
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("780x660")
    root.minsize(680, 580)
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
