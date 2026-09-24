"""brut_ucret için tkinter arayüzü. PyInstaller ile exe'ye çevrilir (bkz. build.bat)."""
import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from brut_ucret import write_excel


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Brüt Ücret Okuyucu")
        self.geometry("720x480")
        self.minsize(560, 380)
        self.pdfs = []
        self.out_path = None

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Button(top, text="PDF Seç...", command=self.pick_files).pack(side="left")
        ttk.Button(top, text="Klasör Seç...", command=self.pick_folder).pack(side="left", padx=6)
        ttk.Button(top, text="Temizle", command=self.clear).pack(side="left")

        ttk.Label(self, text="Seçili PDF'ler:", padding=(10, 0)).pack(anchor="w")
        self.listbox = tk.Listbox(self, height=7)
        self.listbox.pack(fill="x", padx=10)

        run_row = ttk.Frame(self, padding=10)
        run_row.pack(fill="x")
        self.run_btn = ttk.Button(run_row, text="Excel Oluştur", command=self.run)
        self.run_btn.pack(side="left")
        self.open_btn = ttk.Button(run_row, text="Excel'i Aç", command=self.open_excel, state="disabled")
        self.open_btn.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(run_row, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=6)

        self.log_box = tk.Text(self, height=10, state="disabled", wrap="word")
        self.log_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_box.tag_config("ok", foreground="#1a7f37")
        self.log_box.tag_config("err", foreground="#c62828")

        # İşçi thread'i UI'ye dokunmaz; mesajlarını bu kuyruğa koyar, ana thread okur.
        self.events = queue.Queue()
        self.after(100, self.poll_events)

    def set_pdfs(self, pdfs):
        self.pdfs = sorted(set(self.pdfs) | set(pdfs))
        self.listbox.delete(0, "end")
        for p in self.pdfs:
            self.listbox.insert("end", p.name)

    def pick_files(self):
        files = filedialog.askopenfilenames(title="PDF seç", filetypes=[("PDF", "*.pdf")])
        self.set_pdfs(Path(f) for f in files)

    def pick_folder(self):
        folder = filedialog.askdirectory(title="PDF klasörü seç")
        if folder:
            pdfs = list(Path(folder).glob("*.pdf"))
            if not pdfs:
                messagebox.showwarning("PDF yok", "Seçilen klasörde PDF bulunamadı.")
            self.set_pdfs(pdfs)

    def clear(self):
        self.pdfs = []
        self.listbox.delete(0, "end")

    def log(self, msg, tag=None):
        self.events.put(("log", msg, tag))

    def poll_events(self):
        while not self.events.empty():
            kind, *args = self.events.get_nowait()
            if kind == "log":
                msg, tag = args
                self.log_box.config(state="normal")
                self.log_box.insert("end", msg + "\n", tag)
                self.log_box.see("end")
                self.log_box.config(state="disabled")
            else:
                self.done(*args)
        self.after(100, self.poll_events)

    def run(self):
        if not self.pdfs:
            messagebox.showwarning("PDF seçilmedi", "Önce PDF veya klasör seçin.")
            return
        out = filedialog.asksaveasfilename(
            title="Excel'i kaydet", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialdir=self.pdfs[0].parent, initialfile="brut_ucret.xlsx")
        if not out:
            return
        self.out_path = Path(out)
        self.run_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.progress.start(10)
        threading.Thread(target=self.work, args=(list(self.pdfs), self.out_path), daemon=True).start()

    def work(self, pdfs, out):
        try:
            ok = write_excel(pdfs, out, log=lambda m: self.log(m, None if "UYUMSUZ" not in m else "err"))
            if ok:
                self.log("Tüm kontroller TAMAM.", "ok")
            else:
                self.log("Bazı PDF'lerde kontrol tutmadı. Excel'deki 'Kontrol' sayfasına bakın.", "err")
            self.events.put(("done", True))
        except PermissionError:
            self.log(f"Excel yazılamadı: {out.name} açık olabilir, kapatıp tekrar deneyin.", "err")
            self.events.put(("done", False))
        except Exception as e:
            self.log(f"Hata: {e}", "err")
            self.events.put(("done", False))

    def done(self, success):
        self.progress.stop()
        self.run_btn.config(state="normal")
        if success:
            self.open_btn.config(state="normal")

    def open_excel(self):
        if self.out_path and self.out_path.exists():
            os.startfile(self.out_path)


if __name__ == "__main__":
    App().mainloop()
