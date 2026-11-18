"""Local desktop UI. Files never leave this computer."""

from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from convert import InputPDF, convert_pdfs, sheet_names

HELP_TEXT = """How to use

1. Click "Add PDFs…" and select one or more Student Assessment Results PDFs
   (Ctrl-click on Windows, Command-click on Mac to pick several).
2. Each PDF becomes one Excel tab. Tab names default to the PDF filename.
   Double-click a row to type a different tab name.
3. Enter the workbook filename and choose the save folder.
4. Optional: type an index sheet name to add a first sheet listing all PDFs.
   Leave it blank for the same layout as the existing results workbooks.
5. Click "Convert to Excel".

Good to know
- Excel tab names are limited to 31 characters; long names are shortened.
- Marks, totals and grades are copied exactly as printed in the PDF.
- Failed rows (F, F+, F-, U) are shaded; PDF markers such as N appear as
  cell comments.
- If any PDF cannot be read, nothing is saved and the problem is shown.
- Everything runs on this computer. No internet connection is used.
- Only this report format is supported; scanned PDFs cannot be read."""


def default_folder():
    documents = Path.home() / "Documents"
    return documents if documents.is_dir() else Path.home()


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.inputs = []
        self.events = queue.Queue()
        self.busy = False
        self.output = tk.StringVar(value="Assessment Results.xlsx")
        self.main_sheet = tk.StringVar(value="")
        self.folder = tk.StringVar(value=str(default_folder()))
        self.status = tk.StringVar(value="Add PDFs to get started. All processing stays on this computer.")
        root.title("Assessment PDF to Excel")
        root.geometry("1050x690")
        root.minsize(850, 580)
        root.protocol("WM_DELETE_WINDOW", self.close)
        frame = ttk.Frame(root, padding=24)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)
        ttk.Label(frame, text="Assessment PDF to Excel", font=("Helvetica", 22, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(frame, text="Select multiple PDFs, review the tab names, then save one Excel workbook.").grid(row=1, column=0, sticky="w", pady=(6, 20))
        actions = ttk.Frame(frame)
        actions.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        self.controls = []
        for text, command in [
            ("Add PDFs…", self.add_files), ("Remove selected", self.remove),
            ("Rename selected tab…", self.rename), ("Use PDF names", self.reset_names),
        ]:
            button = ttk.Button(actions, text=text, command=command)
            button.pack(side="left", padx=(0, 8))
            self.controls.append(button)
        ttk.Button(actions, text="Help", command=self.show_help).pack(side="right")
        list_frame = ttk.Frame(frame)
        list_frame.grid(row=3, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(list_frame, columns=("pdf", "tab"), show="headings", selectmode="extended")
        self.tree.heading("pdf", text="PDF filename")
        self.tree.heading("tab", text="Excel tab name (double-click to edit)")
        self.tree.column("pdf", width=590, minwidth=300)
        self.tree.column("tab", width=330, minwidth=220)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<Double-1>", self.rename)
        ttk.Label(frame, text="Excel allows 31 characters per tab. Long names are shortened; duplicates get a suffix.", wraplength=960).grid(row=4, column=0, sticky="w", pady=(10, 16))
        options = ttk.LabelFrame(frame, text="Workbook settings", padding=14)
        options.grid(row=5, column=0, sticky="ew")
        options.columnconfigure(1, weight=1)
        for row, (label, variable) in enumerate([
            ("Workbook filename", self.output), ("Index sheet name (optional)", self.main_sheet), ("Save folder", self.folder),
        ]):
            ttk.Label(options, text=label).grid(row=row, column=0, sticky="w", padx=(0, 16), pady=5)
            entry = ttk.Entry(options, textvariable=variable)
            entry.grid(row=row, column=1, sticky="ew", pady=5)
            self.controls.append(entry)
        browse = ttk.Button(options, text="Choose…", command=self.choose_folder)
        browse.grid(row=2, column=2, padx=(10, 0))
        self.controls.append(browse)
        self.main_sheet.trace_add("write", lambda *_: self.refresh())
        bottom = ttk.Frame(frame)
        bottom.grid(row=6, column=0, sticky="ew", pady=(18, 0))
        bottom.columnconfigure(0, weight=1)
        ttk.Label(bottom, textvariable=self.status, wraplength=740).grid(row=0, column=0, sticky="w")
        self.convert_button = ttk.Button(bottom, text="Convert to Excel", command=self.convert, state="disabled")
        self.convert_button.grid(row=0, column=1, padx=(14, 0))
        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.grid(row=7, column=0, sticky="ew", pady=(14, 0))
        root.after(100, self.poll)

    def refresh(self):
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        _, names = sheet_names(self.inputs, self.main_sheet.get())
        for i, (item, name) in enumerate(zip(self.inputs, names)):
            self.tree.insert("", "end", iid=str(i), values=(item.path.name, name))
        for item_id in selected:
            if self.tree.exists(item_id):
                self.tree.selection_add(item_id)
        if not self.busy:
            self.convert_button.configure(state="normal" if self.inputs else "disabled")
            self.status.set(f"{len(self.inputs)} PDF(s) selected. Double-click a row to enter a custom tab name.")

    def add_files(self):
        if self.busy:
            return
        files = filedialog.askopenfilenames(title="Select assessment PDFs (multiple selection allowed)", filetypes=[("PDF files", "*.pdf"), ("All files", "*")])
        existing = {item.path.resolve() for item in self.inputs}
        for filename in files:
            path = Path(filename).resolve()
            if path.suffix.casefold() == ".pdf" and path not in existing:
                self.inputs.append(InputPDF(path))
                existing.add(path)
        self.refresh()

    def remove(self):
        if self.busy:
            return
        selected = {int(item) for item in self.tree.selection()}
        self.inputs = [item for i, item in enumerate(self.inputs) if i not in selected]
        self.refresh()

    def rename(self, event=None):
        if self.busy:
            return
        selected = self.tree.selection()
        if event is not None:
            item = self.tree.identify_row(event.y)
            selected = (item,) if item else ()
        if len(selected) != 1:
            messagebox.showinfo("Choose one PDF", "Select one PDF row to rename its Excel tab.", parent=self.root)
            return
        index = int(selected[0])
        item = self.inputs[index]
        value = simpledialog.askstring("Excel tab name", f"Tab name for {item.path.name}\nLeave blank to use its PDF filename.", initialvalue=item.tab_name or self.tree.item(str(index), "values")[1], parent=self.root)
        if value is not None:
            self.inputs[index] = InputPDF(item.path, value.strip())
            self.refresh()

    def reset_names(self):
        if self.busy:
            return
        selected = {int(item) for item in self.tree.selection()}
        self.inputs = [InputPDF(item.path) if not selected or i in selected else item for i, item in enumerate(self.inputs)]
        self.refresh()

    def choose_folder(self):
        folder = filedialog.askdirectory(title="Where should the Excel workbook be saved?", initialdir=self.folder.get(), parent=self.root)
        if folder:
            self.folder.set(folder)

    def set_busy(self, busy):
        self.busy = busy
        for control in self.controls:
            control.configure(state="disabled" if busy else "normal")
        self.convert_button.configure(state="disabled" if busy or not self.inputs else "normal")
        self.progress.start(12) if busy else self.progress.stop()

    def convert(self):
        if self.busy or not self.inputs:
            return
        filename = self.output.get().strip()
        if not filename or filename in (".", "..") or any(ch in filename for ch in '/\\:*?"<>|'):
            messagebox.showerror("Workbook filename", "Enter a filename without folders or special filename characters.", parent=self.root)
            return
        if not filename.casefold().endswith(".xlsx"):
            filename += ".xlsx"
            self.output.set(filename)
        if not self.folder.get().strip():
            messagebox.showerror("Workbook settings", "Choose a save folder.", parent=self.root)
            return
        output = Path(self.folder.get()).expanduser() / filename
        overwrite = output.exists()
        if overwrite and not messagebox.askyesno("Replace workbook?", f"This file already exists:\n{output}\n\nReplace it after all PDFs pass validation?", parent=self.root):
            return
        inputs = tuple(self.inputs)
        main_sheet = self.main_sheet.get().strip() or None
        self.set_busy(True)

        def worker():
            try:
                result = convert_pdfs(inputs, output, main_sheet, overwrite=overwrite, progress=lambda text: self.events.put(("progress", text)))
                self.events.put(("done", result))
            except Exception as error:
                self.events.put(("error", str(error)))

        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            while True:
                event, value = self.events.get_nowait()
                if event == "progress":
                    self.status.set(value)
                else:
                    self.set_busy(False)
                    if event == "error":
                        self.status.set("Conversion failed. Fix the reported input or choose another output location.")
                        messagebox.showerror("Conversion failed", value, parent=self.root)
                    else:
                        text = f"Saved {len(value.sheet_names)} PDF tabs and {value.student_count:,} student rows."
                        self.status.set(text)
                        messagebox.showinfo("Workbook saved", f"{text}\n\n{value.output}", parent=self.root)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def show_help(self):
        messagebox.showinfo("Help", HELP_TEXT, parent=self.root)

    def close(self):
        if self.busy:
            messagebox.showinfo("Conversion running", "Please wait for the current conversion to finish before closing.", parent=self.root)
        else:
            self.root.destroy()


def launch():
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    launch()
