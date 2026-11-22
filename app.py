import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

try:
    import win32print
    import win32ui
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


def get_installed_printers():
    if not HAS_WIN32:
        return ["SATO SA408 SEPL"]
    printers = []
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    for p in win32print.EnumPrinters(flags):
        printers.append(p[2])
    return printers or ["SATO SA408 SEPL"]


def build_sbpl_label(rows_data, x_pos=30, start_y=40, line_gap=38, copies=1):
    """
    Generates a basic SBPL label string for up to 4 rows.
    rows_data: list of tuples -> (text, scale_multiplier)
    """
    esc = "\x1b"
    parts = [f"{esc}A"]
    cur_y = start_y

    for text, scale in rows_data:
        if not text:
            continue
        s = max(1, min(int(scale), 9))
        h_mult = f"{s:02d}"
        w_mult = f"{s:02d}"
        parts.append(f"{esc}V{cur_y:04d}{esc}H{x_pos:04d}{esc}L{h_mult}{w_mult}{esc}${text}")
        cur_y += int(line_gap * (0.75 + 0.25 * s))

    parts.append(f"{esc}Q{max(1, int(copies))}")
    parts.append(f"{esc}Z")
    return "".join(parts)


def send_raw_sbpl(printer_name, sbpl_payload):
    if not HAS_WIN32:
        raise RuntimeError("pywin32 is required on Windows to send print jobs.")
    h_printer = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(h_printer, 1, ("SATO SBPL Label", None, "RAW"))
        try:
            win32print.StartPagePrinter(h_printer)
            win32print.WritePrinter(h_printer, sbpl_payload.encode("ascii", errors="ignore"))
            win32print.EndPagePrinter(h_printer)
        finally:
            win32print.EndDocPrinter(h_printer)
    finally:
        win32print.ClosePrinter(h_printer)


def print_label_gdi(printer_name, doc_title, lines_cfg, width_mm, height_mm):
    """
    Prints a single sticker via Windows GDI using exact font name, size, and style per row.
    lines_cfg: list of dicts -> {'text': str, 'font': str, 'size': int, 'bold': bool, 'italic': bool}
    """
    if not HAS_WIN32:
        raise RuntimeError("pywin32 is required on Windows to print via driver.")

    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(printer_name)

    dpi_x = hdc.GetDeviceCaps(88)   # LOGPIXELSX
    dpi_y = hdc.GetDeviceCaps(90)   # LOGPIXELSY

    page_w = int((width_mm / 25.4) * dpi_x)
    page_h = int((height_mm / 25.4) * dpi_y)

    margin_x = int(0.06 * dpi_x)
    margin_y = int(0.05 * dpi_y)

    hdc.StartDoc(doc_title)
    hdc.StartPage()

    cur_y = margin_y
    for item in lines_cfg:
        text = item["text"]
        if not text:
            continue

        pt_size = max(5, int(item["size"]))
        lf_height = -round(pt_size * dpi_y / 72.0)
        weight = 700 if item["bold"] else 400
        italic = 1 if item["italic"] else 0

        font_spec = {
            "name": item["font"],
            "height": lf_height,
            "weight": weight,
            "italic": italic,
        }
        font = win32ui.CreateFont(font_spec)
        old_font = hdc.SelectObject(font)

        rect = (margin_x, cur_y, max(margin_x + 10, page_w - margin_x), max(cur_y + 10, page_h - margin_y))
        hdc.DrawText(text, rect, 0x00000000)  # DT_LEFT

        _, text_h = hdc.GetTextExtent(text if text else "A")
        cur_y += max(text_h + int(0.02 * dpi_y), abs(lf_height) + 2)

        hdc.SelectObject(old_font)
        win32print  # keep imported

    hdc.EndPage()
    hdc.EndDoc()
    hdc.DeleteDC()


class SatoStickerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SATO SA4 Excel Continuous Label Printer")
        self.root.geometry("820x680")

        self.excel_path = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.printer_var = tk.StringVar(value=get_installed_printers()[0])
        self.mode_var = tk.StringVar(value="Windows Driver (Font/Style)")
        self.width_mm_var = tk.StringVar(value="50.8")
        self.height_mm_var = tk.StringVar(value="25.4")
        self.copies_var = tk.StringVar(value="1")
        self.start_row_var = tk.StringVar(value="1")
        self.end_row_var = tk.StringVar(value="0")  # 0 = all rows

        self.df = None
        self.row_widgets = []

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # --- File & Sheet Selection ---
        file_frame = ttk.LabelFrame(self.root, text="1. Excel File & Sheet")
        file_frame.pack(fill="x", **pad)

        ttk.Entry(file_frame, textvariable=self.excel_path, width=62).grid(row=0, column=0, padx=6, pady=6)
        ttk.Button(file_frame, text="Browse Excel", command=self.browse_excel).grid(row=0, column=1, padx=4)

        ttk.Label(file_frame, text="Sheet:").grid(row=0, column=2, padx=(10, 4))
        self.sheet_combo = ttk.Combobox(file_frame, textvariable=self.sheet_var, state="readonly", width=18)
        self.sheet_combo.grid(row=0, column=3, padx=4)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_change)

        # --- Printer & Sticker Settings ---
        cfg_frame = ttk.LabelFrame(self.root, text="2. Printer & Sticker Size")
        cfg_frame.pack(fill="x", **pad)

        ttk.Label(cfg_frame, text="Printer:").grid(row=0, column=0, padx=6, pady=4, sticky="w")
        self.printer_combo = ttk.Combobox(
            cfg_frame, textvariable=self.printer_var, values=get_installed_printers(), width=28
        )
        self.printer_combo.grid(row=0, column=1, padx=4, pady=4)

        ttk.Label(cfg_frame, text="Print Mode:").grid(row=0, column=2, padx=6, pady=4, sticky="w")
        mode_combo = ttk.Combobox(
            cfg_frame,
            textvariable=self.mode_var,
            values=["Windows Driver (Font/Style)", "Raw SBPL"],
            state="readonly",
            width=26,
        )
        mode_combo.grid(row=0, column=3, padx=4, pady=4)

        ttk.Label(cfg_frame, text="Width (mm):").grid(row=1, column=0, padx=6, pady=4, sticky="w")
        ttk.Entry(cfg_frame, textvariable=self.width_mm_var, width=10).grid(row=1, column=1, sticky="w", padx=4)

        ttk.Label(cfg_frame, text="Height (mm):").grid(row=1, column=2, padx=6, pady=4, sticky="w")
        ttk.Entry(cfg_frame, textvariable=self.height_mm_var, width=10).grid(row=1, column=3, sticky="w", padx=4)

        ttk.Label(cfg_frame, text="Copies/Sticker:").grid(row=2, column=0, padx=6, pady=4, sticky="w")
        ttk.Entry(cfg_frame, textvariable=self.copies_var, width=10).grid(row=2, column=1, sticky="w", padx=4)

        ttk.Label(cfg_frame, text="Row Range (Start-End, 0=End):").grid(row=2, column=2, padx=6, pady=4, sticky="w")
        rng_sub = ttk.Frame(cfg_frame)
        rng_sub.grid(row=2, column=3, sticky="w")
        ttk.Entry(rng_sub, textvariable=self.start_row_var, width=6).pack(side="left")
        ttk.Label(rng_sub, text=" to ").pack(side="left")
        ttk.Entry(rng_sub, textvariable=self.end_row_var, width=6).pack(side="left")

        # --- Sticker Lines (Up to 4 Rows) ---
        lines_frame = ttk.LabelFrame(self.root, text="3. Sticker Rows (Map Excel Column or Custom Text + Font)")
        lines_frame.pack(fill="x", **pad)

        headers = ["Row", "Source Type", "Column / Custom Text", "Font", "Size", "Bold", "Italic"]
        for col_idx, h in enumerate(headers):
            ttk.Label(lines_frame, text=h, font=("Segoe UI", 9, "bold")).grid(row=0, column=col_idx, padx=4, pady=3)

        font_choices = ["Arial", "Calibri", "Consolas", "Courier New", "Segoe UI", "Times New Roman"]

        for i in range(4):
            row_num = i + 1
            ttk.Label(lines_frame, text=f"Line {row_num}").grid(row=row_num, column=0, padx=4, pady=3)

            src_type = tk.StringVar(value="Column" if i < 3 else "None")
            src_cb = ttk.Combobox(
                lines_frame, textvariable=src_type, values=["Column", "Custom Text", "None"], state="readonly", width=11
            )
            src_cb.grid(row=row_num, column=1, padx=3)

            val_var = tk.StringVar(value="")
            val_entry = ttk.Entry(lines_frame, textvariable=val_var, width=22)
            val_entry.grid(row=row_num, column=2, padx=3)

            font_var = tk.StringVar(value="Arial")
            font_cb = ttk.Combobox(lines_frame, textvariable=font_var, values=font_choices, width=14)
            font_cb.grid(row=row_num, column=3, padx=3)

            default_sizes = [11, 9, 8, 8]
            size_var = tk.StringVar(value=str(default_sizes[i]))
            size_entry = ttk.Entry(lines_frame, textvariable=size_var, width=5)
            size_entry.grid(row=row_num, column=4, padx=3)

            bold_var = tk.BooleanVar(value=(i == 0))
            ttk.Checkbutton(lines_frame, variable=bold_var).grid(row=row_num, column=5, padx=3)

            italic_var = tk.BooleanVar(value=False)
            ttk.Checkbutton(lines_frame, variable=italic_var).grid(row=row_num, column=6, padx=3)

            self.row_widgets.append({
                "src_type": src_type,
                "val_var": val_var,
                "font_var": font_var,
                "size_var": size_var,
                "bold_var": bold_var,
                "italic_var": italic_var,
            })

        # --- Preview & Actions ---
        act_frame = ttk.Frame(self.root)
        act_frame.pack(fill="x", **pad)

        ttk.Button(act_frame, text="Preview First Sticker", command=self.preview_first).pack(side="left", padx=4)
        ttk.Button(act_frame, text="Print Continuous", command=self.print_continuous).pack(side="left", padx=8)

        self.status_label = ttk.Label(act_frame, text="Ready")
        self.status_label.pack(side="right", padx=8)

        preview_frame = ttk.LabelFrame(self.root, text="4. First Sticker Preview & Log")
        preview_frame.pack(fill="both", expand=True, **pad)

        self.log_text = tk.Text(preview_frame, height=12, wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)

    def browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return
        self.excel_path.set(path)
        try:
            xls = pd.ExcelFile(path)
            sheets = xls.sheet_names
            self.sheet_combo["values"] = sheets
            if sheets:
                self.sheet_var.set(sheets[0])
                self.load_sheet(sheets[0])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel file:\n{e}")

    def on_sheet_change(self, _event=None):
        if self.sheet_var.get():
            self.load_sheet(self.sheet_var.get())

    def load_sheet(self, sheet_name):
        try:
            self.df = pd.read_excel(self.excel_path.get(), sheet_name=sheet_name)
            cols = [str(c) for c in self.df.columns]
            for idx, w in enumerate(self.row_widgets):
                if idx < len(cols) and w["src_type"].get() == "Column":
                    w["val_var"].set(cols[idx])
            self.log(f"Loaded sheet '{sheet_name}' ({len(self.df)} rows). Columns: {', '.join(cols)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read sheet:\n{e}")

    def log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")

    def _resolve_lines_for_row(self, df_row):
        lines = []
        for w in self.row_widgets:
            stype = w["src_type"].get()
            raw_val = w["val_var"].get().strip()
            if stype == "None":
                continue
            if stype == "Custom Text":
                text = raw_val
            else:
                if self.df is None or raw_val not in self.df.columns:
                    text = ""
                else:
                    val = df_row[raw_val]
                    text = "" if pd.isna(val) else str(val).strip()

            try:
                size = int(w["size_var"].get())
            except ValueError:
                size = 9

            lines.append({
                "text": text,
                "font": w["font_var"].get() or "Arial",
                "size": size,
                "bold": bool(w["bold_var"].get()),
                "italic": bool(w["italic_var"].get()),
            })
        return lines

    def _get_target_rows(self):
        if self.df is None or self.df.empty:
            raise ValueError("Please load a valid Excel sheet first.")
        try:
            start_r = max(1, int(self.start_row_var.get()))
            end_r = int(self.end_row_var.get())
        except ValueError:
            raise ValueError("Start/End row must be whole numbers.")

        start_idx = start_r - 1
        end_idx = len(self.df) if end_r <= 0 else min(len(self.df), end_r)
        if start_idx >= end_idx:
            raise ValueError("Invalid row range.")
        return self.df.iloc[start_idx:end_idx]

    def preview_first(self):
        try:
            subset = self._get_target_rows()
            first_row = subset.iloc[0]
            lines = self._resolve_lines_for_row(first_row)
            self.log("--- Sticker Preview (First Row in Range) ---")
            for idx, ln in enumerate(lines, 1):
                style = ("Bold " if ln["bold"] else "") + ("Italic" if ln["italic"] else "")
                style = style.strip() or "Regular"
                self.log(f"  Row {idx}: '{ln['text']}' | {ln['font']} {ln['size']}pt ({style})")
        except Exception as e:
            messagebox.showwarning("Preview", str(e))

    def print_continuous(self):
        try:
            subset = self._get_target_rows()
            printer = self.printer_var.get().strip()
            if not printer:
                raise ValueError("Please select or enter a printer name.")

            copies = max(1, int(self.copies_var.get()))
            width_mm = float(self.width_mm_var.get())
            height_mm = float(self.height_mm_var.get())
            mode = self.mode_var.get()

            printed_count = 0
            for idx, row in subset.iterrows():
                lines = self._resolve_lines_for_row(row)
                if not any(ln["text"] for ln in lines):
                    continue

                for _ in range(copies):
                    if mode == "Raw SBPL":
                        sbpl_rows = [(ln["text"], max(1, round(ln["size"] / 8))) for ln in lines]
                        payload = build_sbpl_label(sbpl_rows, copies=1)
                        send_raw_sbpl(printer, payload)
                    else:
                        print_label_gdi(
                            printer,
                            f"SATO Label Row {idx + 1}",
                            lines,
                            width_mm,
                            height_mm,
                        )
                    printed_count += 1
                    self.status_label.config(text=f"Printed: {printed_count}")
                    self.root.update_idletasks()

            self.log(f"Done. Total stickers sent to '{printer}': {printed_count}")
            messagebox.showinfo("Complete", f"Sent {printed_count} sticker(s) to {printer}.")
        except Exception as e:
            messagebox.showerror("Print Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = SatoStickerApp(root)
    root.mainloop()