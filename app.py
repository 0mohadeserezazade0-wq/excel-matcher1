#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""رابط فارسی برای تطبیق جفت‌ستون‌های اکسل و خروجی گرفتن ردیف‌های منطبق."""
import csv
import os
from datetime import date, datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

PALETTES = {
    "orange": "FFFFA500",
    "yellow": "FFFFFF00",
    "green": "FF92D050",
    "blue": "FF5B9BD5",
    "red": "FFFF6666",
    "purple": "FFB4A7D6",
}


def normalize(value):
    """یکسان‌سازی مقدارهای متنی و عددی اکسل؛ مقدار خالی با خالی تطبیق نمی‌یابد."""
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
        return format(number, ".15g")
    except (ValueError, OverflowError):
        return text.casefold()


def find_header(ws, required):
    wanted = {name.casefold() for name in required}
    for row_num in range(1, ws.max_row + 1):
        cols = {}
        for col_num in range(1, ws.max_column + 1):
            val = ws.cell(row_num, col_num).value
            if val is not None:
                cols[str(val).strip().casefold()] = col_num
        if wanted.issubset(cols):
            return row_num, cols
    raise ValueError(
        f"در برگهٔ «{ws.title}» ستون‌های لازم پیدا نشد: «{'، '.join(required)}»"
    )


def locate_header(wb, required):
    for ws in wb.worksheets:
        try:
            row, cols = find_header(ws, required)
            return ws, row, cols
        except ValueError:
            continue
    raise ValueError("ستون‌های موردنیاز در هیچ‌یک از برگه‌ها پیدا نشدند: «" + "، ".join(required) + "»")


def parse_date(value):
    """تبدیل ورودی تاریخ به تاریخ واقعی اکسل؛ از قالب‌های رایج پشتیبانی می‌کند."""
    if not value or not value.strip():
        return None
    text = value.strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError("قالب تاریخ معتبر نیست؛ نمونه: 2025-01-31 یا 31/01/2025")


def match_files(reference_path, main_path, output_dir=None, color_name="orange", date_value=""):
    """ردیف‌های منطبق را در کنار فایل اصلی به صورت اکسل رنگی و CSV متنی ذخیره می‌کند."""
    if os.path.abspath(reference_path) == os.path.abspath(main_path):
        raise ValueError("فایل رفرنس و فایل اصلی باید دو فایل جداگانه باشند.")
    if not os.path.isfile(reference_path) or not os.path.isfile(main_path):
        raise FileNotFoundError("یکی از فایل‌های انتخاب‌شده پیدا نشد. مسیرها را بررسی کنید.")
    color_name = str(color_name).strip().lower()
    if color_name not in PALETTES:
        raise ValueError("رنگ انتخاب‌شده معتبر نیست.")
    replacement_date = parse_date(date_value)
    output_dir = os.path.dirname(os.path.abspath(main_path))

    ref_wb = load_workbook(reference_path, read_only=True, data_only=True)
    try:
        ref_ws, ref_header, ref_cols = locate_header(ref_wb, ("Signal", "Signal_1"))
        ref_pairs = set()
        for r in range(ref_header + 1, ref_ws.max_row + 1):
            pair = (normalize(ref_ws.cell(r, ref_cols["signal"]).value),
                    normalize(ref_ws.cell(r, ref_cols["signal_1"]).value))
            if pair[0] is not None and pair[1] is not None:
                ref_pairs.add(pair)
    finally:
        ref_wb.close()

    keep_vba = main_path.lower().endswith(".xlsm")
    main_wb = load_workbook(main_path, keep_vba=keep_vba)
    try:
        main_ws, header, cols = locate_header(main_wb, ("Signal_1", "Signal_1_Value"))
        matches = []
        fill = PatternFill(fill_type="solid", fgColor=PALETTES[color_name])
        for r in range(header + 1, main_ws.max_row + 1):
            pair = (normalize(main_ws.cell(r, cols["signal_1"]).value),
                    normalize(main_ws.cell(r, cols["signal_1_value"]).value))
            if pair[0] is not None and pair[1] is not None and pair in ref_pairs:
                matches.append(r)
                for c in range(1, main_ws.max_column + 1):
                    main_ws.cell(r, c).fill = fill

        if replacement_date is not None:
            for ws in main_wb.worksheets:
                for row_num in range(1, ws.max_row + 1):
                    for col_num in range(1, ws.max_column + 1):
                        value = ws.cell(row_num, col_num).value
                        if isinstance(value, str) and value.strip().casefold() == "date":
                            for data_row in range(row_num + 1, ws.max_row + 1):
                                ws.cell(data_row, col_num).value = replacement_date

        stem = os.path.splitext(os.path.basename(main_path))[0]
        xlsx_path = os.path.join(output_dir, f"{stem}_{color_name}.xlsx")
        csv_path = os.path.join(output_dir, f"{stem}_matched.csv")
        main_wb.save(xlsx_path)
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow([main_ws.cell(header, c).value for c in range(1, main_ws.max_column + 1)])
            writer.writerow(["-" * 60])
            for r in matches:
                writer.writerow([main_ws.cell(r, c).value for c in range(1, main_ws.max_column + 1)])
                writer.writerow(["-" * 60])
        return len(matches), xlsx_path, csv_path
    finally:
        main_wb.close()

class ToolTip:
    """راهنمای کوچک فارسی هنگام قرارگرفتن نشانگر روی کنترل."""
    def __init__(self, widget, text):
        self.widget, self.text, self.tip = widget, text, None
        widget.bind("<Enter>", self.show, add="+")
        widget.bind("<Leave>", self.hide, add="+")

    def show(self, _event=None):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.text, bg="#fff8dc", fg="#222222",
                 relief="solid", borderwidth=1, padx=7, pady=4,
                 font=("Tahoma", 9), justify="right").pack()

    def hide(self, _event=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


class MatchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("تطبیق فایل‌های اکسل")
        self.geometry("820x440")
        self.minsize(680, 350)
        self.configure(bg="#f4f6f8")
        self.reference = tk.StringVar()
        self.main_file = tk.StringVar()
        self.color_name = tk.StringVar(value="orange")
        self.date_value = tk.StringVar()
        self._build()

    def _build(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TLabel", font=("Tahoma", 10), background="#f4f6f8")
        style.configure("TButton", font=("Tahoma", 10), padding=(10, 6))
        style.configure("Go.TButton", font=("Tahoma", 11, "bold"), padding=(14, 9))
        style.configure("TEntry", font=("Tahoma", 10), padding=5)

        frame = ttk.Frame(self, padding=22)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        title = ttk.Label(frame, text="تطبیق توضیحات رفرنس با فایل اصلی بروکر",
                          font=("Tahoma", 14, "bold"), anchor="e", justify="right")
        title.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        subtitle = ttk.Label(frame, text="جفت‌ستون‌های سیگنال را تطبیق دهید و دو خروجی دریافت کنید.",
                             anchor="e", justify="right")
        subtitle.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 18))

        date_label = ttk.Label(frame, text="تاریخ جایگزین (اختیاری؛ نمونه 2025-01-31)", anchor="e", justify="right")
        date_label.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        ToolTip(date_label, "در صورت واردکردن تاریخ، مقادیر ستون Date در فایل اصلی جایگزین می‌شوند.")
        date_entry = ttk.Entry(frame, textvariable=self.date_value, justify="right")
        date_entry.grid(row=3, column=0, sticky="ew")
        ToolTip(date_entry, "خالی بماند، تاریخ‌های فعلی حفظ می‌شوند.")
        color_label = ttk.Label(frame, text="رنگ هایلایت", anchor="e", justify="right")
        color_label.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(7, 4))
        self.color_box = ttk.Combobox(frame, textvariable=self.color_name,
                                      values=tuple(PALETTES), state="readonly", justify="right")
        self.color_box.grid(row=5, column=0, sticky="ew")
        ToolTip(self.color_box, "رنگ انتخابی در فایل اکسل و نام فایل خروجی استفاده می‌شود.")

        self._file_row(frame, 6, "اکسل توضیحات (رفرنس)", self.reference,
                       "انتخاب فایل رفرنس؛ باید ستون‌های Signal و Signal_1 را داشته باشد.", self.pick_reference)
        self._file_row(frame, 8, "اکسل اصلی (بروکر)", self.main_file,
                       "انتخاب فایل اصلی؛ خروجی‌ها کنار همین فایل ذخیره می‌شوند.", self.pick_main)

        self.run_button = ttk.Button(frame, text="شروع پردازش و تطبیق", style="Go.TButton", command=self.run)
        self.run_button.grid(row=10, column=0, sticky="e", pady=(18, 0))
        ToolTip(self.run_button, "پردازش را آغاز می‌کند و دو خروجی را کنار فایل اصلی می‌سازد.")
        self.status = ttk.Label(frame, text="آمادهٔ پردازش", foreground="#444444", anchor="e")
        self.status.grid(row=11, column=0, columnspan=2, sticky="ew", pady=(12, 0))

    def _file_row(self, parent, row, label_text, variable, tip, command):
        label = ttk.Label(parent, text=label_text, anchor="e", justify="right")
        label.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(7, 4))
        ToolTip(label, tip)
        entry = ttk.Entry(parent, textvariable=variable, justify="right")
        entry.grid(row=row + 1, column=0, sticky="ew")
        ToolTip(entry, tip)
        button = ttk.Button(parent, text="مرور...", command=command)
        button.grid(row=row + 1, column=1, padx=(9, 0), sticky="e")
        ToolTip(button, tip)

    def _pick_excel(self, variable, title):
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("فایل‌های اکسل", "*.xlsx *.xlsm"), ("همهٔ فایل‌ها", "*.*")])
        if path:
            variable.set(path)

    def pick_reference(self):
        self._pick_excel(self.reference, "انتخاب اکسل توضیحات (رفرنس)")

    def pick_main(self):
        self._pick_excel(self.main_file, "انتخاب اکسل اصلی (بروکر)")

    def run(self):
        ref, main = self.reference.get().strip(), self.main_file.get().strip()
        if not ref or not main:
            messagebox.showwarning("انتخاب‌های ناقص", "لطفاً فایل رفرنس و فایل اصلی را انتخاب کنید.", parent=self)
            return
        self.run_button.configure(state="disabled")
        self.status.configure(text="در حال پردازش فایل‌ها...", foreground="#333333")
        self.update_idletasks()
        try:
            count, xlsx_path, csv_path = match_files(ref, main, color_name=self.color_name.get(), date_value=self.date_value.get())
            self.status.configure(text=f"پردازش با موفقیت انجام شد؛ تعداد ردیف‌های منطبق: {count}", foreground="#16733b")
            messagebox.showinfo(
                "پردازش با موفقیت انجام شد",
                f"تطبیق کامل شد.\nتعداد ردیف‌های منطبق: {count}\n\n"
                f"فایل اکسل رنگی‌شده:\n{xlsx_path}\n\nفایل CSV ردیف‌های مطابق:\n{csv_path}", parent=self)
        except Exception as exc:
            self.status.configure(text="پردازش ناموفق بود؛ جزئیات را بررسی کنید.", foreground="#b00020")
            messagebox.showerror("خطا هنگام پردازش", f"پردازش انجام نشد.\n\nجزئیات خطا:\n{exc}", parent=self)
        finally:
            self.run_button.configure(state="normal")


if __name__ == "__main__":
    MatchApp().mainloop()
