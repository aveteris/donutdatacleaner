
"""
Donut Data Cleaner
--------------------
A cute, candy-colored pixel-art desktop GUI for cleaning messy CSV/Excel files.
Buttons are drawn with a textured "pixel game" look: rounded pill shapes,
thick dark outlines, and a glossy highlight streak across the top -
inspired by chunky retro game UI buttons.

Features:
 - Load CSV or Excel (.csv, .xlsx, .xls)
 - Preview raw data
 - One-click cleanup:
    * Remove exact duplicate rows
    * Strip whitespace from text columns
    * Standardize column names (snake_case, lowercase)
    * Auto-detect & reformat date columns
    * Auto-detect & coerce numeric columns
    * Fill/flag missing values
 - Preview cleaned data
 - Export to CSV or Excel

Run locally with: python donut_data_cleaner.py
Requires: pandas, openpyxl (for Excel), tkinter (bundled with most Python installs)
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import re
import os

# ---------- Color Palette (candy pixel game style) ----------
BG_MAIN      = "#fff6f0"   # warm cream
BG_PANEL     = "#ffffff"
COLOR_BLUE   = "#5ec9f0"
COLOR_PINK   = "#f76fa0"
COLOR_PURPLE = "#a879e6"
COLOR_ORANGE = "#ffb24d"
COLOR_YELLOW = "#ffd23f"
OUTLINE      = "#2b2b2b"
TEXT_LIGHT   = "#ffffff"
TEXT_DARK    = "#2b2b2b"
PIXEL_FONT_NAME = "Arial Rounded MT Bold"   # falls back gracefully if unavailable
BTN_FONT     = (PIXEL_FONT_NAME, 11, "bold")
TITLE_FONT   = (PIXEL_FONT_NAME, 20, "bold")
LOG_FONT     = ("Consolas", 9)


class TexturedButton(tk.Canvas):
    """A chunky, rounded, textured pixel-game-style button drawn on a Canvas."""

    def __init__(self, parent, text, command, color=COLOR_PURPLE,
                 width=170, height=48, font=BTN_FONT, fg=TEXT_LIGHT):
        super().__init__(parent, width=width, height=height,
                          bg=parent["bg"], highlightthickness=0, bd=0, cursor="hand2")
        self.command = command
        self.color = color
        self.width = width
        self.height = height
        self.text = text
        self.font = font
        self.fg = fg
        self._pressed = False
        self._draw(color)
        self.bind("<Button-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _lighten(self, hex_color, amount=40):
        hex_color = hex_color.lstrip("#")
        r, g, b = (int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        r, g, b = (min(255, c + amount) for c in (r, g, b))
        return f"#{r:02x}{g:02x}{b:02x}"

    def _darken(self, hex_color, amount=40):
        hex_color = hex_color.lstrip("#")
        r, g, b = (int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        r, g, b = (max(0, c - amount) for c in (r, g, b))
        return f"#{r:02x}{g:02x}{b:02x}"

    def _rounded_rect(self, x1, y1, x2, y2, r, **kwargs):
        points = [
            x1+r, y1, x2-r, y1, x2, y1, x2, y1+r, x2, y2-r, x2, y2,
            x2-r, y2, x1+r, y2, x1, y2, x1, y2-r, x1, y1+r, x1, y1
        ]
        return self.create_polygon(points, smooth=True, **kwargs)

    def _draw(self, color):
        self.delete("all")
        pad = 3
        shadow_color = self._darken(color, 60)
        # drop shadow (bottom-right offset) for a chunky pixel-game feel
        self._rounded_rect(pad+3, pad+4, self.width-pad+3, self.height-pad+4,
                            14, fill=shadow_color, outline="")
        # main button body with thick outline
        body = self._rounded_rect(pad, pad, self.width-pad, self.height-pad,
                                   14, fill=color, outline=OUTLINE, width=3)
        # glossy highlight streak across the top third
        hi_color = self._lighten(color, 55)
        self.create_polygon(
            pad+10, pad+6, self.width-pad-10, pad+6,
            self.width-pad-16, pad+16, pad+16, pad+16,
            fill=hi_color, outline="", smooth=True
        )
        # small pixel "sparkle" dot top-left
        self.create_rectangle(pad+8, pad+6, pad+12, pad+10, fill="#ffffff", outline="")
        self.create_text(self.width/2, self.height/2+2, text=self.text,
                          font=self.font, fill=self.fg)

    def _on_press(self, event):
        self._pressed = True
        self._draw(self._darken(self.color, 20))

    def _on_release(self, event):
        if self._pressed:
            self._draw(self.color)
            self._pressed = False
            if self.command:
                self.command()

    def _on_enter(self, event):
        if not self._pressed:
            self._draw(self._lighten(self.color, 15))

    def _on_leave(self, event):
        if not self._pressed:
            self._draw(self.color)


class DonutDataCleanerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Donut Data Cleaner")
        self.geometry("940x620")
        self.configure(bg=BG_MAIN)
        self.df_raw = None
        self.df_clean = None
        self.file_path = None
        self._build_layout()

    def _build_layout(self):
        header = tk.Frame(self, bg=BG_MAIN)
        header.pack(fill="x", pady=(14, 6))
        title_row = tk.Frame(header, bg=BG_MAIN)
        title_row.pack()
        self._draw_donut_glyph(title_row)
        tk.Label(title_row, text="DONUT DATA CLEANER", font=TITLE_FONT,
                 bg=BG_MAIN, fg=TEXT_DARK).pack(side="left", padx=10)
        tk.Label(header, text="clean up messy spreadsheets in one click",
                 font=(PIXEL_FONT_NAME, 10), bg=BG_MAIN, fg="#7a7a7a").pack()

        btn_frame = tk.Frame(self, bg=BG_MAIN)
        btn_frame.pack(fill="x", padx=24, pady=14)
        TexturedButton(btn_frame, "LOAD FILE", self.load_file, COLOR_BLUE).pack(side="left", padx=8)
        TexturedButton(btn_frame, "CLEAN DATA", self.clean_data, COLOR_PINK).pack(side="left", padx=8)
        TexturedButton(btn_frame, "EXPORT FILE", self.export_file, COLOR_PURPLE).pack(side="left", padx=8)
        TexturedButton(btn_frame, "RESET", self.reset_app, COLOR_ORANGE, width=110).pack(side="left", padx=8)

        self.status_label = tk.Label(self, text="No file loaded yet.",
                                      font=(PIXEL_FONT_NAME, 10, "bold"),
                                      bg=BG_MAIN, fg=TEXT_DARK)
        self.status_label.pack(pady=(0, 10))

        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook", background=BG_MAIN, borderwidth=0)
        style.configure("TNotebook.Tab", background=COLOR_YELLOW, foreground=TEXT_DARK,
                         font=(PIXEL_FONT_NAME, 10, "bold"), padding=[14, 8])
        style.map("TNotebook.Tab", background=[("selected", COLOR_PINK)],
                  foreground=[("selected", TEXT_LIGHT)])

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=24, pady=8)

        self.raw_frame = tk.Frame(self.notebook, bg=BG_PANEL)
        self.clean_frame = tk.Frame(self.notebook, bg=BG_PANEL)
        self.notebook.add(self.raw_frame, text="Raw Data")
        self.notebook.add(self.clean_frame, text="Cleaned Data")

        self.raw_tree = self._make_tree(self.raw_frame)
        self.clean_tree = self._make_tree(self.clean_frame)

        log_wrap = tk.Frame(self, bg=COLOR_BLUE, bd=0)
        log_wrap.pack(fill="x", padx=24, pady=(0, 16))
        self.log_box = tk.Text(log_wrap, height=6, bg="#f5faff", fg=TEXT_DARK,
                                font=LOG_FONT, relief="flat", bd=0,
                                highlightthickness=2, highlightbackground=OUTLINE)
        self.log_box.pack(fill="x", padx=3, pady=3)
        self._log("Ready. Load a CSV or Excel file to begin.")

    def _draw_donut_glyph(self, parent):
        c = tk.Canvas(parent, width=40, height=40, bg=BG_MAIN, highlightthickness=0)
        c.pack(side="left")
        c.create_oval(2, 2, 38, 38, fill=COLOR_ORANGE, outline=OUTLINE, width=3)
        c.create_oval(14, 14, 26, 26, fill=BG_MAIN, outline=OUTLINE, width=3)
        for x, y, col in [(10, 12, COLOR_PINK), (28, 10, COLOR_BLUE),
                           (8, 26, COLOR_YELLOW), (30, 27, COLOR_PURPLE)]:
            c.create_oval(x, y, x+4, y+4, fill=col, outline="")

    def _make_tree(self, parent):
        tree = ttk.Treeview(parent, show="headings")
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        return tree

    def _fill_tree(self, tree, df, max_rows=200):
        tree.delete(*tree.get_children())
        tree["columns"] = list(df.columns)
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="w")
        for _, row in df.head(max_rows).iterrows():
            tree.insert("", "end", values=list(row))

    def _log(self, msg):
        self.log_box.insert("end", f"> {msg}\n")
        self.log_box.see("end")

    # ---------- Core Functions ----------
    def load_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Spreadsheet files", "*.csv *.xlsx *.xls"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            if path.lower().endswith(".csv"):
                df = pd.read_csv(path)
            else:
                df = pd.read_excel(path)
        except Exception as e:
            messagebox.showerror("Load Error", f"Could not read file:\n{e}")
            return
        self.df_raw = df
        self.file_path = path
        self.status_label.config(text=f"Loaded: {os.path.basename(path)}  ({df.shape[0]} rows, {df.shape[1]} cols)")
        self._fill_tree(self.raw_tree, df)
        self._log(f"Loaded '{os.path.basename(path)}' with {df.shape[0]} rows and {df.shape[1]} columns.")

    @staticmethod
    def _snake_case(name):
        name = str(name).strip()
        name = re.sub(r"[^\w\s]", "", name)
        name = re.sub(r"\s+", "_", name)
        return name.lower()

    def clean_data(self):
        if self.df_raw is None:
            messagebox.showwarning("No Data", "Load a file first!")
            return

        df = self.df_raw.copy()
        n_before = len(df)

        df.columns = [self._snake_case(c) for c in df.columns]

        for col in df.select_dtypes(include="object").columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"": None, "nan": None, "None": None})

        df = df.drop_duplicates()
        n_after_dupes = len(df)

        for col in df.columns:
            sample = df[col].dropna().astype(str).head(20)
            if sample.empty:
                continue
            date_like = sample.str.match(r"^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}").mean() > 0.6
            if date_like:
                df[col] = pd.to_datetime(df[col], errors="coerce")
                continue
            numeric_like = pd.to_numeric(sample.str.replace(r"[$,%]", "", regex=True), errors="coerce").notna().mean() > 0.7
            if numeric_like:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(r"[$,%]", "", regex=True), errors="coerce")

        missing_report = df.isna().sum()
        missing_cols = missing_report[missing_report > 0]

        self.df_clean = df
        self._fill_tree(self.clean_tree, df)
        self.notebook.select(self.clean_frame)

        self._log(f"Removed {n_before - n_after_dupes} duplicate rows ({n_before} -> {n_after_dupes}).")
        self._log("Standardized column names to snake_case.")
        self._log("Trimmed whitespace, normalized empty strings to NaN.")
        self._log("Auto-detected & converted date/numeric columns where possible.")
        if not missing_cols.empty:
            self._log("Missing values found in: " + ", ".join(f"{c}({v})" for c, v in missing_cols.items()))
        else:
            self._log("No missing values detected. Data looks clean!")
        self.status_label.config(text=f"Cleaned data ready: {df.shape[0]} rows, {df.shape[1]} cols.")

    def export_file(self):
        if self.df_clean is None:
            messagebox.showwarning("No Cleaned Data", "Run 'Clean Data' first!")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV file", "*.csv"), ("Excel file", "*.xlsx")]
        )
        if not path:
            return
        try:
            if path.lower().endswith(".xlsx"):
                self.df_clean.to_excel(path, index=False)
            else:
                self.df_clean.to_csv(path, index=False)
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not save file:\n{e}")
            return
        self._log(f"Exported cleaned data to '{os.path.basename(path)}'")
        messagebox.showinfo("Export Complete", f"Saved cleaned file to:\n{path}")

    def reset_app(self):
        self.df_raw = None
        self.df_clean = None
        self.file_path = None
        self.raw_tree.delete(*self.raw_tree.get_children())
        self.clean_tree.delete(*self.clean_tree.get_children())
        self.status_label.config(text="No file loaded yet.")
        self.log_box.delete("1.0", "end")
        self._log("Reset. Load a new file to start again.")


if __name__ == "__main__":
    app = DonutDataCleanerApp()
    app.mainloop()
