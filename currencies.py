import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# CONFIG
INPUT_FROM_COLUMN = "From"
INPUT_TO_COLUMN = "To"
INPUT_RATE_COLUMN = "Rate"
INPUT_RATE_DATE_COLUMN = "Rate_Date"

OUTPUT_CURRENCYCODE_COLUMN = "CurrencyCode"
OUTPUT_EXCHANGERATE_COLUMN = "ExchangeRate"

SOURCE_CURRENCY = "CAD"


# FX Rate Filter Class
class FXRateFilter:
    def process_fxrate_file(self, upload_fxrate_file_path):
        with open(upload_fxrate_file_path, "r", encoding="utf-8-sig") as f:
            first_line = f.readline()
            sep = "\t" if "\t" in first_line else ","

        df = pd.read_csv(upload_fxrate_file_path, sep=sep, engine="python")

        required_columns = [
            INPUT_FROM_COLUMN,
            INPUT_TO_COLUMN,
            INPUT_RATE_COLUMN,
            INPUT_RATE_DATE_COLUMN
        ]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(
                "Missing required column(s): " + ", ".join(missing_columns)
            )

        df[INPUT_FROM_COLUMN] = df[INPUT_FROM_COLUMN].astype(str).str.upper().str.strip()
        df[INPUT_TO_COLUMN] = df[INPUT_TO_COLUMN].astype(str).str.upper().str.strip()

        # Keep only CAD rows from the From column
        cad_df = df[df[INPUT_FROM_COLUMN] == SOURCE_CURRENCY].copy()

        if cad_df.empty:
            raise ValueError("No rows found where From = CAD.")

        # Build export from To + Rate only
        export_df = cad_df[[INPUT_TO_COLUMN, INPUT_RATE_COLUMN]].copy()
        export_df = export_df.rename(columns={
            INPUT_TO_COLUMN: OUTPUT_CURRENCYCODE_COLUMN,
            INPUT_RATE_COLUMN: OUTPUT_EXCHANGERATE_COLUMN
        })

        export_df[OUTPUT_CURRENCYCODE_COLUMN] = (
            export_df[OUTPUT_CURRENCYCODE_COLUMN].astype(str).str.upper().str.strip()
        )

        export_df[OUTPUT_EXCHANGERATE_COLUMN] = pd.to_numeric(
            export_df[OUTPUT_EXCHANGERATE_COLUMN],
            errors="coerce"
        )

        # Remove duplicate currencies, keep first
        export_df = export_df.drop_duplicates(
            subset=[OUTPUT_CURRENCYCODE_COLUMN],
            keep="first"
        )

        # Force CAD = 1
        if "CAD" in export_df[OUTPUT_CURRENCYCODE_COLUMN].values:
            export_df.loc[
                export_df[OUTPUT_CURRENCYCODE_COLUMN] == "CAD",
                OUTPUT_EXCHANGERATE_COLUMN
            ] = 1
        else:
            cad_row = pd.DataFrame([{
                OUTPUT_CURRENCYCODE_COLUMN: "CAD",
                OUTPUT_EXCHANGERATE_COLUMN: 1
            }])
            export_df = pd.concat([export_df, cad_row], ignore_index=True)

        # Copy CNY rate to RMB
        cny_rows = export_df.loc[
            export_df[OUTPUT_CURRENCYCODE_COLUMN] == "CNY",
            OUTPUT_EXCHANGERATE_COLUMN
        ]

        if not cny_rows.empty and pd.notna(cny_rows.iloc[0]):
            cny_rate = cny_rows.iloc[0]

            if "RMB" in export_df[OUTPUT_CURRENCYCODE_COLUMN].values:
                export_df.loc[
                    export_df[OUTPUT_CURRENCYCODE_COLUMN] == "RMB",
                    OUTPUT_EXCHANGERATE_COLUMN
                ] = cny_rate
            else:
                rmb_row = pd.DataFrame([{
                    OUTPUT_CURRENCYCODE_COLUMN: "RMB",
                    OUTPUT_EXCHANGERATE_COLUMN: cny_rate
                }])
                export_df = pd.concat([export_df, rmb_row], ignore_index=True)

        # Sort output for cleaner file
        export_df = export_df.sort_values(
            by=OUTPUT_CURRENCYCODE_COLUMN
        ).reset_index(drop=True)

        return export_df

# APP
class FXRateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FX Rate Tool")
        self.root.geometry("760x210")
        self.root.minsize(700, 210)

        self.setup_dark_mode()
        self.fxrate_gui()

    def setup_dark_mode(self):
        bg = "#1e1e1e"
        panel = "#2a2a2a"
        field = "#252526"
        fg = "#f3f3f3"
        green = "#4CAF50"
        hover = "#3a3a3a"
        pressed = "#5a5a5a"
        disabled_green = "#6fae72"

        self.root.configure(bg=bg)

        style = ttk.Style()
        style.theme_use("clam")

        style.configure(".", background=bg, foreground=fg)
        style.configure("TFrame", background=bg)
        style.configure("TLabel", background=bg, foreground=fg, font=("Segoe UI", 9))

        style.configure(
            "TEntry",
            fieldbackground=field,
            foreground=green,
            insertcolor=green,
            padding=4,
            borderwidth=1
        )
        style.map(
            "TEntry",
            fieldbackground=[("readonly", field)],
            foreground=[("readonly", green)]
        )

        style.configure(
            "TButton",
            background=panel,
            foreground=fg,
            padding=(8, 4),
            borderwidth=1,
            font=("Segoe UI", 9)
        )
        style.map(
            "TButton",
            background=[
                ("active", hover),
                ("pressed", pressed)
            ],
            foreground=[
                ("active", fg),
                ("pressed", fg)
            ]
        )

        style.configure(
            "Accent.TButton",
            background=green,
            foreground="white",
            padding=(8, 5),
            font=("Segoe UI", 9, "bold")
        )
        style.map(
            "Accent.TButton",
            background=[
                ("disabled", disabled_green),
                ("active", "#45a049"),
                ("pressed", "#3e8e41")
            ],
            foreground=[
                ("disabled", "#d8d8d8"),
                ("active", "white"),
                ("pressed", "white")
            ]
        )


    def fxrate_gui(self):
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=0)

        ttk.Label(
            main,
            text="Upload FX Rate File Path:"
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        self.entry_file = ttk.Entry(main, state="readonly")
        self.entry_file.grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=(0, 8))

        ttk.Button(
            main,
            text="Upload",
            command=self.upload_fxrate_file,
            width=10
        ).grid(row=1, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(
            main,
            text="Save FX Rate File Path:"
        ).grid(row=2, column=0, sticky="w", pady=(4, 4))

        self.entry_save = ttk.Entry(main, state="readonly")
        self.entry_save.grid(row=3, column=0, sticky="ew", padx=(0, 8), pady=(0, 10))

        ttk.Button(
            main,
            text="Save",
            command=self.fxrate_file_save,
            width=10
        ).grid(row=3, column=1, sticky="ew", pady=(0, 10))

        self.run_button = ttk.Button(
            main,
            text="Run",
            command=self.fxrate_processing,
            state="disabled"
        )
        self.run_button.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(6, 0))

    def check_ready(self):
        file_path = self.entry_file.get()
        save_path = self.entry_save.get()

        if file_path and save_path:
            self.run_button.config(state="normal", style="Accent.TButton")
        else:
            self.run_button.config(state="disabled", style="TButton")

    def upload_fxrate_file(self):
        path = filedialog.askopenfilename(
            title="Select FX Rate File",
            filetypes=[("CSV / Text Files", "*.csv *.txt"), ("All Files", "*.*")]
        )
        if path:
            self.entry_file.config(state="normal")
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, path)
            self.entry_file.config(state="readonly")

            self.check_ready()

    def fxrate_file_save(self):
        path = filedialog.asksaveasfilename(
            title="Save output as",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if path:
            self.entry_save.config(state="normal")
            self.entry_save.delete(0, tk.END)
            self.entry_save.insert(0, path)
            self.entry_save.config(state="readonly")

            self.check_ready()

    def fxrate_processing(self):
        upload_fxrate_file_path = self.entry_file.get().strip()
        save_fxrate_file_path = self.entry_save.get().strip()

        if not upload_fxrate_file_path:
            messagebox.showwarning("Missing Upload", "Please select a file to upload.")
            return

        if not save_fxrate_file_path:
            messagebox.showwarning("Missing Save Path", "Please choose where to save the FX Rate file.")
            return

        try:
            processor = FXRateFilter()
            export_df = processor.process_fxrate_file(upload_fxrate_file_path)
            export_df.to_csv(save_fxrate_file_path, index=False)

            messagebox.showinfo(
                "Saved Successfully",
                f"FX Rate file was saved successfully.\n\nFile:\n{save_fxrate_file_path}"
            )
        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = FXRateApp(root)
    root.mainloop()