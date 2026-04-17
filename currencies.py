import csv
import json
import time
import getpass
from datetime import datetime
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Constants
# Uploaded file columns
INPUT_FROM_COLUMN = "From"
INPUT_TO_COLUMN = "To"
INPUT_RATE_COLUMN = "Rate"
INPUT_RATE_DATE_COLUMN = "Rate_Date"

# Output file columns
OUTPUT_CURRENCYCODE_COLUMN = "CurrencyCode"
OUTPUT_EXCHANGERATE_COLUMN = "ExchangeRate"

# Other constants
SOURCE_CURRENCY = "CAD"
CURRENCT_USER = getpass.getuser()

# FXRateTool_CurrencyConfig.json is a shared config file that contains the list of supported currencies and their mapping to the 'To' column values in the uploaded file
SHARED_CONFIG_PATH = r"C:\Users\andrea.clemente\Downloads\FXRateTool_CurrencyConfig.json"

# Load currency mapping from shared config file
CONFIG_PATH = Path(SHARED_CONFIG_PATH)
LOCK_FILE = CONFIG_PATH.with_suffix(CONFIG_PATH.suffix + ".lock")

# Export order
DEFAULT_CURRENCY_ORDER = [
    "GBP",
    "EUR",
    "USD",
    "CAD",
    "ARS",
    "AUD",
    "BGN",
    "BRL",
    "CHF",
    "CNY",
    "COP",
    "CZK",
    "DKK",
    "HKD",
    "HUF",
    "INR",
    "ISK",
    "JPY",
    "MXN",
    "MYR",
    "NOK",
    "NZD",
    "PHP",
    "PLN",
    "PEN",
    "RMB",
    "RON",
    "SEK",
    "SGD",
    "THB",
    "TWD",
    "UAH",
    "ZAR",
    "AED"
]

# Mapping of 'To' column values to CurrencyCode
CURRENCY_MAPPING_VALUES = {
    "GBP": "GBP",
    "EUR": "EUR",
    "USD": "USD",
    "ARS": "ARS",
    "AUD": "AUD",
    "BGN": "BGN",
    "BRL": "BRL",
    "CHF": "CHF",
    "CNY": "CNY",
    "COP": "COP",
    "CZK": "CZK",
    "DKK": "DKK",
    "HKD": "HKD",
    "HUF": "HUF",
    "INR": "INR",
    "ISK": "ISK",
    "JPY": "JPY",
    "MXN": "MXN",
    "MYR": "MYR",
    "NOK": "NOK",
    "NZD": "NZD",
    "PHP": "PHP",
    "PLN": "PLN",
    "PEN": "PEN",
    "RON": "RON",
    "SEK": "SEK",
    "SGD": "SGD",
    "THB": "THB",
    "TWD": "TWD",
    "UAH": "UAH",
    "ZAR": "ZAR",
    "AED": "AED"
}

# List of protected currencies that should not be deleted from the output even if they are not present in the uploaded file
PROTECTED_CURRENCIES = {
    "CAD",
    "RMB"
}

# Class for managing currency configuration
class CurrencyConfigManager:
    def __init__(self, config_path: Path, lock_file: Path):
        self.config_path = config_path
        self.lock_file = lock_file
        self.currencies = []
        self.mapping = {}
        self.history = []

        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        self.load_config()

    # Validation function to ensure the JSON structure is correct and contains the expected keys and types
    def validation_json_structure(self, data):
        if not isinstance(data, dict):
            raise ValueError("Invalid JSON structure for 'FXRateTool_CurrencyConfig.json': expected a dictionary.")
        for key in ["currencies", "mapping", "history"]:
            if key not in data:
                raise ValueError(f"Missing '{key}' key in JSON structure in 'FXRateTool_CurrencyConfig.json'.")
        if not isinstance(data["currencies"], list):
            raise ValueError("Invalid JSON structure for 'currencies' in 'FXRateTool_CurrencyConfig.json': expected a list.")
        if not isinstance(data["mapping"], dict):
            raise ValueError("Invalid JSON structure for 'mapping' in 'FXRateTool_CurrencyConfig.json': expected a dictionary.")
        if not isinstance(data["history"], list):
            raise ValueError("Invalid JSON structure for 'history' in 'FXRateTool_CurrencyConfig.json': expected a list.")

    # Load the currency configuration from the JSON file, validate its structure, and populate the currencies, mapping, and history attributes
    def load_config(self):
        if not self.config_path.exists():
            raise FileNotFoundError(f"Currency config file not found at:\n{self.config_path}. Please ensure the file exists and try again.")
        with open(self.config_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        self.validation_json_structure(data)
        self.currencies = data["currencies"]
        self.mapping = data["mapping"]
        self.history = data["history"]

        self.sync_mapping_with_currencies()

    # Save the current state of the currencies, mapping, and history back to the JSON file, including metadata about the last update such as the user and timestamp
    def save_config(self):
        data = {
            "currencies": self.currencies,
            "mapping": self.mapping,
            "history": self.history,
            "last_updated_user": CURRENCT_USER,
            "last_updated_timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)

    # Ensure that the mapping is always in sync with the currencies list
    def sync_mapping_with_currencies(self):
        for currency in self.currencies:
            if currency not in self.mapping and currency not in {"CAD", "RMB"}:
                self.mapping[currency] = currency

        for key in list(self.mapping.keys()):
            if key not in self.currencies:
                del self.mapping[key]

    # Add a new currency to the list at the specified index, update the mapping, and log the action in the history with details about the user and timestamp
    def add_currency(self, code, insert_index):
        self.currencies.insert(insert_index, code)

        if code not in {"CAD", "RMB"}:
            self.mapping[code] = code

        self.history.append({
            "action": "ADD",
            "currency": code,
            "user": CURRENCT_USER,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "details": f"Added '{code}' at position {insert_index + 1}"
        })
        self.save_config()

    # Remove the specified currencies from the list, update the mapping accordingly, and log each removal action in the history with details about the user and timestamp
    # Protected currencies will be skipped and not removed
    def remove_currency(self, codes):
        removed_currencies = []
        for code in codes:
            if code in PROTECTED_CURRENCIES:
                continue
            if code in self.currencies:
                self.currencies.remove(code)
                removed_currencies.append(code)
                if code in self.mapping:
                    del self.mapping[code]

                self.history.append({
                    "action": "REMOVE",
                    "currency": code,
                    "user": CURRENCT_USER,
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "details": f"Removed '{code}' from the currency list"
                })
        self.save_config()
        return removed_currencies
    
    # Move the specified currencies up in the list
    def move_currency_up(self, code):
        if code not in self.currencies:
            raise ValueError(f"Currency '{code}' not found in the list.")
        
        selected_indices = sorted([self.currencies.index(c) for c in code if c in self.currencies])

        if not selected_indices:
            return
        
        if selected_indices[0] == 0:
            return
        
        for idx in selected_indices:
            self.currencies[idx - 1], self.currencies[idx] = self.currencies[idx], self.currencies[idx - 1]

        self.history.append({
            "action": "MOVE UP",
            "currency": code,
            "user": CURRENCT_USER,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "details": f"Moved '{', '.join(code)}' up to position(s) {', '.join(str(i + 1) for i in selected_indices)}"
        })

        self.save_config()

    # Move the specified currencies down in the list
    def move_currency_down(self, code):
        if code not in self.currencies:
            raise ValueError(f"Currency '{code}' not found in the list.")
        
        selected_indices = sorted([self.currencies.index(c) for c in code if c in self.currencies], reverse=True)

        if not selected_indices:
            return

        if selected_indices[0] == len(self.currencies) - 1:
            return
        
        for idx in selected_indices:
             self.currencies[idx + 1], self.currencies[idx] = self.currencies[idx], self.currencies[idx + 1]

        self.history.append({
            "action": "MOVE DOWN",
            "currency": code,
            "user": CURRENCT_USER,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "details": f"Moved '{', '.join(code)}' down to position(s) {', '.join(str(i + 1) for i in selected_indices)}"
        })

        self.save_config()

# Class for FX Rate File Processing
class FXRateProcessing:
    def __init__(self, currencies, mapping):
        self.currencies = currencies
        self.mapping = mapping

    # Process the uploaded FX rate file, filter and transform the data according to the specified rules, validate the results, and prepare a DataFrame for export
    def process_fxrate_file(self, upload_fxrate_file_path):
        df = pd.read_excel(upload_fxrate_file_path, engine="openpyxl")

        df = df[df[INPUT_FROM_COLUMN].str.upper() == SOURCE_CURRENCY]

        export_df = df[[INPUT_TO_COLUMN, INPUT_RATE_COLUMN]].copy()

        # Map 'To' to CurrencyCode and keep Rate as ExchangeRate
        export_df[OUTPUT_CURRENCYCODE_COLUMN] = (
            export_df[INPUT_TO_COLUMN]
            .astype(str)
            .str.strip()
            .str.upper()
            .map(lambda x: self.mapping.get(x,x))
        )
        # Validate that all mapped currency codes are in the correct format (3 uppercase letters)
        invalid_currencies = export_df[
            ~export_df[OUTPUT_CURRENCYCODE_COLUMN].astype(str).str.fullmatch(r"[A-Z]{3}")
        ]

        if not invalid_currencies.empty:
            invalid_list = invalid_currencies[OUTPUT_CURRENCYCODE_COLUMN].tolist()
            raise ValueError(
                f"Invalid currency codes found in the uploaded file for the following currencies: {', '.join(map(str, invalid_list))}. "
                "\nCurrency codes must be exactly 3 letters with no numbers or special characters. Please correct the currency codes and try again."
            )

        # Keep only the mapped CurrencyCode and ExchangeRate columns and rename them to the output format
        export_df = export_df[[OUTPUT_CURRENCYCODE_COLUMN, INPUT_RATE_COLUMN]]
        export_df = export_df.rename(columns={
            INPUT_TO_COLUMN: OUTPUT_CURRENCYCODE_COLUMN,
            INPUT_RATE_COLUMN: OUTPUT_EXCHANGERATE_COLUMN
        }, inplace=True)

        # Keep exact rate text as is without any formatting changes to ensure the integrity of the data, especially for very small or very large numbers
        export_df[OUTPUT_CURRENCYCODE_COLUMN] = (
            export_df[OUTPUT_CURRENCYCODE_COLUMN].astype(str)
        )

        # Validate duplicates in the CurrencyCode column after mapping
        duplicate_currencies = export_df[export_df.duplicated(subset=[OUTPUT_CURRENCYCODE_COLUMN], keep=False)]
        if not duplicate_currencies.empty:
            duplicates_currencyList = duplicate_currencies[OUTPUT_CURRENCYCODE_COLUMN].tolist()
            raise ValueError(
                f"Duplicate currency codes found in the uploaded file for the following currencies: {', '.join(duplicates_currencyList)}. "
                "\nPlease ensure there are no duplicates and try again."
            )

        # Add CAD manually and force CAD = 1
        export_df = pd.concat([
            export_df,
            pd.DataFrame([{
                OUTPUT_CURRENCYCODE_COLUMN: "CAD",
                OUTPUT_EXCHANGERATE_COLUMN: "1"
            }])
        ], ignore_index=True)

        # Add RMB manually, copy CNY rate to RMB
        cny_row = export_df.loc[export_df[OUTPUT_CURRENCYCODE_COLUMN] == "CNY"]
        if not cny_row.empty:
            rmb_rate = cny_row.iloc[0][OUTPUT_EXCHANGERATE_COLUMN]
            export_df = pd.concat([
                export_df,
                pd.DataFrame([{
                    OUTPUT_CURRENCYCODE_COLUMN: "RMB",
                    OUTPUT_EXCHANGERATE_COLUMN: rmb_rate
                }])
            ], ignore_index=True)

        # Apply exact export order
        export_df["_sort_order"] = pd.Categorical(
            export_df[OUTPUT_CURRENCYCODE_COLUMN],
            categories=DEFAULT_CURRENCY_ORDER,
            ordered=True
        )
        # For any currencies not in the default order, sort them
        export_df["_sort_order"] = pd.Categorical(
            export_df[OUTPUT_CURRENCYCODE_COLUMN],
            categories=self.currencies,
            ordered=True
        )

        # Sort by the defined order and then drop the sort order column
        export_df = export_df.sort_values("_sort_order").drop(columns="_sort_order")

        # Reset index for clean export
        export_df = export_df.reset_index(drop=True)

        # Return the final DataFrame ready for export
        return export_df

# FX Rate APP
class FXRateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FX Rate Processing Tool")
        self.root.geometry("800x300")
        self.root.minsize(760, 320)

        self.shared_config_manager = CurrencyConfigManager(CONFIG_PATH, LOCK_FILE)

        # Window/widget references
        self.manage_window = None
        self.currency_listbox = None
        self.history = None
        self.currency_entry_widget = None

        # Sorting for History tab
        self.history_sort_column = None
        self.history_sort_reverse = False

        # Build UI
        self.setup_dark_mode()
        self.fxrate_gui()

    # Setup dark mode styles for the application using ttk.Style and configuring colors for various widgets to create a cohesive dark theme throughout the app
    def setup_dark_mode(self):
        bg = "#2b2b2b"
        panel = "#6e6e6e"
        field = "#505050"
        fg = "#f3f3f3"
        green = "#4CAF50"
        hover = "#5a5a5a"
        pressed = "#3a3a3a"

        self.root.configure(bg=bg)

        style = ttk.Style()
        style.theme_use("clam")

        # General styles for the main window, frames, and labels
        style.configure(
            ".",
            background=bg,
            foreground=fg
        )

        # Frame/Label styles for entire application
        style.configure(
            "TFrame",
            background=bg
        )

        # Label style for all labels for entire application
        style.configure(
            "TLabel",
            background=bg,
            foreground=fg,
            font=("TkDefaultFont", 9, "bold")
        )

        # Title label at the top of the main window
        style.configure(
            "Title.TLabel",
            font=("TkDefaultFont", 12, "bold"),
            foreground="white",
            background="#2b2b2b"
        )

        # Entry field style for upload and save paths, with specific colors for normal and readonly states
        style.configure(
            "TEntry",
            fieldbackground=field,
            foreground=green,
            insertcolor=green,
            padding=4,
            borderwidth=1
        )

        # Entry fields style for upload and save paths in readonly state
        style.map(
            "TEntry",
            fieldbackground=[("readonly", field)],
            foreground=[("readonly", green)]
        )

        # General button style for all buttons, with specific styles for the Run and Exit buttons
        style.configure(
            "TButton",
            background=panel,
            foreground=fg,
            padding=(8, 4),
            borderwidth=1,
            font=("TkDefaultFont", 9, "bold"),
            focuscolor=panel,
            focusthickness=0
        )

        # All buttons style except Run and Exit buttons with hover and pressed states
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

        # Run button style
        style.configure(
            "Run.TButton",
            background=green,
            foreground="white",
            padding=(8, 5),
            font=("TkDefaultFont", 9, "bold"),
            focuscolor=green,
            focusthickness=0
        )

        # Run button style with hover and pressed states
        style.map(
            "Run.TButton",
            background=[
                ("disabled", "#FF6F6F"),
                ("active", "#45a049"),
                ("pressed", "#2c812f")
            ],
            foreground=[
                ("disabled", "#dbd9d9"),
                ("active", "white"),
                ("pressed", "white")
            ]
        )

        # Clear button style
        style.configure(
            "Clear.TButton",
            font=("TkDefaultFont", 9, "bold"),
            padding=(10, 6),
            focuscolor=panel,
            focusthickness=0
        )

        # Clear button style with hover and pressed states
        style.map(
            "Clear.TButton",
            background=[
                ("active", "#dab955"),
                ("pressed", field),
                ("!active", panel)
            ],
            foreground=[
                ("active", "white"),
                ("pressed", "white"),
                ("!active", fg)
            ]
        )

        # Header style for the currency list and event history in the Manage Currencies window
        style.configure(
            "Treeview",
            background=field,
            foreground=fg,
            fieldbackground=field
        )

        # Header style for Event History
        style.configure(
            "Treeview.Heading",
            background=panel,
            foreground=fg,
            font=("TkDefaultFont", 9, "bold")
        )

        # Map for Event History header
        style.map(
            "Treeview.Heading",
            background=[
                ("active", bg),
                ("pressed", field)
            ],
            foreground=[
                ("active", fg),
                ("pressed", fg)
            ]   
        )

        # Notebook style for tabs in the Manage Currencies window
        style.configure(
            "TNotebook",
            background=bg,
            borderwidth=0
        )

        # Configure the Notebook tabs style
        style.configure(
            "TNotebook.Tab",
            padding=(10, 5),
            font=("TkDefaultFont", 9, "bold"),
            background=bg,
            foreground=panel,
        )

        # Map for tabs in the Manage Currencies window
        style.map(
            "TNotebook.Tab",
            background=[
                ("selected", bg),
                ("active", field)
            ],
            foreground=[
                ("selected", fg),
                ("active", fg)
            ]
        )

        # Layout configuration for the Notebook tabs in the Manage Currencies window
        style.layout("TNotebook.Tab", [
            ("Notebook.tab", {
                "sticky": "nswe",
                "children": [
                    ("Notebook.padding", {
                        "side": "top",
                        "sticky": "nswe",
                        "children": [
                            ("Notebook.label", {"side": "top", "sticky": ""})
                        ]
                    })
                ]
            })
        ])

    # Build the main GUI layout for the FX Rate Tool, including the title, upload and save sections, run button, and manage currencies and exit buttons, all styled with the defined dark mode theme
    def fxrate_gui(self):
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        # Center everything in the main frame
        main.columnconfigure(0, weight=1)

        top_menu_frame = tk.Frame(main, bg="#2b2b2b")
        top_menu_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        menu_button = tk.Menubutton(
            top_menu_frame,
            text="Options ▼",
            bg="#2b2b2b",
            fg="#f3f3f3",
            activebackground="#3a3a3a",
            activeforeground="#f3f3f3",
            relief="flat",
            bd=0,
            padx=10,
            pady=4
        )
        menu_button.pack(side="right")

        menu = tk.Menu(
            menu_button,
            tearoff=0,
            bg="#2b2b2b",
            fg="#f3f3f3",
            activebackground="#3a3a3a",
            activeforeground="#f3f3f3",
            bd=0
        )

        menu.add_command(label="Manage Currencies", command=self.open_manage_currencies_window)
        menu.add_command(label="Exit", command=self.close_main_window)
        menu_button.config(menu=menu)

        # Title Label/Text
        ttk.Label(
            main,
            text="\nFX Rate Processing Tool",
            style="Title.TLabel"
        ).grid(row=1, column=0, pady=(0, 15))

        # Frame for upload/save/run buttons and entries
        action_frame = tk.Frame(
            main,
            bg="#2b2b2b",
            relief="groove",
            bd=2,
            padx=12,
            pady=12
        )
        action_frame.grid(row=2, column=0, sticky="ew", pady=(0, 15))

        action_frame.grid_columnconfigure(0, weight=1)

        # Row for Upload section
        upload_row = tk.Frame(action_frame, bg="#2b2b2b")
        upload_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        upload_row.grid_columnconfigure(1, weight=1)

        self.entry_file = tk.Entry(
            upload_row,
            bg="#505050",
            fg="#ff6b6b",
            readonlybackground="#505050",
            insertbackground="#ffffff",
            relief="flat",
            bd=0
        )
        self.entry_file.insert(0, "Please select a file to upload...")
        self.entry_file.config(state="readonly")
        self.entry_file.grid(row=0, column=1, sticky="ew")

        # Upload Button
        ttk.Button(
            upload_row,
            text="Upload",
            command=self.upload_fxrate_file,
            width=10
        ).grid(row=0, column=2, padx=(10, 0))

        save_row = tk.Frame(action_frame, bg="#2b2b2b")
        save_row.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        save_row.grid_columnconfigure(1, weight=1)

        self.entry_save = tk.Entry(
            save_row,
            bg="#505050",
            fg="#ff6b6b",
            readonlybackground="#505050",
            insertbackground="#ffffff",
            relief="flat",
            bd=0
        )
        self.entry_save.insert(0, "Please select save location for the FX Rate csv file...")
        self.entry_save.config(state="readonly")
        self.entry_save.grid(row=0, column=1, sticky="ew")

        # Save Button
        ttk.Button(
            save_row,
            text="Save",
            command=self.fxrate_file_save,
            width=10
        ).grid(row=0, column=2, padx=(10, 0))

        # Manage Currencies and Clear buttons at the bottom
        clear_row = tk.Frame(action_frame, bg="#2b2b2b")
        clear_row.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        clear_row.grid_columnconfigure(1, weight=1)

        # Clear Button
        ttk.Button(
            clear_row,
            text="Clear",
            command=self.clear_entries,
            style="Clear.TButton",
            width=10
        ).grid(row=0, column=2, padx=(10, 0))

        run_row = tk.Frame(action_frame, bg="#2b2b2b")
        run_row.grid(row=2, column=0, sticky="ew", pady=(8, 0))

        # Run Button
        self.run_button = ttk.Button(
            run_row,
            text="RUN",
            command=self.fxrate_processing,
            state="disabled",
            width=10,
            style="Run.TButton"
        )
        self.run_button.pack(anchor="center")

    # Refresh the currency configuration by reloading the JSON file, updating the views, and showing a summary of changes including added, removed
    # and order changes in the currency list, with appropriate confirmation and error handling
    def refresh_currency_config(self):
        parent=self.manage_window if self.manage_window and self.manage_window.winfo_exists() else self.root
        confirm = messagebox.askyesno(
            "Confirm Refresh",
            "Do you want to refresh the currency configuration 'FXRateTool_CurrencyConfig.json' file?",
            parent=parent
        )
        if not confirm:
            return

        try:
            old_currencies = list(self.shared_config_manager.currencies)
            self.shared_config_manager.load_config()
            new_currencies = list(self.shared_config_manager.currencies)

            if self.manage_window and self.manage_window.winfo_exists():
                self.refresh_manage_currency_views()

            added_currencies = [c for c in new_currencies if c not in old_currencies]
            removed_currencies = [c for c in old_currencies if c not in new_currencies]
            order_changed = old_currencies != new_currencies and not added_currencies and not removed_currencies

            if not added_currencies and not removed_currencies and not order_changed:
                messagebox.showinfo(
                    "No Changes Found",
                    "The currency configuration is already up to date.\n\nNo changes were detected.",
                    parent=parent
                )
                return

            message = "Currency configuration has been refreshed successfully.\n\n"
            if added_currencies:
                message += f"Added Currency:\n" + "\n".join(added_currencies) + "\n\n"
            if removed_currencies:
                message += f"Removed Currency:\n" + "\n".join(removed_currencies) + "\n\n"
            if order_changed:
                message += "The currency order was updated.\n\n"

            messagebox.showinfo(
                "Refresh Complete",
                message.strip(),
                parent=parent
            )
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"An error occurred while refreshing the currency configuration:\n{str(e)}",
                parent=parent
            )

    # Close the manage currencies window if it is open, with appropriate checks to ensure the window exists before attempting to close it
    def close_manage_window(self):
        if self.manage_window and self.manage_window.winfo_exists():
            self.manage_window.destroy()
            self.manage_window = None

    # Close main application window with confirmation from user
    def close_main_window(self):
        confirm = messagebox.askyesno(
            "Confirm Exit",
            "Are you sure you want to exit the FX Rate Processing Tool?",
            parent=self.root
        )
        if confirm:
            self.root.destroy()
        if not confirm:
            return
    
    # Clear upload/save paths
    def clear_entries(self):
        confirm_clear = messagebox.askyesno(
            "Confirm Clear",
            "Are you sure you want to clear the selected file and save paths?",
            parent=self.root
        )
        if not confirm_clear:
            return

        self.entry_file.config(state="normal", fg="#ff6b6b")
        self.entry_file.delete(0, tk.END)
        self.entry_file.insert(0, "Please select a file to upload...")
        self.entry_file.config(state="readonly")

        self.entry_save.config(state="normal", fg="#ff6b6b")
        self.entry_save.delete(0, tk.END)
        self.entry_save.insert(0, "Please select save location for the FX Rate csv file...")
        self.entry_save.config(state="readonly")
        self.check_ready()

    # Check if both upload/save files are selected and enable the Run button if they are, otherwise keep it disabled to prevent processing without the necessary inputs
    def check_ready(self):
        file_path = self.entry_file.get().strip()
        save_path = self.entry_save.get().strip()

        upload_placeholder = "Please select a file to upload..."
        save_placeholder = "Please select save location for the FX Rate csv file..."

        is_ready = (
            file_path
            and save_path
            and file_path != upload_placeholder
            and save_path != save_placeholder
        )

        if is_ready:
            self.run_button.config(state="normal", style="Run.TButton")
        else:
            self.run_button.config(state="disabled", style="Run.TButton")

    # Open a file dialog to select the FX Rate DTS to RGS file, update the entry field with the selected file path, and check if both upload and save paths are ready to enable the Run button
    def upload_fxrate_file(self):
        upload_path = filedialog.askopenfilename(
            title="Select FX Rate DTS to RGS File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if upload_path:
            self.entry_file.config(state="normal", fg="#4CAF50")
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, upload_path)
            self.entry_file.config(state="readonly")
            self.check_ready()

    # File dialog to select where to save file, update the entry field with the selected save path, and check if both upload and save paths are ready to enable the Run button
    def fxrate_file_save(self):
        save_path = filedialog.asksaveasfilename(
            title="Save FX Rate File",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if save_path:
            self.entry_save.config(state="normal", fg="#4CAF50")
            self.entry_save.delete(0, tk.END)
            self.entry_save.insert(0, save_path)
            self.entry_save.config(state="readonly")

            self.check_ready()

    # Processing uploaded file with validation inputs
    def fxrate_processing(self):
        upload_fxrate_file_path = self.entry_file.get().strip()
        save_fxrate_file_path = self.entry_save.get().strip()

        if not upload_fxrate_file_path:
            messagebox.showwarning(
                "Missing Upload",
                "Please select a file to upload.",
                parent=self.root
            )
            return

        if not save_fxrate_file_path:
            messagebox.showwarning(
                "Missing Save Path",
                "Please choose where to save the FX Rate csv file.",
                parent=self.root
            )
            return

        try:
            self.shared_config_manager.load_config()
            processor = FXRateProcessing(
                self.shared_config_manager.currencies,
                self.shared_config_manager.mapping
            )
            export_df = processor.process_fxrate_file(upload_fxrate_file_path)
            export_df.to_csv(save_fxrate_file_path, index=False, quoting=csv.QUOTE_ALL)

            messagebox.showinfo(
                "Saved Successfully",
                f"FX Rate csv file was saved successfully.\n\nFile:\n{save_fxrate_file_path}",
                parent=self.root
            )
        except Exception as e:
            messagebox.showerror(
                "Error",
                str(e),
                parent=self.root
            )

    # Open Manage Currencies window for managing currencies
    def open_manage_currencies_window(self):
        if self.manage_window and self.manage_window.winfo_exists():
            self.manage_window.lift()
            self.manage_window.focus_force()
            return

        self.shared_config_manager.load_config()

        self.manage_window = tk.Toplevel(self.root)
        self.manage_window.title("Manage Currencies")
        self.manage_window.geometry("760x500")
        self.manage_window.minsize(700, 460)
        self.manage_window.configure(bg="#2b2b2b")

        top_menu_bar = tk.Frame(self.manage_window, bg="#2b2b2b")
        top_menu_bar.pack(fill="x", padx=10, pady=(5, 0))

        menu_button = tk.Menubutton(
            top_menu_bar,
            text="Options ▼",
            bg="#2b2b2b",
            fg="#f3f3f3",
            activebackground="#3a3a3a",
            activeforeground="#f3f3f3",
            relief="flat",
            bd=0,
            padx=10,
            pady=4
        )
        menu_button.pack(side="right")

        menu = tk.Menu(
            menu_button,
            tearoff=0,
            bg="#2b2b2b",
            fg="#f3f3f3",
            activebackground="#3a3a3a",
            activeforeground="#f3f3f3",
            bd=0
        )

        menu.add_command(label="Refresh Currency Config", command=self.refresh_currency_config)
        menu.add_command(label="Exit", command=self.close_manage_window)
        menu_button.config(menu=menu)

        notebook = ttk.Notebook(self.manage_window)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        currencies_tab = ttk.Frame(notebook, padding=12)
        history_tab = ttk.Frame(notebook, padding=12)

        notebook.add(currencies_tab, text="Currencies")
        notebook.add(history_tab, text="Event History")

        currency_list_frame = tk.Frame(currencies_tab, bg="#2b2b2b")
        currency_list_frame.pack(fill="both", expand=True, pady=(0, 10))

        self.currency_listbox = tk.Listbox(
            currency_list_frame,
            height=14,
            selectmode=tk.EXTENDED,
            bg="#505050",
            fg="#e0e0e0",
            selectbackground="#2e7d32",
            selectforeground="#ffffff",
            relief="flat",
            highlightthickness=0,
            activestyle="none",
            bd=0,
            exportselection=False
        )
        self.currency_listbox.pack(side="left", fill="both", expand=True)

        currency_scrollbar = tk.Scrollbar(
            currency_list_frame,
            orient="vertical",
            command=self.currency_listbox.yview
        )
        currency_scrollbar.pack(side="right", fill="y")
        self.currency_listbox.config(yscrollcommand=currency_scrollbar.set)

        entry_frame = ttk.Frame(currencies_tab)
        entry_frame.pack(fill="x", pady=(0, 10))

        self.currency_entry_widget = ttk.Entry(entry_frame)
        self.currency_entry_widget.pack(side="left", fill="x", expand=True, padx=(0, 8))

        # Add Currency Button
        def addCurrency():
            code = self.currency_entry_widget.get().strip().upper()
            if not code:
                messagebox.showwarning(
                    "Invalid Currency",
                    "Please enter a valid currency code.",
                    parent=self.manage_window
                )
                return

            if len(code) != 3 or not code.isalpha():
                messagebox.showwarning(
                    "Invalid Format",
                    "Currency code must be exactly 3 letters with no numbers or special characters (Example: USD, EUR, GBP).",
                    parent=self.manage_window
                )
                return
            
            if code in self.shared_config_manager.currencies:
                messagebox.showwarning(
                    "Currency Exists",
                    f"Currency '{code}' already exists in the list.",
                    parent=self.manage_window
                )
                return
            
            insert_index = len(self.shared_config_manager.currencies)

            confirm_currency_message = (
                f"Are you sure you want to add '{code}' to the currency list at the end?\n\n"
                f"NOTE: This will update the 'FXRateTool_CurrencyConfig.json' file and add it to the currency mapping '{code}': '{code}'."
            )
            confirm = messagebox.askyesno(
                "Confirm Add Currency",
                confirm_currency_message, parent=self.manage_window
            )
            if not confirm:
                return

            try:
                self.shared_config_manager.load_config()
                self.shared_config_manager.add_currency(code, insert_index=insert_index)
                self.currency_entry_widget.delete(0, tk.END)
                self.refresh_manage_currency_views()

                for i, currency in enumerate(self.shared_config_manager.currencies):
                    if currency == code:
                        self.currency_listbox.selection_clear(0, tk.END)
                        self.currency_listbox.selection_set(i)
                        self.currency_listbox.activate(i)
                        break

                messagebox.showinfo(
                    "Currency Added",
                    f"Currency '{code}' has been added and mapped successfully.",
                    parent=self.manage_window
                )
            except Exception as e:
                messagebox.showerror(
                    "Error",
                    f"An error occurred while adding the currency:\n{str(e)}",
                    parent=self.manage_window
                )

        # Remove Currency Button
        def removeCurrency():
            selected_indices = self.currency_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning(
                    "No Selection",
                    "Please select at least one currency to remove.",
                    parent=self.manage_window
                )
                return

            selected_codes = [self.currency_listbox.get(i) for i in selected_indices]
            protected_currencies_selected = [code for code in selected_codes if code in PROTECTED_CURRENCIES]
            removable_currencies = [code for code in selected_codes if code not in PROTECTED_CURRENCIES]

            if not removable_currencies:
                messagebox.showwarning(
                    "Protected Currencies",
                    f"The selected currencies cannot be removed because they are protected and must always be included in the output:\n{', '.join(protected_currencies_selected)}",
                    parent=self.manage_window
                )
                return

            confirm_currency_message = (
                f"Are you sure you want to remove '{', '.join(removable_currencies)}' from the currency list?\n\n"
                + "NOTE: This will update the 'FXRateTool_CurrencyConfig.json' file and remove from the currency mapping."
            )
            if protected_currencies_selected:
                confirm_currency_message += "\n\nNOTE: The following selected currencies are protected and will not be removed:\n"
                confirm_currency_message += "\n".join(protected_currencies_selected)

            confirm = messagebox.askyesno(
                "Confirm Remove Currencies",
                confirm_currency_message,
                parent=self.manage_window
            )
            if not confirm:
                return

            try:
                self.shared_config_manager.load_config()
                removed_currencies = self.shared_config_manager.remove_currency(removable_currencies)
                self.refresh_manage_currency_views()

                if removed_currencies:
                    messagebox.showinfo(
                        "Currencies Removed",
                        f"Currency '{', '.join(removed_currencies)}' has been removed successfully.",
                        parent=self.manage_window
                    )
            except Exception as e:
                messagebox.showerror(
                    "Error",
                    f"An error occurred while removing the currencies:\n{str(e)}",
                    parent=self.manage_window
                )

        # Move Up Button
        def move_up():
            selected_indices = self.currency_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning(
                    "No Selection",
                    "Please select at least one currency to move.",
                    parent=self.manage_window
                )
                return

            codes = [self.currency_listbox.get(i) for i in selected_indices]

            try:
                self.shared_config_manager.load_config()
                self.shared_config_manager.move_currency_up(codes)
                self.refresh_manage_currency_views()
                self.currency_listbox.selection_clear(0, tk.END)

                for i, currency in enumerate(self.shared_config_manager.currencies):
                    if currency in codes:
                        self.currency_listbox.selection_set(i)
                        self.currency_listbox.activate(i)
            except Exception as e:
                messagebox.showerror(
                    "Error",
                    f"An error occurred while moving the currency up:\n{str(e)}",
                    parent=self.manage_window
                )

        # Move Down Button
        def move_down():
            selected_indices = self.currency_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning(
                    "No Selection",
                    "Please select at least one currency to move.",
                    parent=self.manage_window
                )
                return

            codes = [self.currency_listbox.get(i) for i in selected_indices]

            try:
                self.shared_config_manager.load_config()
                self.shared_config_manager.move_currency_down(codes)
                self.refresh_manage_currency_views()
                self.currency_listbox.selection_clear(0, tk.END)

                for i, currency in enumerate(self.shared_config_manager.currencies):
                    if currency in codes:
                        self.currency_listbox.selection_set(i)
                        self.currency_listbox.activate(i)
            except Exception as e:
                messagebox.showerror(
                    "Error",
                    f"An error occurred while moving the currency down:\n{str(e)}",
                    parent=self.manage_window
                )

        button_frame = ttk.Frame(currencies_tab)
        button_frame.pack(fill="x")

        left_button_frame = ttk.Frame(button_frame)
        left_button_frame.pack(side="left")

        ttk.Button(
            left_button_frame,
            text="Add ✚",
            command=addCurrency
        ).pack(side="left")

        ttk.Button(
            left_button_frame,
            text="Remove ✖",
            command=removeCurrency
        ).pack(side="left", padx=(8, 0))

        right_button_frame = ttk.Frame(button_frame)
        right_button_frame.pack(side="right")

        ttk.Button(
            right_button_frame,
            text="Move Up ▲",
            command=move_up
        ).pack(side="right", padx=(8, 0))

        ttk.Button(
            right_button_frame,
            text="Move Down ▼",
            command=move_down
        ).pack(side="right", padx=(8, 0))

        right_frame = ttk.Frame(button_frame)
        right_frame.pack(side="right")

        columns = ("time", "action", "currency", "user", "details")
        self.history = ttk.Treeview(
            history_tab,
            columns=columns,
            show="headings"
        )

        # Configure headings with sorting functionality
        self.history.heading("time", text="Time", command=lambda: self.sort_history("time"))
        self.history.heading("action", text="Action", command=lambda: self.sort_history("action"))
        self.history.heading("currency", text="Currency", command=lambda: self.sort_history("currency"))
        self.history.heading("user", text="User", command=lambda: self.sort_history("user"))
        self.history.heading("details", text="Details", command=lambda: self.sort_history("details"))

        # Configure column widths and alignment
        self.history.column("time", width=180)
        self.history.column("action", width=100, anchor="center")
        self.history.column("currency", width=90, anchor="center")
        self.history.column("user", width=120, anchor="center")
        self.history.column("details", width=220)

        scrollbar = ttk.Scrollbar(history_tab, orient="vertical", command=self.history.yview) # Add vertical scrollbar to the history Treeview
        self.history.configure(yscrollcommand=scrollbar.set) # Link the scrollbar to the Treeview
        self.history.pack(side="left", fill="both", expand=True) # Pack the Treeview to the left and allow it to expand, filling available space
        scrollbar.pack(side="right", fill="y") # Pack the scrollbar to the right of the history tab and fill vertically

        self.refresh_manage_currency_views()

    # Refresh the currency listbox/event history by reloading Configuration and updating accordingly
    def refresh_manage_currency_views(self):
        self.shared_config_manager.load_config()

        if self.currency_listbox and self.currency_listbox.winfo_exists():
            selected_values = [
                self.currency_listbox.get(i) for i in self.currency_listbox.curselection()
            ]

            self.currency_listbox.delete(0, tk.END)

            for currency in self.shared_config_manager.currencies:
                self.currency_listbox.insert(tk.END, currency)

            self.currency_listbox.selection_clear(0, tk.END)
            for i, currency in enumerate(self.shared_config_manager.currencies):
                if currency in selected_values:
                    self.currency_listbox.selection_set(i)

        if self.history and self.history.winfo_exists():
            for item in self.history.get_children():
                self.history.delete(item)

            for event in reversed(self.shared_config_manager.history):
                self.history.insert(
                    "",
                    "end",
                    values=(
                        event.get("timestamp", ""),
                        event.get("action", ""),
                        event.get("currency", ""),
                        event.get("user", ""),
                        event.get("details", "")
                    )
                )

    # Sort the event historty in History tab by specified column
    def sort_history(self, column_name):
        # Check if history Treeview exists before attempting to sort
        if not self.history or not self.history.winfo_exists():
            return

        rows = [] # List to hold the rows of the Treeview for sorting
        # Retrieve all rows from the Treeview and store them in a list for sorting
        for item_id in self.history.get_children():
            values = self.history.item(item_id, "values")
            rows.append(values)

        # Determine the index of the column to sort by
        if self.history_sort_column == column_name:
            self.history_sort_reverse = not self.history_sort_reverse
        else:
            self.history_sort_column = column_name
            self.history_sort_reverse = True if column_name == "time" else False

        # Update the column headings to indicate the current sort column and direction with an arrow symbol
        self.history.heading(
            "time",
            text="Time",
            command=lambda: self.sort_history("time")
        )
        self.history.heading(
            "action",
            text="Action",
            command=lambda: self.sort_history("action")
        )
        self.history.heading(
            "currency",
            text="Currency",
            command=lambda: self.sort_history("currency")
        )
        self.history.heading(
            "user",
            text="User",
            command=lambda: self.sort_history("user")
        )
        self.history.heading(
            "details",
            text="Details",
            command=lambda: self.sort_history("details")
        )

        # Add arrow symbol to active column to indicate sort direction (up arrow for ascending, down arrow for descending)
        arrow = " ▼" if self.history_sort_reverse else " ▲"

        # Map column names to their default header text for resetting non-sorted columns to their default text without arrows
        header_text_map = {
            "time": "Time",
            "action": "Action",
            "currency": "Currency",
            "user": "User",
            "details": "Details"
        }

        # Update the heading of the sorted column to include an arrow indicating the sort direction, while resetting other column headings to their default text without arrows
        self.history.heading(
            column_name,
            text=f"{header_text_map[column_name]} {arrow}",
            command=lambda: self.sort_history(column_name)
        )

        # If there are no rows to sort, simply return without attempting to sort or update the Treeview
        if not rows:
            return

        # Map column names to their respective indices in the Treeview values for sorting
        column_index_map = {
            "time": 0,
            "action": 1,
            "currency": 2,
            "user": 3,
            "details": 4
        }
        column_index = column_index_map[column_name]

        # Define sorting key function based on the column type (datetime for "time" column, string for others) and sort the rows accordingly
        if column_name == "time":
            def sort_key(row):
                try:
                    return datetime.strptime(row[column_index], "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    return datetime.min
        else:
            def sort_key(row):
                return str(row[column_index]).upper()

        # Sort the rows based on the defined key function and the current sort direction (ascending or descending)
        rows.sort(key=sort_key, reverse=self.history_sort_reverse)

        # Clear table
        for item_id in self.history.get_children():
            self.history.delete(item_id)

        # Re-insert sorted rows into the Treeview after sorting
        for row in rows:
            self.history.insert("", "end", values=row)

# Main entry point to start the FX Rate Tool application, with error handling to catch any exceptions that occur during startup and display an appropriate error message to the user
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide the root window while initializing
    try:
        app = FXRateApp(root)
        root.deiconify()  # Show the root window after initialization
        root.mainloop()
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"An error occurred while starting the application:\n{str(e)}"
        )
    root.destroy()
