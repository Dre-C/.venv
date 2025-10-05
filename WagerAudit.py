import tkinter as tk #Tkinter library for building the GUI
from tkinter import filedialog, messagebox #file dialog and messagebox for interaction
import pandas as pd #Pandas library for handling Excel fles
import xlsxwriter #XlsxWriter library for writing Excel files with formatting
import math #module to allow for mathematical functions and constants
import re #Regular expression module used for pattern-based string manipulation
import csv #for reading from and writing to CSV
import numpy as np #numerical python
import unicodedata #module for the Unicode Character Database
from pathlib import Path #module for modern object-oriented way to handle filesystem paths


class WagerAuditProgram:
    def __init__(self, master=None):
        self.window = tk.Toplevel(master)
        self.window.title("Wager Audit Comparison Tool") #window title
        self.window.configure(bg="#2b2b2b") #set window background color to white

        self.window.protocol("WM_DELETE_WINDOW", self.close_window) #X button will confirm if user wants to close

        self.wagerAudit_Staging_path = "" #path for Wager Staging Audit File
        self.wagerAudit_Production_path = "" #path for Wager Production Audit File
        self.operator_wagerSheet_path = "" #path for Op Wager Config Sheet File

        self.create_widgets() #function for UI components
        self.adjust_window() #function for screen function

        #Default and min size settings
        self.window.geometry("800x600")
        self.window.minsize(800, 600)

    def close_window(self): #function for cancel confirmation
        confirm = messagebox.askyesno(
            "Exit Wager Audit",
            "Are you sure you want to close the Wager Audit?"
        )
        if confirm:
            self.window.destroy() #To close this window only
        else:
            messagebox.showinfo(
                "Canceled!",
                "Close canceled."
            )

    def adjust_window(self):
        #Get the screen's full width/height
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()

        #Defines the desired window dimensions
        window_width = 800
        window_height = 600

        #Calculate position to center the window
        position_top = (screen_height - window_height) // 2
        position_left = (screen_width - window_width) // 2

        #Update the window's geometry to apply size and position
        self.window.geometry(f'{screen_width}x{window_height}+{position_left}+{position_top}')

    def create_widgets(self):
        #Main content frame for all buttons/labels
        content_frame = tk.Frame(self.window, bg="#2b2b2b", height=300)
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)

        #Welcome display text
        welcome_text = "\nWager\nAudit Comparison Tool\n"
        self.welcome_label = tk.Label(content_frame, text=welcome_text, font=("TkDefaultFont", 15, "bold"), fg='white', bg='#2b2b2b')
        self.welcome_label.pack(pady=10)

        #Group container
        group_container = tk.Frame(content_frame, bg="#2b2b2b")
        group_container.pack()

        #Center group for Staging Wager Audit File/Production Wager Audit File/Op Wager Config Sheet
        center_group = tk.LabelFrame(group_container, text="Wager Audit Files", font=("TkDefaultFont", 8, "bold"), fg='white', bd=3, relief="groove", bg="#2b2b2b", padx=10, pady=10)
        center_group.pack(anchor='center', padx=10)

        #Button style dictionary for all buttons
        button_style = {
            "bg": "#6e6e6e",
            "fg": "white",
            "activebackground": "#505050",
            "activeforeground": "white",
            "borderwidth": 1,
            "highlightthickness": 0,
            "font": ("TkDefaultFont", 10, "bold")
        }
        
        #Label styles dictionary for all labels
        label_style = {
            "bg": "#2b2b2b",
            "fg": "#FF6F6F",
            "padx": 10,
            "pady": 2,
            "font": ("TkDefaultFont", 10)
        }
        #Staging Wager Audit label and upload button
        self.wagerAudit_Staging_label = tk.Label(center_group, text="Select Staging Wager Audit File", **label_style)
        self.wagerAudit_Staging_label.pack(pady=(0, 5))
        self.wagerAudit_Staging_button = tk.Button(center_group, text="Upload Staging Wager Audit File", width=38, command=self.upload_wagerAudit_Staging, **button_style)
        self.wagerAudit_Staging_button.pack(pady=(0, 10))
        self.button_hover_effect(self.wagerAudit_Staging_button)

        #Production Wager Audit label and upload button
        self.wagerAudit_Production_label = tk.Label(center_group, text="Select Wager Production Audit File", **label_style)
        self.wagerAudit_Production_label.pack(pady=(10, 5))
        self.wagerAudit_Production_button = tk.Button(center_group, text="Upload Wager Production Audit File", width=38, command=self.upload_wagerAudit_Production, **button_style)
        self.wagerAudit_Production_button.pack(pady=(0, 10))
        self.button_hover_effect(self.wagerAudit_Production_button)

        #Operator Wager Config Sheet label and upload button
        self.operator_wagerSheet_label = tk.Label(center_group, text="Select Operator Wager Configuration Sheet", **label_style)
        self.operator_wagerSheet_label.pack(pady=(10, 5))
        self.operator_wagerSheet_button = tk.Button(center_group, text="Upload Operator Wager Configuration Sheet", width=38, command=self.upload_operatorWagerSheet, **button_style)
        self.operator_wagerSheet_button.pack(pady=(0, 10))
        self.button_hover_effect(self.operator_wagerSheet_button)

        #Frame for submit/clear buttons
        action_frame = tk.Frame(content_frame, bg="#2b2b2b")
        action_frame.pack(pady=20)

        #Submit button
        self.submit_button = tk.Button(action_frame, text="SUBMIT FILES", font=("TkDefaultFont", 12, "bold"), command=self.submit_files, state=tk.DISABLED, fg='white', bg="#FF6F6F", borderwidth=1)
        self.submit_button.pack(side="left", padx=10)
        self.button_hover_effect(self.submit_button)

        #Clear button
        self.clear_button = tk.Button(action_frame, text="CLEAR FILES", font=("TkDefaultFont", 12, "bold"), command=self.clear_button, fg='white', bg="#6e6e6e", borderwidth=1)
        self.clear_button.pack(side="left", padx=10)
        self.button_hover_effect(self.clear_button)

        #Center action_frame for submit/clear buttons
        action_frame.pack_configure(anchor="center")

    #Adds a hover effect to the buttons
    def button_hover_effect(self, button, hover_bg="#5a5a5a", normal_bg="#6e6e6e"):
        #Takes into consideration the submit buttons color effects of green/red
        def on_enter(e):
            if button.cget("state") == tk.NORMAL:
                if button is self.submit_button and button.cget("bg") == "green":
                    button.config(bg="dark green")
                elif button.cget("bg") != "green":
                    button.config(bg=hover_bg)

        def on_leave(e):
            if button.cget("state") == tk.NORMAL:
                if button is self.submit_button and button.cget("bg") == "dark green":
                    button.config(bg="green")
                elif button.cget("bg") == hover_bg:
                    button.config(bg=normal_bg)
                else:
                    button.config(bg="#FF6F6F")

        #Bind hover effect to the button
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    def enable_submit_button(self):
        #Enables the submit button if files are not empty and turns green. Otherwise, displays red if no files are selected and remains disabled
        if all([self.wagerAudit_Staging_path, self.wagerAudit_Production_path, self.operator_wagerSheet_path]):
            self.submit_button.config(state=tk.NORMAL, bg='green')
        else:
            self.submit_button.config(state=tk.DISABLED, bg='#FF6F6F')
        self.button_hover_effect(self.submit_button)

    def upload_wagerAudit_Staging(self):
        self.wagerAudit_Staging_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("CSV Files", "*.csv")]
            ) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.wagerAudit_Staging_path: #Checks if a file is selected
            self.wagerAudit_Staging_label.config(text=f"Staging Wager Audit File Uploaded: \n{self.wagerAudit_Staging_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Select Staging Wager Audit File to proceed.") #Show warning if no staging wager audit file is selected
            self.wagerAudit_Staging_label.config(text="Select Staging Wager Audit File", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.wagerAudit_Staging_path = "" if not self.wagerAudit_Staging_path else self.wagerAudit_Staging_path
        self.enable_submit_button() #Enables submit button after selection

    def upload_wagerAudit_Production(self):
        self.wagerAudit_Production_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("CSV Files", "*.csv")]
            ) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.wagerAudit_Production_path: #Checks if a file is selected
            self.wagerAudit_Production_label.config(text=f"Wager Production Audit File Uploaded: \n{self.wagerAudit_Production_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Select Production Wager Audit File to proceed.") #Show warning if no production wager audit file is selected
            self.wagerAudit_Production_label.config(text="Select Production Wager Audit File", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.wagerAudit_Production_path = "" if not self.wagerAudit_Production_path else self.wagerAudit_Production_path
        self.enable_submit_button() #Enables submit button after selection 

    def upload_operatorWagerSheet(self):
        self.operator_wagerSheet_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("Excel Files", "*.xlsx")]
            ) #Allows user to upload excel file (this is the file type when file is downloaded)

        if self.operator_wagerSheet_path: #Checks if a file is selected
            self.operator_wagerSheet_label.config(text=f"Operator Wager Configuration Sheet Uploaded: \n{self.operator_wagerSheet_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Select Operator Wager Configuration Sheet to proceed.") #Show warning if no op wager config sheet is selected
            self.operator_wagerSheet_label.config(text="Select Operator Wager Configuration Sheet", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.operator_wagerSheet_path = "" if not self.operator_wagerSheet_path else self.operator_wagerSheet_path
        self.enable_submit_button() #Enables submit button after selection

    def submit_files(self):
        #Checks if all files are uploaded
        if not all([self.wagerAudit_Staging_path, self.wagerAudit_Production_path, self.operator_wagerSheet_path]):
            messagebox.showwarning("Incomplete files!", "Upload all required files before submitting.") #Show warning if not all files were uploaded
            return

        #Allows user to select the file save location
        file_path = filedialog.asksaveasfilename(
            parent=self.window,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")], #file types filter
            title="File Save Location" #dialog title
        )

        if not file_path:
            messagebox.showinfo("Missing File Path!",
                                "Select file path to save Wager Audit Results and try again.") #Show cancelled message if no save file path was selected
            self.enable_submit_button() #Enables submit button
            return

        #Message box to confirm user selected files for submission and allows user to hit cancel if needed to re-upload files
        if messagebox.askyesno("Confirm Submit",
                               "Are you sure you want to submit files for comparison?"):
            try:
                result = self.compare_files(file_path) #Call the function to compare files and save
                if result:
                    messagebox.showinfo("Wager Audit Results Saved!",
                                        f"Wager Audit Results successfully saved at: {file_path}.") #Success message and show user save location
                else:
                    messagebox.showerror("Error!",
                                         "Failed to save file. Check the correct file formats were submitted and try again.") #Show failure message if results fail
            except Exception as e:
                messagebox.showerror("Error!",
                                     f"Error occurred during export: {str(e)}") #Show error if there's an exception while saving files
        else:
            messagebox.showinfo("Canceled!",
                                "File submission canceled. Upload all required files to submit and try again.") #Display cancel message if user hits cancel

        self.enable_submit_button() #Resets submit button to it's default state after handling success, cancellation, or missing file path

    def clear_button(self):
        answer = messagebox.askyesno(
            "Confirm Clear",
            "Are you sure you want to clear all files selected?"
        )
        if answer:
            #Clear all file paths if yes is selected
            self.wagerAudit_Staging_path = ""
            self.wagerAudit_Production_path = ""
            self.operator_wagerSheet_path = ""
            
            #Clear all labels and display red text
            self.wagerAudit_Staging_label.config(text="Select Staging Wager Audit File", fg="#FF6F6F")
            self.wagerAudit_Production_label.config(text="Select Production Wager Audit File", fg="#FF6F6F")
            self.operator_wagerSheet_label.config(text="Select Operator Wager Configuration Sheet", fg="#FF6F6F")

            #Disable the submit button and turn red
            self.submit_button.config(state=tk.DISABLED, bg="#FF6F6F")

            #Show message box to user stating cleared files
            messagebox.showinfo("All Files Cleared!",
                                "All uploaded files were cleared. Select new files to upload.")
            
        else: #Show message box to user the clear was canceled
            messagebox.showinfo("Canceled!",
                                "Clear canceled.")

    def normalize_name(self, name):
        #Standardize game name column; convert to lowercase, removes all spaces, removes apostrophes
        if isinstance(name, str):
            name = unicodedata.normalize('NFKD', name) #Normalize any smart quotes or accents (Ex: Jack O'Lantern Jackpots)
            name = re.sub(r'(?<!^)(?=[A-Z][a-z])', ' ', name) #Split only before capital letters followed by lowercase (to avoid splitting acronyms)
            name = name.replace('_', ' ') #Remove underscores and adds a space (specific for postfix games)
            name = re.sub(r"[’';:]", '', name) #Remove straight and curly apostrophes using regex
            name = re.sub(r'\s+', '', name).strip() #Replace multiple spaces with no space, then strip leading/trailing
            return name.lower() #Convert to lowercase
        return name

    def normalize_value(self, val, is_percent_column=False):
        #Standardize values to handle percentages, currencies, and NaN values
        if pd.isna(val) or str(val).strip == '': #Return empty string for NaN, empty string, or whitespace
            return '' if not is_percent_column else pd.NA

        val_str = str(val).strip()

        #Handles converting percentages first
        if is_percent_column:
            try:
                if '%' in val_str:
                    numeric_val = float(val_str.replace('%', '').replace(',', '').strip())
                else:
                    numeric_val = float(val_str.replace(',', '').strip())
                if numeric_val > 1:
                    numeric_val /= 100
                return numeric_val
            except (ValueError, TypeError):
                raise ValueError(f"Invalid percent value: {val_str}")

        #Handles multiple values separated by commas or space separated values (ex: $0.01, $0.05, $0.10, etc.)
        if any(sym in val_str for sym in ('$', '€', '£')): #Can add more currencies as needed
            currency_values = re.findall(r'[\$€£]?\d[\d,]*\.?\d*', val_str) #Regex to detect multiple values vs single values

            if len(currency_values) > 1: #Checks for more than one currency value
                parts = [v.strip() for v in val_str.split(',')] #Split by commas and strip whitespace from each individual value
                normalized_values = [self.normalize_currency_values(p) for p in parts] #Normalize each stripped currency value using def normalize_currency_values methood
                return ','.join(normalized_values) #Join normalized values back into a single comma-separated string
            else:
                return self.normalize_currency_values(val_str) #Single value - normalize it directly

        val_str = val_str.replace(' ', '') #Remove all spaces

        #Handle comma-separated when currency is not included (ex: .25,1.00,5.00 -> 0.25,1,5)
        if ',' in val_str:
            parts = val_str.split(',')
            normalized_parts = [self.clean_number_string(p) for p in parts]
            val_str = ','.join(normalized_parts)
        else:
            val_str = self.clean_number_string(val_str)

        #Capitalize any letters in the final normalized value (ex: 243Ways -> 243WAYS)
        val_str = ''.join(c.upper() if c.isalpha() else c for c in val_str)

        return val_str

    def clean_number_string(self, val):
        #Handles values without currency symbols such as default lines & bet multipliers (or data entries for denoms such as .25, 1.00, 5.00 -> 0.25,1,5)
        try:
            val_clean = str(val).strip()
            if val_clean.startswith("."):
                val_clean = "0" + val_clean
            num = float(val_clean)
            if num.is_integer():
                return str(int(num))
            else:
                return str(num).rstrip("0").rstrip(".")
        except (ValueError, TypeError):
            return str(val).strip() #If conversion fails (val is not a number), return original value as a stripped string

    def normalize_currency_values(self, val):
    #Helper method to handle currency symbols (ex: $€£) and commas; can expand currencies as needed
        try:
            val_clean = re.sub(r'[$€£,]', '', str(val)).strip() #Remove the currency symbols/commas using regex
            if val_clean.startswith("."):
                val_clean = "0" + val_clean
            num = float(val_clean) #Convert to float
            if num.is_integer(): #Checks if float is an integer
                return str(int(num)) #If integer, convert to an integer then to a string (removes the decimal point)
            else:
                return str(num).rstrip("0").rstrip(".") #If not an integer, return as string formatted with two decimal places
        except (ValueError, TypeError):
            return '' #if conversion fails, return empty string

    def detect_header_row(self, file_path, header_indicator="Game"):
    #Handles automatically detecting header rows by scanning all rows
        if file_path.endswith('.xlsx'): #Read Excel file
            wager_data = pd.read_excel(file_path, header=None, engine='openpyxl') #Checks all rows for header
            wager_data = wager_data.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x)) #Cleans up unwanted spaces before further processing

            #DEBUG: Print first 5 rows for inspection
            print("\nDEBUG Excel Files: Preview of first 5 raw rows:")
            print(wager_data.head())

        elif file_path.endswith('.csv'): #Handles csv files differently
            rows = [] #Empty list to store rows
            with open(file_path, 'r', encoding='ISO-8859-1') as f: #DEBUG to print first 5 lines
                reader = csv.reader(f)
                print("\nDEBUG CSV Files: Preview of first 5 raw rows:")

                for i, row in enumerate(reader): #Iterate over each row
                    standardized_row = [cell.strip() if isinstance(cell, str) and cell.strip() else '' for cell in row]
                    
                    if i < 5: #DEBUG: print standardized row for first 5 rows
                        print(f"Line {i}: {standardized_row}")
                    rows.append(standardized_row) #Append normalized row to the list of rows

            #Convert rows to DataFrame after reading rows, replace empty strings, None values with NaN for easier handling
            wager_data = pd.DataFrame(rows).replace(['', None], np.nan)
        else:
            raise ValueError("Unsupported file format. Only ('.csv') and ('.xlsx') file types are supported.") #Raise error for incorrect file formats
                        
        for idx, row in wager_data.iterrows(): #Iterate through each row, convert all values to string, strip spaces
            row_values = [str(cell).strip() for cell in row.values if isinstance(cell, str)]
            print(f"Checking row {idx}: {row_values}") #DEBUG to print specific header rows it's detecting

            #Check if 'Game' is a part of any column names in this row
            if any(header_indicator in value for value in row_values):
                print(f"Header row detected at index {idx}")

                new_header = wager_data.iloc[idx] #Grab header row use it as new column names
                wager_data = wager_data[(idx + 1):].copy() #Drop all rows above header, keep data rows below header
                wager_data.columns = new_header #Assign new header row to the DataFrame columns
                wager_data.columns = wager_data.columns.astype(str).str.replace('\n', ' ', regex=False).str.strip() #Clean column names
                wager_data = wager_data.loc[:, ~wager_data.columns.duplicated()] #Remove duplicate column names

                return idx
            
        raise ValueError("No matching header row found. Check files to ensure proper files were uploaded and try again.") #Raise error for when headers are not found

    def compare_files(self, file_path):
            #Checks if all required files are missing
            if not all([self.wagerAudit_Staging_path, self.wagerAudit_Production_path, self.operator_wagerSheet_path]):
                messagebox.showerror("Error!", "Upload all required files to proceed.") #Show error if any files are missing
                return False #Stop further execution if files are incomplete
            
            all_valid = True #Set the validation flag to True if all files are present and proceed with processing
            
            #Process Wager Staging/Production Audit Files and Operator Wager Config Sheet
            try:
                #Checks required columns are present in both files
                wagerAudit_columns = ["Everi Game ID", "RTP MAX", "Denom", "Line Selection", "Bet Multiplier Selection", "Default Denom", "Default Line", 
                                      "Default Bet Multiplier", "Default Bet", "Min Bet", "Max Bet"]
            
                operatorSheet_columns = ["Game", "RTP%", "Denom Selection", "Line/Ways Selection", "Bet Multiplier Selection", "Default Denom Selection", "Default Line/Ways", 
                                        "Default Bet Multiplier", "Total Default Bet", "Min Bet", "Max Bet"]

                #Defining column mapping for wager audit manually so that names match data
                column_mapping_wager = {
                    "Everi Game ID": "Game",
                    "RTP MAX": "RTP%",
                    "Denom": "Denom Selection",
                    "Line Selection": "Line/Ways Selection",
                    "Bet Multiplier Selection": "Bet Multiplier Selection",
                    "Default Denom": "Default Denom Selection",
                    "Default Line": "Default Line/Ways",
                    "Default Bet Multiplier": "Default Bet Multiplier",
                    "Default Bet": "Total Default Bet",
                    "Min Bet": "Min Bet",
                    "Max Bet": "Max Bet"
                }

                #Detect the header rows for files automatically finding column names
                wagerAudit_Staging_header_row = self.detect_header_row(self.wagerAudit_Staging_path, header_indicator="Everi Game ID")
                wagerAudit_Production_header_row = self.detect_header_row(self.wagerAudit_Production_path, header_indicator="Everi Game ID")
                operatorSheet_header_row = self.detect_header_row(self.operator_wagerSheet_path, header_indicator="Game")

                #Throws an error if no valid header rows are found in the files
                if wagerAudit_Staging_header_row is None or wagerAudit_Production_header_row is None or operatorSheet_header_row is None:
                    messagebox.showerror("Error!", "Could not find valid header rows in all selected files.")
                    return False
            
                #Read full files, skipping the detected header rows
                if self.wagerAudit_Staging_path.endswith('.csv'):
                    wagerAudit_StagingFile = pd.read_csv(self.wagerAudit_Staging_path, skiprows=wagerAudit_Staging_header_row, encoding='ISO-8859-1') #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Staging Wager File. Only ('.csv') file type is supported.") #Raise error if incorrect file type is selected
                
                #Read full files, skipping the detected header rows
                if self.wagerAudit_Production_path.endswith('.csv'):
                    wagerAudit_ProductionFile = pd.read_csv(self.wagerAudit_Production_path, skiprows=wagerAudit_Production_header_row, encoding='ISO-8859-1') #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Production Wager Audit File. Only ('.csv') file type is supported.") #Raise error if incorrect file type is selected

                if self.operator_wagerSheet_path.endswith('.xlsx'):
                    operatorSheet_file = pd.read_excel(self.operator_wagerSheet_path, header=operatorSheet_header_row, engine='openpyxl') #File format is downloaded as xlsx therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Operator Wager Configuration Sheet. Only ('.xlsx') file type is supported.") #Raise error if incorrect file type is selected
                               
                #Normalize column names, strip spaces
                wagerAudit_StagingFile.columns = wagerAudit_StagingFile.columns.astype(str).str.strip()
                wagerAudit_ProductionFile.columns = wagerAudit_ProductionFile.columns.astype(str).str.strip()
                operatorSheet_file.columns = operatorSheet_file.columns.astype(str).str.strip()

                #Filter only relevant columns
                wagerAudit_StagingFile = wagerAudit_StagingFile[wagerAudit_columns]
                wagerAudit_ProductionFile = wagerAudit_ProductionFile[wagerAudit_columns]
                operatorSheet_file = operatorSheet_file[operatorSheet_columns]

                #Identify if expected columns are missing
                missing_wagerAudit_Staging_columns = [col for col in wagerAudit_columns if col not in wagerAudit_StagingFile.columns]
                missing_wagerAudit_Production_columns = [col for col in wagerAudit_columns if col not in wagerAudit_ProductionFile.columns]
                missing_operatorSheet_columns = [col for col in operatorSheet_columns if col not in operatorSheet_file.columns]

                #Checks for missing columns
                if missing_wagerAudit_Staging_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Staging Wager Audit File: {', '.join(missing_wagerAudit_Staging_columns)}")
                    return False
                if missing_wagerAudit_Production_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Production Wager Audit File: {', '.join(missing_wagerAudit_Production_columns)}")
                    return False
                if missing_operatorSheet_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Operator Wager Configuration Sheet: {', '.join(missing_operatorSheet_columns)}")
                    return False
                              
                #Renames columns to match column mapping
                try:
                    wagerAudit_StagingFile = wagerAudit_StagingFile.rename(columns=column_mapping_wager)
                    wagerAudit_ProductionFile = wagerAudit_ProductionFile.rename(columns=column_mapping_wager)
                    operatorSheet_file = operatorSheet_file.rename(columns=column_mapping_wager)
                except Exception as e:
                    messagebox.showerror("Error in column_mapping_wager", str(e))
                    return False
                
                #Handles all missing columns by adding them with NaN values to all DataFrames
                for col in column_mapping_wager.values():
                    if col not in wagerAudit_StagingFile.columns:
                        wagerAudit_StagingFile[col] = pd.NA
                    if col not in wagerAudit_ProductionFile.columns:
                        wagerAudit_ProductionFile[col] = pd.NA
                    if col not in operatorSheet_file.columns:
                        operatorSheet_file[col] = pd.NA

                #Applies normalization to Game Name column only
                wagerAudit_StagingFile['Game'] = wagerAudit_StagingFile['Game'].apply(self.normalize_name)
                wagerAudit_ProductionFile['Game'] = wagerAudit_ProductionFile['Game'].apply(self.normalize_name)
                operatorSheet_file['Game'] = operatorSheet_file['Game'].apply(self.normalize_name)

                #Handle RTP% column for operatorSheet_file specifically
                percent_column = ['RTP%']

                #Skip Game column and normalize values in the other columns
                for f in [wagerAudit_StagingFile, wagerAudit_ProductionFile, operatorSheet_file]:
                    for wager_column in f.columns:
                        if wager_column != 'Game':
                            continue
                        f[wager_column] = f[wager_column].apply(lambda x: self.normalize_value(x, is_percent_column=(wager_column in percent_column)))

                #Sorts Game columns alphabetically in all DataFrames
                wagerAudit_StagingFile = wagerAudit_StagingFile.sort_values(by='Game', ascending=True)
                wagerAudit_ProductionFile = wagerAudit_ProductionFile.sort_values(by='Game', ascending=True)
                operatorSheet_file = operatorSheet_file.sort_values(by='Game', ascending=True)

                #Removes duplicates in DataFrames to ensure it only appears once
                wagerAudit_StagingFile = wagerAudit_StagingFile.drop_duplicates(subset='Game')
                wagerAudit_ProductionFile = wagerAudit_ProductionFile.drop_duplicates(subset='Game')
                operatorSheet_file = operatorSheet_file.drop_duplicates(subset='Game')

                #File labels for labeling on Missing Games sheet
                file_labels = ['Staging Wager Audit File',
                               'Production Wager Audit File',
                               'Operator Wager Configuration Sheet']

                #Get sets of Game Names from each file
                games_wagerAudit_StagingFile = set(wagerAudit_StagingFile['Game'])
                games_wagerAudit_ProductionFile = set(wagerAudit_ProductionFile['Game'])
                games_operatorSheet_file = set(operatorSheet_file['Game'])

                #Ensures DataFrames have only matching Game values
                common_games_wager = games_wagerAudit_StagingFile & games_wagerAudit_ProductionFile & games_operatorSheet_file

                #Union of all Game Names across all three files
                all_games = sorted(games_wagerAudit_StagingFile | games_wagerAudit_ProductionFile | games_operatorSheet_file)
                allFiles_set = [games_wagerAudit_StagingFile, games_wagerAudit_ProductionFile, games_operatorSheet_file]

                allMissing_games = [] #Empty list to collect missing Game Names

                for game in all_games:
                    missing_in = [
                        file_labels[i] for i, game_sets in enumerate(allFiles_set)
                        if game not in game_sets
                    ]
                    #Append one row per Game Name with combined missing info
                    if missing_in:
                        combined_status = ', '.join(missing_in)
                        allMissing_games.append({
                            'Game': game,
                            'Status': f'Missing in {combined_status}'
                        })

                #Convert missing Game Names list of dicts into a DataFrame and sort it for Missing Games Sheet
                missing_games_wager = pd.DataFrame(allMissing_games).sort_values(by='Game').reset_index(drop=True)

                #Filter rows based on common Game Names in all three files and sort by game
                wagerAudit_StagingFile_matchedGameNames = wagerAudit_StagingFile[wagerAudit_StagingFile['Game'].isin(common_games_wager)].sort_values(by='Game', ascending=True).reset_index(drop=True)
                wagerAudit_ProductionFile_matchedGameNames = wagerAudit_ProductionFile[wagerAudit_ProductionFile['Game'].isin(common_games_wager)].sort_values(by='Game', ascending=True).reset_index(drop=True)
                operatorSheet_file_matchedGameNames = operatorSheet_file[operatorSheet_file['Game'].isin(common_games_wager)].sort_values(by='Game', ascending=True).reset_index(drop=True)

                #DataFrame for Wager Audit Results to hold side-by-side columns for comparison
                audit_results_wagers = pd.DataFrame({'Game': wagerAudit_StagingFile_matchedGameNames['Game'].values})

                #Single loop to handle renamed columns to normalize values and add columns side by side
                for wager_column in wagerAudit_StagingFile_matchedGameNames.columns:
                    if wager_column == 'Game':
                        continue
                    if wager_column not in wagerAudit_StagingFile_matchedGameNames.columns:
                        raise KeyError(f"'{wager_column}' not found in 'wagerAudit_StagingFile_matchedGameNames' matched rows dataset.")
                    if wager_column not in wagerAudit_ProductionFile_matchedGameNames.columns:
                        raise KeyError(f"'{wager_column}' not found in 'wagerAudit_ProductionFile_matchedGameNames' matched rows dataset.")
                    if wager_column not in operatorSheet_file_matchedGameNames.columns:
                        raise KeyError(f"'{wager_column}' not found in 'operatorSheet_file_matchedGameNames' matched rows dataset.")
                    
                    if wager_column in wagerAudit_StagingFile_matchedGameNames.columns and wager_column in wagerAudit_ProductionFile_matchedGameNames.columns and wager_column in operatorSheet_file_matchedGameNames.columns:
                        wagerAudit_StagingFile_matchedGameNames[wager_column] = wagerAudit_StagingFile_matchedGameNames[wager_column].reset_index(drop=True)
                        wagerAudit_ProductionFile_matchedGameNames[wager_column] = wagerAudit_ProductionFile_matchedGameNames[wager_column].reset_index(drop=True)
                        operatorSheet_file_matchedGameNames[wager_column] = operatorSheet_file_matchedGameNames[wager_column].reset_index(drop=True)

                        #Side by side columns from all sheets to the DataFrame
                        audit_results_wagers[f"{wager_column}\n(Wager Staging Audit File): "] = wagerAudit_StagingFile_matchedGameNames[wager_column]
                        audit_results_wagers[f"{wager_column}\n(Wager Production Audit File): "] = wagerAudit_ProductionFile_matchedGameNames[wager_column]
                        audit_results_wagers[f"{wager_column}\n({Path(self.operator_wagerSheet_path).stem[:31]}): "] = operatorSheet_file_matchedGameNames[wager_column]

                audit_results_wagers['Game'] = wagerAudit_StagingFile_matchedGameNames['Game'].values
                cols = list(audit_results_wagers.columns)
                cols.remove('Game')
                cols.insert(0, 'Game')
                audit_results_wagers = audit_results_wagers[cols]

                audit_results_wagers = audit_results_wagers.sort_values(by='Game', ascending=True).reset_index(drop=True)

                for f in [wagerAudit_StagingFile, wagerAudit_ProductionFile, operatorSheet_file, audit_results_wagers]:
                    for wager_column in percent_column:
                        if wager_column in f.columns:
                            f[wager_column] = f[wager_column].apply(lambda x: f"{int(round(x * 100))}%" if isinstance(x, (int, float)) and not pd.isna(x) else 'N/A')

            except Exception as e:
                all_valid = False
                print(f"Error caught in except block: {e}")
                messagebox.showerror("Error", f"An error has occured for Staging Wager Audit File, Production Wager Audit File, and Operator Wager Configuration Sheet: {str(e)}")
                return False
            
            #If all files are processed successfully and True, proceed with Excel writing
            if all_valid:
                #Safety check to ensure file names are not the same so that it does not overwrite sheets accidently
                sheet_names_wagerAuditGroup = [
                    Path(self.wagerAudit_Staging_path).stem[:31],
                    Path(self.wagerAudit_Production_path).stem[:31],
                    Path(self.operator_wagerSheet_path).stem[:31]
                ]
                #Check for duplicates in wagerAuditGroup
                if len(sheet_names_wagerAuditGroup) != len(set(sheet_names_wagerAuditGroup)):
                    messagebox.showerror(
                        "Error Duplicate File Names Detected!",
                        f'Duplicate file names detected for files: {sheet_names_wagerAuditGroup}.\n'
                        'Rename files to ensure unique sheet names and re-upload again.'
                    )
                    return #Stop execution until files are renamed properly

                try:
                    #Write to excel with formatting if file names are unique
                    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                        #Write to Excel with sheet names (based on selected file paths) truncated to 31 characters
                        wagerAudit_StagingFile.to_excel(writer, sheet_name=sheet_names_wagerAuditGroup[0], index=False) #Staging Wager Audit File raw data on sheet 1
                        wagerAudit_ProductionFile.to_excel(writer, sheet_name=sheet_names_wagerAuditGroup[1], index=False) #Production Wager Audit File raw data on sheet 2
                        operatorSheet_file.to_excel(writer, sheet_name=sheet_names_wagerAuditGroup[2], index=False) #Op Wager Config Sheet raw data on sheet 3
                        audit_results_wagers.to_excel(writer, sheet_name='Wager Audit Results', index=False) #Wager Audit Results with side by side comparison on sheet 4
                        missing_games_wager.to_excel(writer, sheet_name='Missing Games', index=False) #Missing games from all files on sheet 5

                        #Access the workbook and worksheet to apply formatting
                        workbook = writer.book

                        #Define formats
                        header_format = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'border': 2, 'text_wrap': True}) #Grey header format (bold, thick borders)
                        cell_format = workbook.add_format({'border': 1, 'border_color': '#BFBFBF'}) #Borders for data cells
                        red_format = workbook.add_format({'bg_color': '#FF6F6F'}) #Red format highlights cells red when there's a mismatch on the Wager Audit Comparison Results

                        #Loop & apply formats to all sheets
                        for df, sheet_name in [
                            (wagerAudit_StagingFile, Path(self.wagerAudit_Staging_path).stem[:31]),
                            (wagerAudit_ProductionFile, Path(self.wagerAudit_Production_path).stem[:31]),
                            (operatorSheet_file, Path(self.operator_wagerSheet_path).stem[:31]),
                            (audit_results_wagers, 'Wager Audit Results'),
                            (missing_games_wager, 'Missing Games')
                        ]:
                            worksheet = writer.sheets[sheet_name]

                            #Header row formatting and auto-adjust column widths
                            for col_num, column_name in enumerate(df.columns):
                                worksheet.write(0, col_num, column_name, header_format)

                                #Auto-adjust column width to fit contents by calculating optimal column widths based on header/data length
                                if df[column_name].notna().any():
                                    max_val_len = df[column_name].astype(str).map(len).max()
                                else:
                                    max_val_len = 0

                                max_len = max(max_val_len, len(column_name))
                                worksheet.set_column(col_num, col_num, max_len + 2) #Add padding

                            worksheet.autofilter(0, 0, 0, len(df.columns) - 1) #Add filter to header row
                            worksheet.freeze_panes(1, 0) #Freeze top row to keep headers visible when scrolling

                            #Write all data cells w/border formatting
                            for row in range(1, len(df) + 1):
                                for col in range(len(df.columns)):
                                    val = df.iat[row - 1, col]
                                    if pd.isna(val) or val in [float('inf'), float('-inf')]:
                                        worksheet.write(row, col, "", cell_format)
                                    else:
                                        worksheet.write(row, col, val, cell_format)

                            #Iterates through rows/columns to apply formatting for mismatches
                            for row in range(1, len(df) + 1):
                                col_idx = 0 #Start at the first column
                                while col_idx < len(df.columns):
                                    try:
                                        remaining_columns = len(df.columns) - col_idx #Calculate remaining columns
                                        column_name = df.columns[col_idx]

                                        #Detect single columns for combined columns dynamically
                                        single_column = column_name == 'Game' or remaining_columns < 3
                                        #Excluding column name 'Game' column from being highlighted since it is combined
                                        if single_column:
                                            val = df.iat[row - 1, col_idx]
                                            worksheet.write(row, col_idx, val)
                                            col_idx += 1
                                            continue

                                        #Access up to 3 values for comparison
                                        val1 = df.iat[row - 1, col_idx]
                                        val2 = df.iat[row - 1, col_idx + 1] if remaining_columns > 1 else None
                                        val3 = df.iat[row - 1, col_idx + 2] if remaining_columns > 2 else None

                                        val1 = val1 if val1 is not None else ""
                                        val2 = val2 if val2 is not None else ""
                                        val3 = val3 if val3 is not None else ""

                                        #Replace NaN or None with empty string (if necessary)
                                        columns_in_groups = ["" if pd.isna(val1) or val1 is None else val1]
                                        if val2 is not None:
                                            columns_in_groups.append("" if pd.isna(val2) or val2 is None else val2)
                                        if val3 is not None:
                                            columns_in_groups.append("" if pd.isna(val3) or val3 is None else val3)

                                        #Only apply red highlighting to audit_results_wagers
                                        if df is audit_results_wagers:
                                            n_vals = len(columns_in_groups)
                                            highlight_flags = [False] * n_vals #Flags for highlighting

                                            #Count occurrences of each value
                                            value_counts = {}
                                            for v in columns_in_groups:
                                                value_counts[v] = value_counts.get(v, 0) + 1

                                            #Highlight all values if all do not match
                                            if len(value_counts) ==  n_vals:
                                                highlight_flags = [True] * n_vals
                                            else:
                                                max_frequency = max(value_counts.values()) #Find the majority value
                                                #Flag any value that is not in majority
                                                for i, v in enumerate(columns_in_groups):
                                                    if value_counts[v] < max_frequency:
                                                        highlight_flags[i] = True

                                            #Write values to worksheet with red formatting
                                            for i, v in enumerate(columns_in_groups):
                                                fmt = red_format if highlight_flags[i] else None
                                                worksheet.write(row, col_idx + i, v, fmt)

                                            col_idx += n_vals #Move past this group

                                        else: #For all other sheets write normally without highlighting/comparison
                                            write_columns = min(3, remaining_columns)
                                            for i in range(write_columns):
                                                val = df.iat[row - 1, col_idx + i]
                                                worksheet.write(row, col_idx + i, val)
                                            col_idx += write_columns

                                    except Exception as e: #Debug to catch errors per row/column
                                        print(f"Error processing row: '{row}', col_idx: '{col_idx}': '{e}'")
                                        col_idx += 1 #Move to next column to avoid an infinite loop

                except Exception as e:
                    all_valid = False
                    messagebox.showerror("Error writing to Excel", str(e))
                    return False

            #Success message when results are True and all passes successfully
            if all_valid:
                messagebox.showinfo("Success!", "All files processed successfully and Wager Audit Results are complete.")
                return True
            else:
                return False
            
