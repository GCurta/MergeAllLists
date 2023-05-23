import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
import sys
import subprocess
import os

class ExcelMerger(tk.Tk):
    def __init__(self):
        super().__init__()

        # Check and install necessary packages
        self._check_and_install("pandas")
        self._check_and_install("openpyxl")

        self.title("Excel Merger")
        self.state('zoomed')  # Start with maximized window
        self.configure(bg='white')  # Set a white background for the whole window

        # Variables
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_file = tk.StringVar()
        self.open_output_file = tk.BooleanVar(value=True)  # Option to open output file after merge

        # Color scheme
        self.primary_color = "#3366CC"  # Blue
        self.secondary_color = "#FFFFFF"  # White
        self.button_text_color = "#FFFFFF"  # White
        self.merge_button_color = "#00CC66"  # Green
        self.how_to_button_color = "#BBBBBB"  # Light Gray
        self.exit_button_color = "#BBBBBB"  # Light Gray

        # Font
        self.font_family = "Source Sans Pro"

        # UI setup
        tk.Label(self, text="SimplifyIT", bg='white', font=(self.font_family, 30, "bold"), fg=self.primary_color).pack(side='top', anchor='ne', padx=10, pady=10)

        # File 1 area
        file1_frame = tk.Frame(self, bg='white')
        file1_frame.pack(side='left', anchor='nw', padx=10, pady=10)

        tk.Label(file1_frame, text="File 1", bg='white', font=(self.font_family, 18)).pack(side='top', anchor='w')
        tk.Button(file1_frame, text="Choose File", command=lambda: self._browse_file(self.file1_path, self.file1_listbox),
                   font=(self.font_family, 15), relief='raised', bd=3, padx=10, pady=5, highlightthickness=0, bg=self.primary_color, fg=self.button_text_color).pack(side='top', anchor='w')
        self.file1_listbox = tk.Listbox(file1_frame, selectmode='multiple', exportselection=0, width=40, height=10, font=(self.font_family, 15))
        self.file1_listbox.pack(side='top', anchor='w')
        tk.Label(file1_frame, text="File 1 Path:", bg='white', font=(self.font_family, 12)).pack(side='top', anchor='w')
        tk.Label(file1_frame, textvariable=self.file1_path, bg='white', font=(self.font_family, 12)).pack(side='top', anchor='w')

        # File 2 area
        file2_frame = tk.Frame(self, bg='white')
        file2_frame.pack(side='right', anchor='ne', padx=10, pady=10)

        tk.Label(file2_frame, text="File 2", bg='white', font=(self.font_family, 18)).pack(side='top', anchor='w')
        tk.Button(file2_frame, text="Choose File", command=lambda: self._browse_file(self.file2_path, self.file2_listbox),
                   font=(self.font_family, 15), relief='raised', bd=3, padx=10, pady=5, highlightthickness=0, bg=self.primary_color, fg=self.button_text_color).pack(side='top', anchor='w')
        self.file2_listbox = tk.Listbox(file2_frame, selectmode='multiple', exportselection=0, width=40, height=10, font=(self.font_family, 15))
        self.file2_listbox.pack(side='top', anchor='w')
        tk.Label(file2_frame, text="File 2 Path:", bg='white', font=(self.font_family, 12)).pack(side='top', anchor='w')
        tk.Label(file2_frame, textvariable=self.file2_path, bg='white', font=(self.font_family, 12)).pack(side='top', anchor='w')

        # Output file area
        output_frame = tk.Frame(self, bg='white')
        output_frame.pack(side='top', anchor='w', padx=10, pady=10)

        tk.Label(output_frame, text="Output File", bg='white', font=(self.font_family, 18)).pack(side='top', anchor='w')
        tk.Button(output_frame, text="Choose File", command=self._browse_output_file,
                   font=(self.font_family, 15), relief='raised', bd=3, padx=10, pady=5, highlightthickness=0, bg=self.primary_color, fg=self.button_text_color).pack(side='top', anchor='w')
        tk.Label(output_frame, text="Output File Path:", bg='white', font=(self.font_family, 12)).pack(side='top', anchor='w')
        tk.Label(output_frame, textvariable=self.output_file, bg='white', font=(self.font_family, 12)).pack(side='top', anchor='w')

        # Open output file option
        open_output_frame = tk.Frame(self, bg='white')
        open_output_frame.pack(side='top', anchor='w', padx=10, pady=10)

        self.open_output_file_checkbox = tk.Checkbutton(open_output_frame, text="Open Output File After Merge",
                                                        variable=self.open_output_file, bg='white',
                                                        font=(self.font_family, 12))
        self.open_output_file_checkbox.pack(side='top', anchor='w')

        # Buttons
        buttons_frame = tk.Frame(self, bg='white')
        buttons_frame.pack(side='top', anchor='e', padx=10, pady=10)

        tk.Button(buttons_frame, text="Merge", command=self._merge_files, font=(self.font_family, 30, "bold"), fg=self.button_text_color, bg=self.merge_button_color, relief='raised', bd=3, padx=10, pady=5, highlightthickness=0).pack(side='right', padx=10)
        tk.Button(buttons_frame, text="Exit", command=self.destroy, font=(self.font_family, 18), fg=self.button_text_color, bg=self.exit_button_color, relief='raised', bd=3, padx=10, pady=5, highlightthickness=0).pack(side='right', padx=10)
        tk.Button(buttons_frame, text="HowTo", command=self._open_how_to_window, font=(self.font_family, 18), fg=self.button_text_color, bg=self.how_to_button_color, relief='raised', bd=3, padx=10, pady=5, highlightthickness=0).pack(side='left', padx=10)

        # Adding a text area for the logs
        self.log_area = scrolledtext.ScrolledText(self, width=100, height=10, font=(self.font_family, 18), bg=self.secondary_color, fg="black")
        self.log_area.pack(side='bottom', fill='both', padx=10, pady=10)
        tk.Label(self, text="Log", bg='white', font=(self.font_family, 18)).pack(side='bottom', anchor='w', padx=10)

        # Configure resizing behavior
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(2, weight=1)
        self.rowconfigure(3, weight=1)

        # Initialize the first script elements
        self.input_entry = None
        self.output_entry = None
        self.delimiter_combobox = None
        self.custom_delimiter_entry = None

        # Call the function to add the first script elements to the GUI
        self._add_csv_to_excel_converter()

    def _add_csv_to_excel_converter(self):
        # Create the frame for the CSV to Excel Converter
        csv_converter_frame = tk.LabelFrame(self, text="CSV to Excel Converter", bg='white', font=(self.font_family, 18))
        csv_converter_frame.pack(side='top', anchor='w', padx=10, pady=10)

        # Input CSV File
        input_label = tk.Label(csv_converter_frame, text="Input CSV File:", bg='white', font=(self.font_family, 12))
        input_label.grid(row=0, column=0)
        self.input_entry = tk.Entry(csv_converter_frame, width=30)
        self.input_entry.grid(row=0, column=1)
        input_button = tk.Button(csv_converter_frame, text="Browse...", command=self.select_input_file, font=(self.font_family, 12))
        input_button.grid(row=0, column=2)

        # Output Excel File
        output_label = tk.Label(csv_converter_frame, text="Output Excel File:", bg='white', font=(self.font_family, 12))
        output_label.grid(row=1, column=0)
        self.output_entry = tk.Entry(csv_converter_frame, width=30)
        self.output_entry.grid(row=1, column=1)
        output_button = tk.Button(csv_converter_frame, text="Browse...", command=self.select_output_file, font=(self.font_family, 12))
        output_button.grid(row=1, column=2)

        # CSV Delimiter
        delimiter_label = tk.Label(csv_converter_frame, text="CSV Delimiter:", bg='white', font=(self.font_family, 12))
        delimiter_label.grid(row=2, column=0)

        self.delimiter_combobox = ttk.Combobox(csv_converter_frame, values=[",", ";", "\t", "|", "Custom Delimiter"])
        self.delimiter_combobox.grid(row=2, column=1)
        self.delimiter_combobox.current(0)  # Set default selection
        self.delimiter_combobox.bind("<<ComboboxSelected>>", lambda event: self.handle_delimiter_selection(event.widget.get()))

        custom_delimiter_label = tk.Label(csv_converter_frame, text="Custom Delimiter:", bg='white', font=(self.font_family, 12))
        custom_delimiter_label.grid(row=3, column=0, sticky="e")
        self.custom_delimiter_entry = tk.Entry(csv_converter_frame, width=5, state="disabled")
        self.custom_delimiter_entry.grid(row=3, column=1)

        convert_button = tk.Button(csv_converter_frame, text="Convert", command=self.convert, font=(self.font_family, 12))
        convert_button.grid(row=4, column=1, pady=10)

    def select_input_file(self):
        filename = filedialog.askopenfilename(filetypes=(("CSV files", "*.csv"), ("All files", "*.*")))
        self.input_entry.delete(0, tk.END)
        self.input_entry.insert(0, filename)

    def select_output_file(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        self.output_entry.delete(0, tk.END)
        self.output_entry.insert(0, filename)

    def convert(self):
        input_file = self.input_entry.get()
        output_file = self.output_entry.get()
        delimiter = self.delimiter_combobox.get()

        if not input_file or not output_file or not delimiter:  # If no file paths or delimiter are entered
            messagebox.showwarning("Warning", "Please enter input and output files and delimiter.")
            return

        if delimiter == "Custom Delimiter":
            delimiter = self.custom_delimiter_entry.get()

        if len(delimiter) != 1:  # If the delimiter is not a single character
            messagebox.showwarning("Warning", "Delimiter should be a single character.")
            return

        try:
            try:
                df = pd.read_csv(input_file, encoding='utf-8', engine='python', sep=delimiter)
            except UnicodeDecodeError:
                df = pd.read_csv(input_file, encoding='ISO-8859-1', engine='python', sep=delimiter)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read CSV file.\n{e}")
            return

        try:
            df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", "File has been converted successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to write Excel file.\n{e}")

    def handle_delimiter_selection(self, selected_delimiter):
        if selected_delimiter == "Custom Delimiter":
            self.custom_delimiter_entry.config(state="normal")
        else:
            self.custom_delimiter_entry.config(state="disabled")

    def _check_and_install(self, package):
        try:
            __import__(package)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

    def _browse_file(self, path_var, listbox):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if filename:
            path_var.set(filename)
            listbox.delete(0, 'end')
            listbox.insert('end', *pd.read_excel(filename).columns)

    def _browse_output_file(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if filename:
            self.output_file.set(filename)

    def _merge_files(self):
        if not all([self.file1_path.get(), self.file2_path.get(), self.file1_listbox.curselection(), self.file2_listbox.curselection(), self.output_file.get()]):
            messagebox.showerror("Error", "Please select both input files, common columns, and output file.")
            return

        file1_columns = [self.file1_listbox.get(i) for i in self.file1_listbox.curselection()]
        file2_columns = [self.file2_listbox.get(i) for i in self.file2_listbox.curselection()]

        try:
            df1 = pd.read_excel(self.file1_path.get())
            df2 = pd.read_excel(self.file2_path.get())
            merged_df = pd.merge(df1, df2, left_on=file1_columns, right_on=file2_columns)
            merged_df.to_excel(self.output_file.get(), index=False)
            self.log_area.insert(tk.END, "Merge completed successfully!\n")
            if self.open_output_file.get():
                self._open_output_file()
        except Exception as e:
            self.log_area.insert(tk.END, f"Merge failed. Error: {str(e)}\n")
            messagebox.showerror("Error", f"Merge failed. Error: {str(e)}")

    def _open_output_file(self):
        output_file = self.output_file.get()
        if output_file:
            try:
                os.startfile(output_file)
            except Exception as e:
                messagebox.showwarning("Warning", f"Failed to open output file: {str(e)}")

    def _open_how_to_window(self):
        how_to_text = """
        How to Use Excel Merger:
        
        1. Click on the "Choose File" button next to "File 1" to select the first input Excel file.
        2. Select one or more columns from the listbox below, which represent the common columns for merging.
        3. Repeat the above two steps for "File 2", selecting the second input Excel file and the corresponding columns.
        4. Click on the "Choose File" button next to "Output File" to specify the output file path and name.
        5. Once all the necessary inputs are selected, click the "Merge" button to start the merging process.
        6. The merged data will be saved to the specified output file.
        7. The progress and any error messages will be displayed in the log area at the bottom of the window.
        8. After a successful merge, the output file will be automatically opened for your convenience.
        
        Note: Ensure that the selected input files are in Excel format (.xlsx or .xls).
        """

        window = tk.Toplevel(self)
        window.title("How To Use Excel Merger")
        window.geometry("800x600")

        text_widget = tk.Text(window, font=(self.font_family, 14), bg=self.secondary_color, fg="black")
        text_widget.pack(expand=True, fill="both")
        text_widget.insert("1.0", how_to_text)
        text_widget.config(state="disabled")

    def _add_csv_to_excel_converter(self):
        # Create the frame for the CSV to Excel Converter
        csv_converter_frame = tk.LabelFrame(self, text="CSV to Excel Converter", bg='white', font=(self.font_family, 18))
        csv_converter_frame.pack(side='top', anchor='w', padx=10, pady=10)

        # Input CSV File
        input_label = tk.Label(csv_converter_frame, text="Input CSV File:", bg='white', font=(self.font_family, 12))
        input_label.grid(row=0, column=0)
        self.input_entry = tk.Entry(csv_converter_frame, width=30)
        self.input_entry.grid(row=0, column=1)
        input_button = tk.Button(csv_converter_frame, text="Browse...", command=self.select_input_file, font=(self.font_family, 12))
        input_button.grid(row=0, column=2)

        # Output Excel File
        output_label = tk.Label(csv_converter_frame, text="Output Excel File:", bg='white', font=(self.font_family, 12))
        output_label.grid(row=1, column=0)
        self.output_entry = tk.Entry(csv_converter_frame, width=30)
        self.output_entry.grid(row=1, column=1)
        output_button = tk.Button(csv_converter_frame, text="Browse...", command=self.select_output_file, font=(self.font_family, 12))
        output_button.grid(row=1, column=2)

        # CSV Delimiter
        delimiter_label = tk.Label(csv_converter_frame, text="CSV Delimiter:", bg='white', font=(self.font_family, 12))
        delimiter_label.grid(row=2, column=0)

        self.delimiter_combobox = ttk.Combobox(csv_converter_frame, values=[",", ";", "\t", "|", "Custom Delimiter"])
        self.delimiter_combobox.grid(row=2, column=1)
        self.delimiter_combobox.current(0)  # Set default selection
        self.delimiter_combobox.bind("<<ComboboxSelected>>", lambda event: self.handle_delimiter_selection(event.widget.get()))

        custom_delimiter_label = tk.Label(csv_converter_frame, text="Custom Delimiter:", bg='white', font=(self.font_family, 12))
        custom_delimiter_label.grid(row=3, column=0, sticky="e")
        self.custom_delimiter_entry = tk.Entry(csv_converter_frame, width=5, state="disabled")
        self.custom_delimiter_entry.grid(row=3, column=1)

        convert_button = tk.Button(csv_converter_frame, text="Convert", command=self.convert, font=(self.font_family, 12))
        convert_button.grid(row=4, column=1, pady=10)

    def select_input_file(self):
        filename = filedialog.askopenfilename(filetypes=(("CSV files", "*.csv"), ("All files", "*.*")))
        self.input_entry.delete(0, tk.END)
        self.input_entry.insert(0, filename)

    def select_output_file(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        self.output_entry.delete(0, tk.END)
        self.output_entry.insert(0, filename)

    def convert(self):
        input_file = self.input_entry.get()
        output_file = self.output_entry.get()
        delimiter = self.delimiter_combobox.get()

        if not input_file or not output_file or not delimiter:  # If no file paths or delimiter are entered
            messagebox.showwarning("Warning", "Please enter input and output files and delimiter.")
            return

        if delimiter == "Custom Delimiter":
            delimiter = self.custom_delimiter_entry.get()

        if len(delimiter) != 1:  # If the delimiter is not a single character
            messagebox.showwarning("Warning", "Delimiter should be a single character.")
            return

        try:
            try:
                df = pd.read_csv(input_file, encoding='utf-8', engine='python', sep=delimiter)
            except UnicodeDecodeError:
                df = pd.read_csv(input_file, encoding='ISO-8859-1', engine='python', sep=delimiter)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read CSV file.\n{e}")
            return

        try:
            df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", "File has been converted successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to write Excel file.\n{e}")

    def handle_delimiter_selection(self, selected_delimiter):
        if selected_delimiter == "Custom Delimiter":
            self.custom_delimiter_entry.config(state="normal")
        else:
            self.custom_delimiter_entry.config(state="disabled")

if __name__ == "__main__":
    app = ExcelMerger()
    app.mainloop()
