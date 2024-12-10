import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

class ExcelComparatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel File Comparator")

        # File paths
        self.file1 = None
        self.file2 = None

        # Dropdowns
        self.columns_file1 = []
        self.columns_file2 = []
        self.selected_column1 = tk.StringVar()
        self.selected_column2 = tk.StringVar()

        # File selection buttons
        tk.Button(root, text="Select File 1", command=self.load_file1).pack()
        tk.Button(root, text="Select File 2", command=self.load_file2).pack()

        # Placeholder labels for dropdowns (will be replaced after file loading)
        self.dropdown_file1_label = tk.Label(root, text="Select column from File 1")
        self.dropdown_file1_label.pack()
        self.dropdown_file2_label = tk.Label(root, text="Select column from File 2")
        self.dropdown_file2_label.pack()

        # Compare button
        tk.Button(root, text="Compare", command=self.compare_files).pack()

        # Output box for showing differences
        self.output_box = tk.Text(root, height=10, width=50)
        self.output_box.pack()

    def load_file1(self):
        self.file1 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.file1:
            df = pd.read_excel(self.file1)
            self.columns_file1 = df.columns.tolist()
            self.update_dropdown(self.dropdown_file1_label, self.selected_column1, self.columns_file1)

    def load_file2(self):
        self.file2 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.file2:
            df = pd.read_excel(self.file2)
            self.columns_file2 = df.columns.tolist()
            self.update_dropdown(self.dropdown_file2_label, self.selected_column2, self.columns_file2)

    def update_dropdown(self, dropdown_label, selected_var, column_list):
        # Replace the label with an OptionMenu
        dropdown_label.pack_forget()  # Remove the label
        dropdown = tk.OptionMenu(self.root, selected_var, *column_list)
        dropdown.pack()

    def compare_files(self):
        if not self.file1 or not self.file2:
            messagebox.showerror("Error", "Please select both files.")
            return

        column1 = self.selected_column1.get()
        column2 = self.selected_column2.get()

        if not column1 or not column2:
            messagebox.showerror("Error", "Please select columns to compare.")
            return

        try:
            df1 = pd.read_excel(self.file1)
            df2 = pd.read_excel(self.file2)

            # Extract the selected columns and reset the index
            col1_series = df1[column1].dropna().reset_index(drop=True)
            col2_series = df2[column2].dropna().reset_index(drop=True)

            # Find differences
            only_in_file1 = col1_series[~col1_series.isin(col2_series)]
            only_in_file2 = col2_series[~col2_series.isin(col1_series)]

            # Display differences
            self.output_box.delete("1.0", tk.END)  # Clear previous output
            if only_in_file1.empty and only_in_file2.empty:
                self.output_box.insert(tk.END, "No differences found.")
            else:
                self.output_box.insert(tk.END, "Items only in File 1:\n")
                self.output_box.insert(tk.END, only_in_file1.to_string(index=False) + "\n\n")
                self.output_box.insert(tk.END, "Items only in File 2:\n")
                self.output_box.insert(tk.END, only_in_file2.to_string(index=False))

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

# Initialize the application
root = tk.Tk()
app = ExcelComparatorApp(root)
root.mainloop()
