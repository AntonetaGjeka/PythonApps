import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

class ExcelComparatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel File Comparator")
        self.root.geometry("600x300")

        # File paths
        self.file1 = None
        self.file2 = None

        # Column variables
        self.columns_file1 = []
        self.columns_file2 = []
        self.selected_column1 = tk.StringVar()
        self.selected_column2 = tk.StringVar()

        # Styles
        style = ttk.Style()
        style.configure("TButton", font=("Arial", 12))
        style.configure("TLabel", font=("Arial", 11))
        style.configure("TCombobox", font=("Arial", 11))

        # File selection frame
        file_frame = ttk.LabelFrame(root, text="File Selection", padding=10)
        file_frame.pack(fill="x", padx=10, pady=10)

        ttk.Button(file_frame, text="Select File 1", command=self.load_file1).grid(row=0, column=0, padx=5, pady=5)
        self.file1_label = ttk.Label(file_frame, text="No file selected for File 1")
        self.file1_label.grid(row=0, column=1, sticky="w")

        ttk.Button(file_frame, text="Select File 2", command=self.load_file2).grid(row=1, column=0, padx=5, pady=5)
        self.file2_label = ttk.Label(file_frame, text="No file selected for File 2")
        self.file2_label.grid(row=1, column=1, sticky="w")

        # Column selection frame
        column_frame = ttk.LabelFrame(root, text="Column Selection", padding=10)
        column_frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(column_frame, text="Select column from File 1:").grid(row=0, column=0, sticky="w", padx=5)
        self.dropdown_file1 = ttk.Combobox(column_frame, textvariable=self.selected_column1, state="readonly")
        self.dropdown_file1.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(column_frame, text="Select column from File 2:").grid(row=1, column=0, sticky="w", padx=5)
        self.dropdown_file2 = ttk.Combobox(column_frame, textvariable=self.selected_column2, state="readonly")
        self.dropdown_file2.grid(row=1, column=1, padx=5, pady=5)

        # Compare button
        ttk.Button(root, text="Compare and Save", command=self.compare_files).pack(pady=5)

    def load_file1(self):
        self.file1 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.file1:
            self.file1_label.config(text=f"Selected File 1: {self.file1.split('/')[-1]}")
            df = pd.read_excel(self.file1)
            self.columns_file1 = df.columns.tolist()
            self.dropdown_file1.config(values=self.columns_file1)

    def load_file2(self):
        self.file2 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.file2:
            self.file2_label.config(text=f"Selected File 2: {self.file2.split('/')[-1]}")
            df = pd.read_excel(self.file2)
            self.columns_file2 = df.columns.tolist()
            self.dropdown_file2.config(values=self.columns_file2)

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

            # Find differences
            only_in_file1 = df1[~df1[column1].isin(df2[column2])]
            only_in_file2 = df2[~df2[column2].isin(df1[column1])]

            # Ask user where to save the results
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx *.xls")],
                title="Save Differences"
            )
            if save_path:
                # Save to Excel
                with pd.ExcelWriter(save_path) as writer:
                    only_in_file1.to_excel(writer, sheet_name="Only in File 1", index=False)
                    only_in_file2.to_excel(writer, sheet_name="Only in File 2", index=False)
                messagebox.showinfo("Success", f"Differences saved to {save_path}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

# Initialize the application
root = tk.Tk()
app = ExcelComparatorApp(root)
root.mainloop()
