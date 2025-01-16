import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

class ExcelComparatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel File Comparator")
        self.root.geometry("700x500")

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
        ttk.Button(root, text="Compare Files", command=self.compare_files).pack(pady=10)

        # Output frame
        output_frame = ttk.LabelFrame(root, text="Output", padding=10)
        output_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.output_box = tk.Text(output_frame, height=10, wrap="word", font=("Arial", 10))
        self.output_box.pack(fill="both", expand=True)

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

            # Extract selected columns and reset index
            col1_series = df1[column1].dropna().reset_index(drop=True)
            col2_series = df2[column2].dropna().reset_index(drop=True)

            # Find differences
            only_in_file1 = col1_series[~col1_series.isin(col2_series)]
            only_in_file2 = col2_series[~col2_series.isin(col1_series)]

            # Display differences
            self.output_box.delete("1.0", tk.END)
            if only_in_file1.empty and only_in_file2.empty:
                self.output_box.insert(tk.END, "No differences found.")
            else:
                self.output_box.insert(tk.END, "Items only in File 1:\n")
                self.output_box.insert(tk.END, only_in_file1.to_string(index=False) + "\n\n")
                self.output_box.insert(tk.END, "Items only in File 2:\n")
                self.output_box.insert(tk.END, only_in_file2.to_string(index=False))

                # Ask user where to save the results
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx *.xls")],
                    title="Save Differences"
                )
                if save_path:
                    # Save to Excel
                    result_df = pd.DataFrame({
                        "Only in File 1": only_in_file1,
                        "Only in File 2": only_in_file2
                    })
                    result_df.to_excel(save_path, index=False)
                    messagebox.showinfo("Success", f"Differences saved to {save_path}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

# Initialize the application
root = tk.Tk()
app = ExcelComparatorApp(root)
root.mainloop()
