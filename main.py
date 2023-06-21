import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


class ExcelSorter:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Sorting Creation File Sorter")
        self.window.configure(bg="white")

        self.window.geometry("500x500")

        self.file_path = None
        self.sort_columns = []

        self.create_widgets()

    def create_widgets(self):
        self.file_label = tk.Label(self.window, text="", bg="white")
        self.file_label.pack(pady=10)

        self.sort_award_button = tk.Button(self.window, text="Sort Award File!", command=self.sort_award_file,
                                            font=("Times New Roman", 16, "bold"), bg="green", fg="white", width=20,
                                            height=2)
        self.sort_award_button.pack(pady=10)

        self.sort_backlog_button = tk.Button(self.window, text="Sort Backlog File!", command=self.sort_backlog_file,
                                            font=("Times New Roman", 16, "bold"), bg="blue", fg="white", width=20,
                                            height=2)
        self.sort_backlog_button.pack(pady=10)

    def select_file(self):
        file_path = filedialog.askopenfilename(title="Select Excel file",
                                               filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")))

        if file_path:
            self.file_path = file_path
            self.file_label.config(text=f"Selected File: {file_path}")
        else:
            self.file_path = None
            self.file_label.config(text="No file selected")

    def sort_award_file(self):
        self.select_file()

        if self.file_path:
            self.sort_columns = ['Product ID', 'Award Cust ID']
            self.sort_excel()

    def sort_backlog_file(self):
        self.select_file()

        if self.file_path:
            self.sort_columns = ['Product ID', 'Backlog Entry']
            self.sort_excel()

    def sort_excel(self):
        if not self.sort_columns:
            messagebox.showerror("Error", "No columns selected for sorting.")
            return

        try:
            # Read the Excel file into a pandas DataFrame
            df = pd.read_excel(self.file_path)

            # Sort the DataFrame based on the selected columns
            df = df.sort_values(by=self.sort_columns, ascending=[True, False])

            # Save the sorted DataFrame back to the Excel file
            df.to_excel(self.file_path, index=False)

            messagebox.showinfo("Success", "Excel file sorted and saved successfully.")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def run(self):
        self.window.mainloop()


# Create an instance of the ExcelSorter and run the program
sorter = ExcelSorter()
sorter.run()
