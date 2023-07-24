import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import sqlite3

class ExcelToDBExporter:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Excel to Database Exporter")
        self.window.configure(bg="white")
        self.window.geometry("500x500")

        self.file_path = None

        self.create_widgets()

    def create_widgets(self):
        self.file_label = tk.Label(self.window, text="", bg="white")
        self.file_label.pack(pady=10)

        self.export_to_db_button = tk.Button(self.window, text="Export to Database", command=self.export_to_db,
                                            font=("Times New Roman", 16, "bold"), bg="green", fg="white", width=20,
                                            height=2)
        self.export_to_db_button.pack(pady=10)

    def select_file(self):
        file_path = filedialog.askopenfilename(title="Select Excel file",
                                               filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")))

        if file_path:
            self.file_path = file_path
            self.file_label.config(text=f"Selected File: {file_path}")
        else:
            self.file_path = None
            self.file_label.config(text="No file selected")

    def export_to_db(self):
        self.select_file()

        if self.file_path:
            try:
                # Read all sheets in the Excel file as a dictionary of DataFrames
                excel_data = pd.read_excel(self.file_path, sheet_name=None)

                # Ask the user to choose the location and name for the new database file
                db_file_path = filedialog.asksaveasfilename(title="Save Database File",
                                                            filetypes=(("SQLite Database", "*.db"),
                                                                       ("All files", "*.*")))
                if not db_file_path:
                    # User canceled the save dialog
                    return

                # Export each sheet to the SQLite database as a separate table
                with sqlite3.connect(db_file_path) as conn:
                    for sheet_name, sheet_data in excel_data.items():
                        # Use the sheet_name as the table name
                        sheet_data.to_sql(name=sheet_name, con=conn, index=False, if_exists="replace")

                messagebox.showinfo("Success", "Excel sheets exported to the database as separate tables.")

            except Exception as e:
                messagebox.showerror("Error", str(e))

    def run(self):
        self.window.mainloop()


# Create an instance of the ExcelToDBExporter and run the program
exporter = ExcelToDBExporter()
exporter.run()
