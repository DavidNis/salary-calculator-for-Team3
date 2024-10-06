import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime, timedelta
from ttkthemes import ThemedTk
from salary_calc import SalaryCalculator 

class SalaryGui:
    def __init__(self, root):
        self.root = root
        self.root.title('Salary Calculator with Excel Import')
        self.root.geometry("1000x700")
        self.root.set_theme("arc")

        self.df = None  # to hold the Excel data
        self.salary_calculator = SalaryCalculator()

        self.create_widgets()




    def create_widgets(self):
        """
        Create the widgets for the GUI
        """
        # button to load an Excel file
        ttk.Button(self.root, text="Load Excel File", command=self.load_excel_file).grid(row=0, column=0, pady=10)

        # Treeview to display data
        self.tree = ttk.Treeview(self.root, columns=("Date", "Role", "Entry Time", "Exit Time", "Control Room", "Pay"), show='headings', selectmode='browse')
        self.tree.heading("Date", text="Date")
        self.tree.heading("Role", text="Role")
        self.tree.heading("Entry Time", text="Entry Time")
        self.tree.heading("Exit Time", text="Exit Time")
        self.tree.heading("Control Room", text="Control Room")
        self.tree.heading("Pay", text="Pay")
        self.tree.column("Date", width=100)
        self.tree.column("Role", width=200)
        self.tree.column("Entry Time", width=100)
        self.tree.column("Exit Time", width=100)
        self.tree.column("Control Room", width=100)
        self.tree.column("Pay", width=100)
        self.tree.grid(row=1, column=0, columnspan=6, padx=10, pady=10)

        # event to handle row selection
        self.tree.bind('<<TreeviewSelect>>', self.on_row_select)

        # form to manually edit row details
        ttk.Label(self.root, text="Date:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.date_var = tk.StringVar()
        self.date_entry = ttk.Entry(self.root, textvariable=self.date_var, state='readonly')
        self.date_entry.grid(row=2, column=1, padx=10, pady=5)

        ttk.Label(self.root, text="Role:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.role_var = tk.StringVar()
        self.role_entry = ttk.Entry(self.root, textvariable=self.role_var, state='readonly')
        self.role_entry.grid(row=3, column=1, padx=10, pady=5)

        ttk.Label(self.root, text="Entry Time (HH:MM):").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        self.entry_time_var = tk.StringVar()
        self.entry_time_entry = ttk.Entry(self.root, textvariable=self.entry_time_var)
        self.entry_time_entry.grid(row=4, column=1, padx=10, pady=5)

        ttk.Label(self.root, text="Exit Time (HH:MM):").grid(row=5, column=0, sticky="w", padx=10, pady=5)
        self.exit_time_var = tk.StringVar()
        self.exit_time_entry = ttk.Entry(self.root, textvariable=self.exit_time_var)
        self.exit_time_entry.grid(row=5, column=1, padx=10, pady=5)

        self.control_room_var = tk.BooleanVar()
        ttk.Checkbutton(self.root, text="In Control Room", variable=self.control_room_var).grid(row=6, column=0, columnspan=2, padx=10, pady=5)

        # button to update selected row
        ttk.Button(self.root, text="Update Row", command=self.update_row).grid(row=7, column=0, columnspan=2, pady=10)

        # button to calculate the pay
        ttk.Button(self.root, text="Calculate Pay for the day", command=self.calculate_pay).grid(row=8, column=0, columnspan=6, pady=10)

        # button to calculate the total pay
        ttk.Button(self.root, text="Calculate Total Pay", command=self.calculate_total_pay).grid(row=9, column=0, columnspan=6, pady=10)





    def load_excel_file(self):
        """
        Load an Excel file and populate the treeview with the data.
        """
        # select an Excel file
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        
        if not file_path:
            return  

        # check if the file is valid, should only be an Excel file
        if not file_path.endswith(('.xlsx', '.xls')):
            messagebox.showerror("Invalid File", "Please upload a valid Excel file (.xlsx or .xls).")
            return

        try:
            # load Excel data using pandas, skipping the first 4 rows if necessary
            self.df = pd.read_excel(file_path, skiprows=4, header=0)

            # rename columns from Hebrew to English to handle the code more easily instead of using Hebrew
            column_mapping = {
                'תאריך': 'Date',
                'תפקיד': 'Role',
                'כניסה': 'Entry Time',
                'יציאה': 'Exit Time',
                'סיכום': 'Summary'
            }
            self.df.rename(columns=column_mapping, inplace=True)

            # check if the necessary columns (date and role) exist in the Excel file
            required_columns = {'Date', 'Role'}
            if not required_columns.issubset(self.df.columns):
                messagebox.showerror("Invalid Format", f"Excel file must contain columns: {', '.join(required_columns)}. Current columns are: {', '.join(self.df.columns)}")
                return

            # populate the treeview with the data, removing rows where Role is "N/A" - means the person didn't work that day
            for item in self.tree.get_children():
                self.tree.delete(item)  # Clear existing items in the treeview before loading new data

            for index, row in self.df.iterrows():
                if pd.isna(row['Role']) or row['Role'] == 'N/A':
                    continue  # Skip rows where Role is "N/A"

                date_str = pd.to_datetime(row['Date']).strftime('%Y-%m-%d') if pd.notna(row['Date']) else "Missing"
                role = row['Role']
                entry_time = row['Entry Time'] if pd.notna(row['Entry Time']) else "Missing"
                exit_time = row['Exit Time'] if pd.notna(row['Exit Time']) else "Missing"
                self.tree.insert("", "end", values=(date_str, role, entry_time, exit_time, "No", ""))  # Default "Control Room" to "No"

        except Exception as e:
            return
            # messagebox.showerror("Error", f"Error loading Excel file: {e}") # TODO: check if needed later.




    def on_row_select(self, event):
        """
        Event handler when a row is selected in the tree view.
        """
        # get selected row
        selected_item = self.tree.selection()
        if not selected_item:
            return

        values = self.tree.item(selected_item, 'values')
        self.date_var.set(values[0])
        self.role_var.set(values[1])
        self.entry_time_var.set(values[2])
        self.exit_time_var.set(values[3])
        self.control_room_var.set(True if values[4] == "Yes" else False)




    def update_row(self):
        """
        Update the selected row in the treeview with the values from the form.
        """
        # get the selected item from the treeview
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("No Selection", "Please select a row to update.")
            return

        # update the values in the treeview
        date_str = self.date_var.get()
        role = self.role_var.get()
        entry_time = self.entry_time_var.get()
        exit_time = self.exit_time_var.get()
        in_control_room = "Yes" if self.control_room_var.get() else "No"

        self.tree.item(selected_item, values=(date_str, role, entry_time, exit_time, in_control_room, ""))




    def calculate_pay(self):
        """
        Calculate the pay for each row in the treeview.
        """
        if self.df is None:
            messagebox.showerror("No Data", "Please load an Excel file first.")
            return

        try:
            total_pay = 0.0  # To store the sum of all non-zero pays

            for item in self.tree.get_children():
                values = self.tree.item(item, 'values')
                date_str, role, entry_time_str, exit_time_str, in_control_room, pay = values

                # If entry or exit time is missing, set pay to zero
                if entry_time_str == "Missing" or exit_time_str == "Missing":
                    total_pay_day = 0.0
                else:
                    # Use SalaryCalculator to calculate pay
                    date = datetime.strptime(date_str, "%Y-%m-%d").date()
                    start_time = datetime.strptime(entry_time_str, "%H:%M").time()
                    end_time = datetime.strptime(exit_time_str, "%H:%M").time()
                    in_control_room_flag = True if in_control_room == "Yes" else False

                    # Calculate the pay for this day using the SalaryCalculator
                    self.salary_calculator.add_work_day(
                        date=date,
                        start_time=start_time,
                        end_time=end_time,
                        in_control_room=in_control_room_flag,
                        is_friday=False,  # Assume other conditions are False, adjust as needed
                        is_saturday=False,
                        is_holiday=False,
                        is_last_day_of_holiday=False
                    )
                    total_pay_day = self.salary_calculator.total_pay()

                # Update Treeview with the calculated pay
                self.tree.item(item, values=(date_str, role, entry_time_str, exit_time_str, in_control_room, f"{total_pay_day:.2f}"))

                # Add to the total pay if the day's pay is greater than zero
                if total_pay_day > 0:
                    total_pay += total_pay_day

            # messagebox.showinfo("Success", f"Pay calculated successfully for all entries.\nTotal Pay: {total_pay:.2f} shekels.")
        except Exception as e:
            messagebox.showerror("Error", f"Error calculating pay: {e}")


    def calculate_total_pay(self):
        """"
        Calculate the total pay for all days in the treeview.
        """
        if self.df is None:
            messagebox.showerror("No Data", "Please load an Excel file first.")
            return

        try:
            total_pay = 0.0  # To store the sum of all non-zero pays

            for item in self.tree.get_children():
                values = self.tree.item(item, 'values')
                pay = values[5]
                total_pay += float(pay)

            messagebox.showinfo("Total Pay", f"Total Pay: {total_pay:.2f} shekels.")
        except Exception as e:
            messagebox.showerror("Error", f"Error calculating total pay: {e}")
        

if __name__ == "__main__":
    root = ThemedTk(theme="arc")
    app = SalaryGui(root)
    root.mainloop()
