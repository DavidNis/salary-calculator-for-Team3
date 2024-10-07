import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkFont
import pandas as pd
from datetime import datetime
from ttkthemes import ThemedTk
from salary_calc import SalaryCalculator
import threading


class SalaryGui:
    def __init__(self, root):
        self.root = root
        self.root.title('Salary Calculator with Excel Import')
        self.root.geometry("1200x900")
        self.root.set_theme("arc")

        self.df = None
        self.salary_calculator = SalaryCalculator()

        # Define a custom font for widgets with increased size and bold weight
        self.custom_font = tkFont.Font(family="Rubik", size=14)

        # Create a custom style
        self.style = ttk.Style()
        self.style.configure("Custom.TButton", font=self.custom_font)
        self.style.configure("Custom.TLabel", font=self.custom_font)
        self.style.configure("Custom.TCheckbutton", font=self.custom_font)
        self.style.configure("Custom.TLabelframe.Label", font=self.custom_font)

        self.create_widgets()






    def create_widgets(self):
        # Frame for loading the Excel file
        load_frame = ttk.Frame(self.root)
        load_frame.pack(pady=15)
        ttk.Button(load_frame, text="Load Excel File", command=self.load_excel_file, style="Custom.TButton").pack()

        # Frame for Treeview (with a scrollbar)
        tree_frame = ttk.Frame(self.root)
        tree_frame.pack(pady=15, fill='x')

        # Scrollbar for the Treeview
        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Treeview to display data
        self.tree = ttk.Treeview(tree_frame, columns=("Date","Day of Week", "Role", "Entry Time", "Exit Time", "Control Room", "Pay"),
                                 show='headings', selectmode='browse', yscrollcommand=tree_scroll.set)
        self.tree.heading("Date", text="Date", anchor='center')
        self.tree.heading("Day of Week", text="Day of Week", anchor='center')
        self.tree.heading("Role", text="Role", anchor='center')
        self.tree.heading("Entry Time", text="Entry Time", anchor='center')
        self.tree.heading("Exit Time", text="Exit Time", anchor='center')
        self.tree.heading("Control Room", text="Control Room", anchor='center')
        self.tree.heading("Pay", text="Pay", anchor='center')

        self.tree.column("Date", anchor='center', width=100)
        self.tree.column("Day of Week", anchor='center', width=120)
        self.tree.column("Role", anchor='center', width=200)
        self.tree.column("Entry Time", anchor='center', width=100)
        self.tree.column("Exit Time", anchor='center', width=100)
        self.tree.column("Control Room", anchor='center', width=100)
        self.tree.column("Pay", anchor='center', width=100)

        self.tree.pack(fill='x', padx=15)
        tree_scroll.config(command=self.tree.yview)

        # Event to handle row selection
        self.tree.bind('<<TreeviewSelect>>', self.on_row_select)

        # Frame for editing selected row
        edit_frame = ttk.LabelFrame(self.root, text="Edit Selected Row", style="Custom.TLabelframe")
        edit_frame.pack(pady=20, padx=20, fill='x', expand='yes')


        # Entry for date
        ttk.Label(edit_frame, text="Date:", style="Custom.TLabel").grid(row=0, column=0, sticky="w", padx=15, pady=10)
        self.date_var = tk.StringVar()
        self.date_entry = ttk.Entry(edit_frame, textvariable=self.date_var, state='readonly', width=25, font=self.custom_font)
        self.date_entry.grid(row=0, column=1, padx=15, pady=10)


        # Entry for day of the week
        ttk.Label(edit_frame, text="Day of the Week:", style="Custom.TLabel").grid(row=0, column=2, sticky="w", padx=15, pady=10)
        self.date_of_week_var = tk.StringVar()
        self.date_of_week_entry = ttk.Entry(edit_frame, textvariable=self.date_of_week_var, width=25, font=self.custom_font)
        self.date_of_week_entry.grid(row=0, column=3, padx=15, pady=10)


        # Entry for role
        ttk.Label(edit_frame, text="Role:", style="Custom.TLabel").grid(row=1, column=0, sticky="w", padx=15, pady=10)
        self.role_var = tk.StringVar()
        self.role_entry = ttk.Entry(edit_frame, textvariable=self.role_var, width=25, font=self.custom_font)
        self.role_entry.grid(row=1, column=1, padx=15, pady=10)


        # Entry for start time
        ttk.Label(edit_frame, text="Entry Time (HH:MM):", style="Custom.TLabel").grid(row=2, column=0, sticky="w", padx=15, pady=10)
        self.entry_time_var = tk.StringVar()
        self.entry_time_entry = ttk.Entry(edit_frame, textvariable=self.entry_time_var, width=25, font=self.custom_font)
        self.entry_time_entry.grid(row=2, column=1, padx=15, pady=10)

        # Entry for exit time
        ttk.Label(edit_frame, text="Exit Time (HH:MM):", style="Custom.TLabel").grid(row=3, column=0, sticky="w", padx=15, pady=10)
        self.exit_time_var = tk.StringVar()
        self.exit_time_entry = ttk.Entry(edit_frame, textvariable=self.exit_time_var, width=25, font=self.custom_font)
        self.exit_time_entry.grid(row=3, column=1, padx=15, pady=10)



        # Checkbutton for control room
        self.control_room_var = tk.BooleanVar()
        ttk.Checkbutton(edit_frame, text="In Control Room", variable=self.control_room_var, 
                command=self.update_control_room, style="Custom.TCheckbutton").grid(row=4, column=0, columnspan=2, padx=15, pady=10)

        # Checkbutton for holiday
        self.holiday_var = tk.BooleanVar()
        ttk.Checkbutton(edit_frame, text="Holiday", variable=self.holiday_var, style="Custom.TCheckbutton").grid(row=4, column=1, columnspan=2, padx=15, pady=10)

        # Checkbutton for last day of holiday
        self.last_day_holiday_var = tk.BooleanVar()
        ttk.Checkbutton(edit_frame, text="Last Day of Holiday", variable=self.last_day_holiday_var, style="Custom.TCheckbutton").grid(row=4, column=2, columnspan=2, padx=15, pady=10)

        # Button to update selected row
        ttk.Button(edit_frame, text="Update Selected Row", command=self.update_row, style="Custom.TButton").grid(row=5, column=0, columnspan=2, pady=20)

        # Frame for action buttons
        action_frame = ttk.Frame(self.root)
        action_frame.pack(pady=25)

        # Button to calculate the pay for each day
        calculate_pay_button = ttk.Button(action_frame, text="Calculate Pay for Each Day", command=self.calculate_pay, style="Custom.TButton")
        calculate_pay_button.grid(row=0, column=0, padx=20, pady=10)

        # Button to calculate the total pay
        total_pay_button = ttk.Button(action_frame, text="Calculate Total Pay", command=self.calculate_total_pay, style="Custom.TButton")
        total_pay_button.grid(row=0, column=1, padx=20, pady=10)

        # Label to display the total pay
        ttk.Label(action_frame, text="Total Pay:", style="Custom.TLabel").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        self.total_pay_var = tk.StringVar(value="0.00")
        self.total_pay_label = ttk.Label(action_frame, textvariable=self.total_pay_var, style="Custom.TLabel")
        self.total_pay_label.grid(row=1, column=1, padx=10, pady=10, sticky="w")

    
    
    
    def update_control_room(self):
        """
        Update the "Control Room" value in the selected row based on the checkbox state.
        """
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("No Selection", "Please select a row to update.")
            return

        # Get the value of the checkbox and set "Yes" or "No"
        in_control_room = "Yes" if self.control_room_var.get() else "No"

        # Get existing values from the selected item
        values = self.tree.item(selected_item, 'values')

        # Create a new tuple with the updated "Control Room" value
        updated_values = (
            values[0],  # Date
            values[1],  # Day of Week
            values[2],  # Role
            values[3],  # Entry Time
            values[4],  # Exit Time
            in_control_room,  # Updated Control Room value
            values[6]  # Pay
        )

        # Update the Treeview with the new values
        self.tree.item(selected_item, values=updated_values)




    def load_excel_file(self):
        """
        Load an Excel file and populate the treeview with the data.
        """
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        
        if not file_path:
            return

        if not file_path.endswith(('.xlsx', '.xls')):
            messagebox.showerror("Invalid File", "Please upload a valid Excel file (.xlsx or .xls).")
            return

        try:
            # Offload heavy file loading to a separate thread to keep the GUI responsive
            threading.Thread(target=self._load_file_thread, args=(file_path,)).start()
        except Exception as e:
            messagebox.showerror("Error", f"Error loading Excel file: {e}")






    def _load_file_thread(self, file_path):
        try:
            # Load the file and rename columns
            self.df = pd.read_excel(file_path, skiprows=4, header=0, engine='openpyxl')

            column_mapping = {
                'תאריך': 'Date',
                'תפקיד': 'Role',
                'כניסה': 'Entry Time',
                'יציאה': 'Exit Time',
                'סיכום': 'Summary'
            }
            self.df.rename(columns=column_mapping, inplace=True)

            # Check for required columns
            required_columns = {'Date', 'Role'}
            if not required_columns.issubset(self.df.columns):
                messagebox.showerror("Invalid Format", f"Excel file must contain columns: {', '.join(required_columns)}. Current columns are: {', '.join(self.df.columns)}")
                return

            # Update the Treeview on the main thread
            self.root.after(0, self._update_treeview)

        except Exception as e:
            messagebox.showerror("Error", f"Error loading Excel file: {e}")







    def _update_treeview(self):
        """
        Update the Treeview with the data from the loaded Excel file.
        """

        # clear any existing data in the Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Populate the Treeview
        for index, row in self.df.iterrows():
            if pd.isna(row['Role']) or row['Role'] == 'N/A':
                continue

            # Format the date
            date_str = pd.to_datetime(row['Date']).strftime('%Y-%m-%d') if pd.notna(row['Date']) else "Missing"
            role = row['Role']

            # Format entry and exit times to HH:MM, handle potential parsing issues
            try:
                entry_time = datetime.strptime(str(row['Entry Time']), '%H:%M:%S').strftime('%H:%M') if pd.notna(row['Entry Time']) else "Missing"
            except ValueError:
                entry_time = "Missing"

            try:
                exit_time = datetime.strptime(str(row['Exit Time']), '%H:%M:%S').strftime('%H:%M') if pd.notna(row['Exit Time']) else "Missing"
            except ValueError:
                exit_time = "Missing"

            # Calculate the day of the week from the date
            if pd.notna(row['Date']):
                day_of_week = pd.to_datetime(row['Date']).strftime('%A')  # Get full weekday name
            else:
                day_of_week = "Missing"

            # Insert the row into the Treeview, including the day of the week
            self.tree.insert("", "end", values=(date_str, day_of_week, role, entry_time, exit_time, "No", ""))





    def on_row_select(self, event):
        """
        Handle the event when a row is selected in the Tree.
        """
        selected_item = self.tree.selection()
        if not selected_item:
            return

        values = self.tree.item(selected_item, 'values')

        # extract values from each column correctly
        self.date_var.set(values[0])
        self.role_var.set(values[2])
        self.entry_time_var.set(values[3]) 
        self.exit_time_var.set(values[4])  

        # update the "In Control Room" checkbox
        self.control_room_var.set(True if values[5] == "Yes" else False)







    def update_row(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("No Selection", "Please select a row to update.")
            return

        # Get the updated values from the entry fields and checkbox
        date_str = self.date_var.get()
        role = self.role_var.get()
        entry_time = self.entry_time_var.get()
        exit_time = self.exit_time_var.get()
        in_control_room = "Yes" if self.control_room_var.get() else "No"  # Get the correct value based on checkbox state

        # Update Treeview with new values
        day_of_week = self.get_day_of_week(date_str)  # Get the day of the week for the updated date
        self.tree.item(selected_item, values=(date_str, day_of_week, role, entry_time, exit_time, in_control_room, ""))

    # Helper function to get day of the week
    def get_day_of_week(self, date_str):
        try:
            return datetime.strptime(date_str, "%Y-%m-%d").strftime('%A')
        except ValueError:
            return "Missing"







    def calculate_pay(self):
        """
        Calculate the pay for each row in the treeview and update the 'Pay' column.
        """
        if self.df is None:
            messagebox.showerror("No Data", "Please load an Excel file first.")
            return

        try:
            for item in self.tree.get_children():
                values = self.tree.item(item, 'values')
                # The values tuple has 7 elements now: Date, DayOfWeek, Role, Entry Time, Exit Time, Control Room, Pay
                date_str, day_of_week, role, entry_time_str, exit_time_str, in_control_room, _ = values

                if entry_time_str == "Missing" or exit_time_str == "Missing":
                    total_pay_day = 0.0
                else:
                    # Parse the date, start time, and end time
                    date = datetime.strptime(date_str, "%Y-%m-%d").date()
                    start_time = datetime.strptime(entry_time_str, "%H:%M").time()
                    end_time = datetime.strptime(exit_time_str, "%H:%M").time()
                    in_control_room_flag = True if in_control_room == "Yes" else False

                    # Determine if the workday is a Friday, Saturday, holiday, or last day of a holiday
                    is_friday = date.weekday() == 4  # 0 = Monday, 4 = Friday
                    is_saturday = date.weekday() == 5  # 5 = Saturday
                    is_holiday = self.holiday_var.get()  # From checkbox
                    is_last_day_of_holiday = self.last_day_holiday_var.get()  # From checkbox

                    # Use a temporary instance of SalaryCalculator to avoid accumulating pay
                    temp_calculator = SalaryCalculator()
                    temp_calculator.add_work_day(
                        date=date,
                        start_time=start_time,
                        end_time=end_time,
                        in_control_room=in_control_room_flag,
                        is_friday=is_friday,
                        is_saturday=is_saturday,
                        is_holiday=is_holiday,
                        is_last_day_of_holiday=is_last_day_of_holiday
                    )
                    total_pay_day = temp_calculator.total_pay()

                # Update Treeview with the calculated pay for that specific day
                self.tree.item(item, values=(date_str, day_of_week, role, entry_time_str, exit_time_str, in_control_room, f"{total_pay_day:.2f}"))

        except Exception as e:
            messagebox.showerror("Error", f"Error calculating pay: {e}")






    def calculate_total_pay(self):
        """
        Calculate the total pay for all days in the treeview.
        """
        total_pay = 0.0  # To store the sum of all valid pays

        try:
            for item in self.tree.get_children():
                values = self.tree.item(item, 'values')
                pay = values[6]  # Assuming 'Pay' is the seventh column (index 6)

                # Attempt to convert the pay value to float if it's not "Missing" or empty
                if pay != "Missing" and pay != "" and pay != "No":
                    try:
                        total_pay += float(pay)
                    except ValueError:
                        # Ignore any value that cannot be converted to a float
                        continue

            # Show the total pay in a message box
            messagebox.showinfo("Total Pay", f"Total Pay for all entries: {total_pay:.2f} shekels.")
        except Exception as e:
            messagebox.showerror("Error", f"Error calculating total pay: {e}")


if __name__ == "__main__":
    root = ThemedTk(theme="arc")
    app = SalaryGui(root)
    root.mainloop()

    
