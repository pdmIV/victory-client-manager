import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime, timedelta
import os

# Change this to the exact name of your Excel file in the same directory
EXCEL_FILE = "investments.xlsx"

# COLUMN CONSTANTS
COLUMN_FIRST_NAME = "First Name"
COLUMN_LAST_NAME = "Last Name"
COLUMN_ORIGIN_DATE = "Note Origin Date"
COLUMN_MATURITY_DATE = "Note Maturity Date"
COLUMN_PRINCIPAL = "Principal"
COLUMN_INTEREST_RATE = "Interest Rate"
COLUMN_PRINCIPAL_PLUS_INTEREST = "Principal + Interest"
COLUMN_PROJECT_NAME = "Project Name"   # NEW COLUMN

def load_investments():
    """
    Reads the Excel file into a pandas DataFrame.
    If the file doesn't exist or is empty, returns an empty DataFrame with necessary columns.
    """
    if not os.path.exists(EXCEL_FILE):
        # Create an empty DataFrame with all the columns (including Project Name)
        df = pd.DataFrame(columns=[
            COLUMN_FIRST_NAME, COLUMN_LAST_NAME, 
            COLUMN_ORIGIN_DATE, COLUMN_MATURITY_DATE, 
            COLUMN_PRINCIPAL, COLUMN_INTEREST_RATE, 
            COLUMN_PRINCIPAL_PLUS_INTEREST, COLUMN_PROJECT_NAME
        ])
        df.to_excel(EXCEL_FILE, index=False)
        return df
    
    df = pd.read_excel(EXCEL_FILE)
    
    # Ensure all required columns exist. If they don't, add them.
    required_cols = [
        COLUMN_FIRST_NAME, COLUMN_LAST_NAME, COLUMN_ORIGIN_DATE, 
        COLUMN_MATURITY_DATE, COLUMN_PRINCIPAL, COLUMN_INTEREST_RATE, 
        COLUMN_PRINCIPAL_PLUS_INTEREST, COLUMN_PROJECT_NAME
    ]
    for col in required_cols:
        if col not in df.columns:
            df[col] = None

    return df


def save_investments(df):
    """Writes the pandas DataFrame to the Excel file."""
    df.to_excel(EXCEL_FILE, index=False)


def calculate_maturity_date(origin_date_str):
    """
    Given a string of note origin date (e.g. '2024-02-15'), 
    returns a string for maturity date ~9 months later.
    """
    # Adjust parsing as needed for your date formats
    try:
        origin_date = datetime.strptime(origin_date_str, "%Y-%m-%d")
    except ValueError:
        # Attempt different common format: e.g. 'MM/DD/YYYY'
        origin_date = datetime.strptime(origin_date_str, "%m/%d/%Y")
    maturity_date = origin_date + timedelta(days=9*30)  # approximate 9 months as 270 days
    return maturity_date.strftime("%Y-%m-%d")


def calculate_principal_plus_interest(principal, interest_rate):
    """
    principal: float
    interest_rate: float (representing annual rate, e.g. 0.07 for 7%)
    Approximates 9-month interest using simple interest: principal * rate * (9/12).
    """
    try:
        principal = float(principal)
        rate = float(interest_rate)
        interest = principal * rate * (9/12)
        total = principal + interest
        return round(total, 2)
    except:
        return None


class InvestmentApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Client Investments Manager")
        
        # Load the data from Excel
        self.df = load_investments()

        # Main Frame
        self.main_frame = ttk.Frame(self.master, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Define the columns that the Treeview will show
        self.columns = [
            COLUMN_FIRST_NAME, 
            COLUMN_LAST_NAME, 
            COLUMN_PROJECT_NAME,   # SHOW PROJECT NAME HERE
            COLUMN_ORIGIN_DATE, 
            COLUMN_MATURITY_DATE, 
            COLUMN_PRINCIPAL, 
            COLUMN_INTEREST_RATE, 
            COLUMN_PRINCIPAL_PLUS_INTEREST
        ]

        # Treeview for showing data
        self.tree = ttk.Treeview(self.main_frame, columns=self.columns, show="headings", height=15)
        
        self.tree.heading(COLUMN_FIRST_NAME, text="First Name")
        self.tree.heading(COLUMN_LAST_NAME, text="Last Name")
        self.tree.heading(COLUMN_PROJECT_NAME, text="Project Name")    # NEW
        self.tree.heading(COLUMN_ORIGIN_DATE, text="Origin Date")
        self.tree.heading(COLUMN_MATURITY_DATE, text="Maturity Date")
        self.tree.heading(COLUMN_PRINCIPAL, text="Principal")
        self.tree.heading(COLUMN_INTEREST_RATE, text="Interest Rate")
        self.tree.heading(COLUMN_PRINCIPAL_PLUS_INTEREST, text="Principal + Interest")
        
        for col in self.columns:
            self.tree.column(col, width=120, anchor=tk.CENTER)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbar
        scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Populate the tree
        self.load_tree()

        # Buttons frame
        self.btn_frame = ttk.Frame(self.master, padding="10")
        self.btn_frame.pack(fill=tk.X)

        self.add_btn = ttk.Button(self.btn_frame, text="Add Entry", command=self.add_entry)
        self.add_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.edit_btn = ttk.Button(self.btn_frame, text="Edit Entry", command=self.edit_entry)
        self.edit_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.del_btn = ttk.Button(self.btn_frame, text="Delete Entry", command=self.del_entry)
        self.del_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.save_btn = ttk.Button(self.btn_frame, text="Save to Excel", command=self.save_changes)
        self.save_btn.pack(side=tk.LEFT, padx=(0, 5))

    def load_tree(self):
        # Clear current rows
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Insert data from self.df
        for _, row_data in self.df.iterrows():
            self.tree.insert("", tk.END, values=(
                row_data.get(COLUMN_FIRST_NAME, ""),
                row_data.get(COLUMN_LAST_NAME, ""),
                row_data.get(COLUMN_PROJECT_NAME, ""),   # NEW
                row_data.get(COLUMN_ORIGIN_DATE, ""),
                row_data.get(COLUMN_MATURITY_DATE, ""),
                row_data.get(COLUMN_PRINCIPAL, ""),
                row_data.get(COLUMN_INTEREST_RATE, ""),
                row_data.get(COLUMN_PRINCIPAL_PLUS_INTEREST, "")
            ))

    def add_entry(self):
        """Opens a new window to add a new investment entry."""
        EntryWindow(self.master, self, mode="add")

    def edit_entry(self):
        """Opens a new window to edit the selected investment entry."""
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Select an entry to edit.")
            return
        item_values = self.tree.item(selected_item[0], "values")
        EntryWindow(self.master, self, mode="edit", item_values=item_values)

    def del_entry(self):
        """Deletes the selected investment entry from the DataFrame and refreshes the table."""
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Select an entry to delete.")
            return
        
        confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete?")
        if confirm:
            # Remove from DataFrame
            item_values = self.tree.item(selected_item[0], "values")
            condition = (
                (self.df[COLUMN_FIRST_NAME] == item_values[0]) & 
                (self.df[COLUMN_LAST_NAME] == item_values[1]) &
                (self.df[COLUMN_PROJECT_NAME] == item_values[2]) &    # NEW
                (self.df[COLUMN_ORIGIN_DATE].astype(str) == str(item_values[3])) &
                (self.df[COLUMN_MATURITY_DATE].astype(str) == str(item_values[4])) &
                (self.df[COLUMN_PRINCIPAL].astype(str) == str(item_values[5])) &
                (self.df[COLUMN_INTEREST_RATE].astype(str) == str(item_values[6])) &
                (self.df[COLUMN_PRINCIPAL_PLUS_INTEREST].astype(str) == str(item_values[7]))
            )
            self.df.drop(self.df[condition].index, inplace=True)
            
            # Remove from tree
            self.tree.delete(selected_item[0])

    def save_changes(self):
        """Saves the current DataFrame to the Excel file."""
        save_investments(self.df)
        messagebox.showinfo("Saved", "Changes saved to Excel file successfully.")


class EntryWindow(tk.Toplevel):
    def __init__(self, master, app, mode="add", item_values=None):
        super().__init__(master)
        self.app = app
        self.mode = mode
        self.item_values = item_values

        if self.mode == "add":
            self.title("Add New Entry")
        else:
            self.title("Edit Entry")

        # Labels
        lbl_fn = ttk.Label(self, text="First Name:")
        lbl_ln = ttk.Label(self, text="Last Name:")
        lbl_project = ttk.Label(self, text="Project Name:")  # NEW
        lbl_origin = ttk.Label(self, text="Note Origin Date (YYYY-MM-DD):")
        lbl_principal = ttk.Label(self, text="Principal:")
        lbl_interest = ttk.Label(self, text="Interest Rate (decimal):")

        # Entry widgets
        self.entry_fn = ttk.Entry(self)
        self.entry_ln = ttk.Entry(self)
        self.entry_project = ttk.Entry(self)  # NEW
        self.entry_origin = ttk.Entry(self)
        self.entry_principal = ttk.Entry(self)
        self.entry_interest = ttk.Entry(self)

        # Grid layout
        lbl_fn.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        lbl_ln.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        lbl_project.grid(row=2, column=0, padx=5, pady=5, sticky="e")  # NEW
        lbl_origin.grid(row=3, column=0, padx=5, pady=5, sticky="e")
        lbl_principal.grid(row=4, column=0, padx=5, pady=5, sticky="e")
        lbl_interest.grid(row=5, column=0, padx=5, pady=5, sticky="e")

        self.entry_fn.grid(row=0, column=1, padx=5, pady=5)
        self.entry_ln.grid(row=1, column=1, padx=5, pady=5)
        self.entry_project.grid(row=2, column=1, padx=5, pady=5)  # NEW
        self.entry_origin.grid(row=3, column=1, padx=5, pady=5)
        self.entry_principal.grid(row=4, column=1, padx=5, pady=5)
        self.entry_interest.grid(row=5, column=1, padx=5, pady=5)

        # If edit mode, populate fields
        if self.mode == "edit" and self.item_values:
            # item_values structure:
            # 0: First Name
            # 1: Last Name
            # 2: Project Name
            # 3: Origin Date
            # 4: Maturity Date
            # 5: Principal
            # 6: Interest Rate
            # 7: Principal + Interest
            self.entry_fn.insert(0, self.item_values[0])
            self.entry_ln.insert(0, self.item_values[1])
            self.entry_project.insert(0, self.item_values[2])  # NEW
            self.entry_origin.insert(0, self.item_values[3])
            self.entry_principal.insert(0, self.item_values[5])
            self.entry_interest.insert(0, self.item_values[6])

        # Action button
        btn_action = ttk.Button(self, text="Save", command=self.save_entry)
        btn_action.grid(row=6, column=0, columnspan=2, pady=10)

    def save_entry(self):
        fn = self.entry_fn.get().strip()
        ln = self.entry_ln.get().strip()
        project_name = self.entry_project.get().strip()  # NEW
        origin_date_str = self.entry_origin.get().strip()
        principal_str = self.entry_principal.get().strip()
        interest_str = self.entry_interest.get().strip()

        if not (fn and ln and project_name and origin_date_str and principal_str and interest_str):
            messagebox.showerror("Error", "All fields are required.")
            return

        maturity_date_str = calculate_maturity_date(origin_date_str)
        principal_plus_interest = calculate_principal_plus_interest(principal_str, interest_str)

        if self.mode == "add":
            # Use pd.concat instead of df.append (pandas 2.0+)
            new_row = {
                COLUMN_FIRST_NAME: fn,
                COLUMN_LAST_NAME: ln,
                COLUMN_PROJECT_NAME: project_name,  # NEW
                COLUMN_ORIGIN_DATE: origin_date_str,
                COLUMN_MATURITY_DATE: maturity_date_str,
                COLUMN_PRINCIPAL: float(principal_str),
                COLUMN_INTEREST_RATE: float(interest_str),
                COLUMN_PRINCIPAL_PLUS_INTEREST: principal_plus_interest
            }
            self.app.df = pd.concat([self.app.df, pd.DataFrame([new_row])], ignore_index=True)
        else:
            # Editing an existing row
            item_values = self.item_values
            cond = (
                (self.app.df[COLUMN_FIRST_NAME] == item_values[0]) & 
                (self.app.df[COLUMN_LAST_NAME] == item_values[1]) &
                (self.app.df[COLUMN_PROJECT_NAME] == item_values[2]) &  # NEW
                (self.app.df[COLUMN_ORIGIN_DATE].astype(str) == str(item_values[3])) &
                (self.app.df[COLUMN_MATURITY_DATE].astype(str) == str(item_values[4])) &
                (self.app.df[COLUMN_PRINCIPAL].astype(str) == str(item_values[5])) &
                (self.app.df[COLUMN_INTEREST_RATE].astype(str) == str(item_values[6])) &
                (self.app.df[COLUMN_PRINCIPAL_PLUS_INTEREST].astype(str) == str(item_values[7]))
            )
            idx = self.app.df[cond].index
            if not idx.empty:
                self.app.df.at[idx, COLUMN_FIRST_NAME] = fn
                self.app.df.at[idx, COLUMN_LAST_NAME] = ln
                self.app.df.at[idx, COLUMN_PROJECT_NAME] = project_name  # NEW
                self.app.df.at[idx, COLUMN_ORIGIN_DATE] = origin_date_str
                self.app.df.at[idx, COLUMN_MATURITY_DATE] = maturity_date_str
                self.app.df.at[idx, COLUMN_PRINCIPAL] = float(principal_str)
                self.app.df.at[idx, COLUMN_INTEREST_RATE] = float(interest_str)
                self.app.df.at[idx, COLUMN_PRINCIPAL_PLUS_INTEREST] = principal_plus_interest

        self.app.load_tree()
        self.destroy()


def main():
    root = tk.Tk()
    app = InvestmentApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
