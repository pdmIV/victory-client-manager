import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime, timedelta
import os

from tkcalendar import DateEntry  # pip install tkcalendar

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

EXCEL_FILE = "investments.xlsx"

COLUMN_FIRST_NAME = "First Name"
COLUMN_LAST_NAME = "Last Name"
COLUMN_PROJECT_NAME = "Project Name"
COLUMN_ORIGIN_DATE = "Note Origin Date"
COLUMN_MATURITY_DATE = "Note Maturity Date"
COLUMN_PRINCIPAL = "Principal"
COLUMN_INTEREST_RATE = "Interest Rate"
COLUMN_PRINCIPAL_PLUS_INTEREST = "Principal + Interest"


def load_investments():
    """Reads the Excel file into a pandas DataFrame, ensuring all columns exist."""
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=[
            COLUMN_FIRST_NAME, COLUMN_LAST_NAME, COLUMN_PROJECT_NAME,
            COLUMN_ORIGIN_DATE, COLUMN_MATURITY_DATE, COLUMN_PRINCIPAL,
            COLUMN_INTEREST_RATE, COLUMN_PRINCIPAL_PLUS_INTEREST
        ])
        df.to_excel(EXCEL_FILE, index=False)
        return df
    
    df = pd.read_excel(EXCEL_FILE)
    required_cols = [
        COLUMN_FIRST_NAME, COLUMN_LAST_NAME, COLUMN_PROJECT_NAME,
        COLUMN_ORIGIN_DATE, COLUMN_MATURITY_DATE, COLUMN_PRINCIPAL,
        COLUMN_INTEREST_RATE, COLUMN_PRINCIPAL_PLUS_INTEREST
    ]
    for col in required_cols:
        if col not in df.columns:
            df[col] = None

    return df


def save_investments(df):
    """Writes the DataFrame to the Excel file."""
    df.to_excel(EXCEL_FILE, index=False)


def calculate_maturity_date(origin_date_str):
    """Return ~9-month later maturity date string, given an origin date string."""
    try:
        origin_date = datetime.strptime(origin_date_str, "%Y-%m-%d")
    except ValueError:
        # Attempt a different format
        origin_date = datetime.strptime(origin_date_str, "%m/%d/%Y")
    maturity_date = origin_date + timedelta(days=9*30)  # approximate 9 months as 270 days
    return maturity_date.strftime("%Y-%m-%d")


def calculate_principal_plus_interest(principal, interest_rate):
    """Approximate 9-month simple interest on principal at given annual interest rate."""
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

        # Load DataFrame from Excel
        self.df = load_investments()

        # Main frame
        self.main_frame = ctk.CTkFrame(self.master, corner_radius=10)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Treeview style overrides for larger/bold text
        style = ttk.Style(self.main_frame)
        style.configure("Treeview.Heading", font=("Helvetica", 12, "bold"))
        style.configure("Treeview", font=("Helvetica", 11, "bold"))

        self.columns = [
            COLUMN_FIRST_NAME, COLUMN_LAST_NAME, COLUMN_PROJECT_NAME,
            COLUMN_ORIGIN_DATE, COLUMN_MATURITY_DATE,
            COLUMN_PRINCIPAL, COLUMN_INTEREST_RATE,
            COLUMN_PRINCIPAL_PLUS_INTEREST
        ]

        self.tree = ttk.Treeview(
            self.main_frame, 
            columns=self.columns, 
            show="headings", 
            height=15, 
            style="Treeview"
        )

        # Define red highlight tag
        self.tree.tag_configure("alert", background="red", foreground="white")

        self.tree.heading(COLUMN_FIRST_NAME, text="First Name")
        self.tree.heading(COLUMN_LAST_NAME, text="Last Name")
        self.tree.heading(COLUMN_PROJECT_NAME, text="Project Name")
        self.tree.heading(COLUMN_ORIGIN_DATE, text="Origin Date")
        self.tree.heading(COLUMN_MATURITY_DATE, text="Maturity Date")
        self.tree.heading(COLUMN_PRINCIPAL, text="Principal")
        self.tree.heading(COLUMN_INTEREST_RATE, text="Interest Rate")
        self.tree.heading(COLUMN_PRINCIPAL_PLUS_INTEREST, text="Principal + Interest")

        for col in self.columns:
            self.tree.column(col, width=130, anchor=tk.CENTER)

        self.tree.pack(side="left", fill="both", expand=True)

        # Scrollbar
        scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # Populate the tree
        self.load_tree()

        # Buttons + Search Frame
        self.btn_frame = ctk.CTkFrame(self.master, corner_radius=10)
        self.btn_frame.pack(fill="x", padx=10, pady=5)

        self.add_btn = ctk.CTkButton(self.btn_frame, text="Add Entry", command=self.add_entry)
        self.add_btn.pack(side="left", padx=(0, 5))

        self.edit_btn = ctk.CTkButton(self.btn_frame, text="Edit Entry", command=self.edit_entry)
        self.edit_btn.pack(side="left", padx=(0, 5))

        self.del_btn = ctk.CTkButton(self.btn_frame, text="Delete Entry", fg_color="red", command=self.del_entry)
        self.del_btn.pack(side="left", padx=(0, 5))

        self.save_btn = ctk.CTkButton(self.btn_frame, text="Save to Excel", command=self.save_changes)
        self.save_btn.pack(side="left", padx=(0, 15))

        self.search_label = ctk.CTkLabel(self.btn_frame, text="Search Project Name:")
        self.search_label.pack(side="left", padx=(20, 5))

        self.search_entry = ctk.CTkEntry(self.btn_frame, width=140)
        self.search_entry.pack(side="left", padx=(0, 5))

        self.search_btn = ctk.CTkButton(self.btn_frame, text="Search", command=self.search_by_project_name)
        self.search_btn.pack(side="left", padx=(0, 5))

        self.reset_btn = ctk.CTkButton(self.btn_frame, text="Reset", command=self.reset_view)
        self.reset_btn.pack(side="left", padx=(0, 5))


    def load_tree(self, df=None):
        """
        Clear and repopulate the Treeview. 
        Highlights rows in red if maturity <= 7 days away or if past due.
        """
        # Clear existing items
        for row in self.tree.get_children():
            self.tree.delete(row)

        if df is None:
            df = self.df

        for _, row_data in df.iterrows():
            # Default row style
            row_tags = ()

            maturity_str = row_data.get(COLUMN_MATURITY_DATE, "")
            try:
                maturity_dt = datetime.strptime(maturity_str, "%Y-%m-%d")
                days_to_maturity = (maturity_dt - datetime.now()).days
                # If maturity is a week or less away (days_to_maturity <= 7) 
                # OR it's already past due (days_to_maturity < 0), highlight in red
                if days_to_maturity <= 7:
                    row_tags = ("alert",)
            except:
                pass

            self.tree.insert(
                "", 
                tk.END, 
                values=(
                    row_data.get(COLUMN_FIRST_NAME, ""),
                    row_data.get(COLUMN_LAST_NAME, ""),
                    row_data.get(COLUMN_PROJECT_NAME, ""),
                    row_data.get(COLUMN_ORIGIN_DATE, ""),
                    row_data.get(COLUMN_MATURITY_DATE, ""),
                    row_data.get(COLUMN_PRINCIPAL, ""),
                    row_data.get(COLUMN_INTEREST_RATE, ""),
                    row_data.get(COLUMN_PRINCIPAL_PLUS_INTEREST, "")
                ),
                tags=row_tags
            )


    def add_entry(self):
        EntryWindow(self.master, self, mode="add")

    def edit_entry(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Select an entry to edit.")
            return
        item_values = self.tree.item(selected_item[0], "values")
        EntryWindow(self.master, self, mode="edit", item_values=item_values)

    def del_entry(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Select an entry to delete.")
            return
        
        confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete?")
        if confirm:
            item_values = self.tree.item(selected_item[0], "values")
            condition = (
                (self.df[COLUMN_FIRST_NAME] == item_values[0]) & 
                (self.df[COLUMN_LAST_NAME] == item_values[1]) &
                (self.df[COLUMN_PROJECT_NAME] == item_values[2]) &
                (self.df[COLUMN_ORIGIN_DATE].astype(str) == str(item_values[3])) &
                (self.df[COLUMN_MATURITY_DATE].astype(str) == str(item_values[4])) &
                (self.df[COLUMN_PRINCIPAL].astype(str) == str(item_values[5])) &
                (self.df[COLUMN_INTEREST_RATE].astype(str) == str(item_values[6])) &
                (self.df[COLUMN_PRINCIPAL_PLUS_INTEREST].astype(str) == str(item_values[7]))
            )
            self.df.drop(self.df[condition].index, inplace=True)
            self.tree.delete(selected_item[0])

    def save_changes(self):
        save_investments(self.df)
        messagebox.showinfo("Saved", "Changes saved to Excel file successfully.")

    def search_by_project_name(self):
        query = self.search_entry.get().strip()
        if not query:
            messagebox.showwarning("Warning", "Enter a project name to search.")
            return

        filtered_df = self.df[self.df[COLUMN_PROJECT_NAME]
                              .str.contains(query, case=False, na=False)]
        self.load_tree(filtered_df)

    def reset_view(self):
        self.search_entry.delete(0, tk.END)
        self.load_tree(self.df)


class EntryWindow(ctk.CTkToplevel):
    def __init__(self, master, app, mode="add", item_values=None):
        super().__init__(master)
        self.app = app
        self.mode = mode
        self.item_values = item_values

        if self.mode == "add":
            self.title("Add New Entry")
        else:
            self.title("Edit Entry")

        self.frame = ctk.CTkFrame(self, corner_radius=10)
        self.frame.pack(padx=20, pady=20, fill="both", expand=True)

        lbl_fn = ctk.CTkLabel(self.frame, text="First Name:")
        lbl_ln = ctk.CTkLabel(self.frame, text="Last Name:")
        lbl_project = ctk.CTkLabel(self.frame, text="Project Name:")
        lbl_origin = ctk.CTkLabel(self.frame, text="Note Origin Date:")
        lbl_principal = ctk.CTkLabel(self.frame, text="Principal:")
        lbl_interest = ctk.CTkLabel(self.frame, text="Interest Rate (decimal):")

        self.entry_fn = ctk.CTkEntry(self.frame, width=200)
        self.entry_ln = ctk.CTkEntry(self.frame, width=200)
        self.entry_project = ctk.CTkEntry(self.frame, width=200)
        self.calendar_origin = DateEntry(self.frame, date_pattern="yyyy-mm-dd", selectmode='day', width=18)
        self.entry_principal = ctk.CTkEntry(self.frame, width=200)
        self.entry_interest = ctk.CTkEntry(self.frame, width=200)

        lbl_fn.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        lbl_ln.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        lbl_project.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        lbl_origin.grid(row=3, column=0, padx=5, pady=5, sticky="e")
        lbl_principal.grid(row=4, column=0, padx=5, pady=5, sticky="e")
        lbl_interest.grid(row=5, column=0, padx=5, pady=5, sticky="e")

        self.entry_fn.grid(row=0, column=1, padx=5, pady=5)
        self.entry_ln.grid(row=1, column=1, padx=5, pady=5)
        self.entry_project.grid(row=2, column=1, padx=5, pady=5)
        self.calendar_origin.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.entry_principal.grid(row=4, column=1, padx=5, pady=5)
        self.entry_interest.grid(row=5, column=1, padx=5, pady=5)

        if self.mode == "edit" and self.item_values:
            self.entry_fn.insert(0, self.item_values[0])
            self.entry_ln.insert(0, self.item_values[1])
            self.entry_project.insert(0, self.item_values[2])
            try:
                origin_dt = datetime.strptime(self.item_values[3], "%Y-%m-%d")
                self.calendar_origin.set_date(origin_dt)
            except:
                pass
            self.entry_principal.insert(0, self.item_values[5])
            self.entry_interest.insert(0, self.item_values[6])

        btn_action = ctk.CTkButton(self.frame, text="Save", command=self.save_entry)
        btn_action.grid(row=6, column=0, columnspan=2, pady=10)

    def save_entry(self):
        fn = self.entry_fn.get().strip()
        ln = self.entry_ln.get().strip()
        project_name = self.entry_project.get().strip()
        origin_date_obj = self.calendar_origin.get_date()
        origin_date_str = origin_date_obj.strftime("%Y-%m-%d")
        principal_str = self.entry_principal.get().strip()
        interest_str = self.entry_interest.get().strip()

        if not (fn and ln and project_name and origin_date_str and principal_str and interest_str):
            messagebox.showerror("Error", "All fields are required.")
            return

        maturity_date_str = calculate_maturity_date(origin_date_str)
        principal_plus_interest = calculate_principal_plus_interest(principal_str, interest_str)

        if self.mode == "add":
            new_row = {
                COLUMN_FIRST_NAME: fn,
                COLUMN_LAST_NAME: ln,
                COLUMN_PROJECT_NAME: project_name,
                COLUMN_ORIGIN_DATE: origin_date_str,
                COLUMN_MATURITY_DATE: maturity_date_str,
                COLUMN_PRINCIPAL: float(principal_str),
                COLUMN_INTEREST_RATE: float(interest_str),
                COLUMN_PRINCIPAL_PLUS_INTEREST: principal_plus_interest
            }
            self.app.df = pd.concat([self.app.df, pd.DataFrame([new_row])], ignore_index=True)
        else:
            item_values = self.item_values
            cond = (
                (self.app.df[COLUMN_FIRST_NAME] == item_values[0]) &
                (self.app.df[COLUMN_LAST_NAME] == item_values[1]) &
                (self.app.df[COLUMN_PROJECT_NAME] == item_values[2]) &
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
                self.app.df.at[idx, COLUMN_PROJECT_NAME] = project_name
                self.app.df.at[idx, COLUMN_ORIGIN_DATE] = origin_date_str
                self.app.df.at[idx, COLUMN_MATURITY_DATE] = maturity_date_str
                self.app.df.at[idx, COLUMN_PRINCIPAL] = float(principal_str)
                self.app.df.at[idx, COLUMN_INTEREST_RATE] = float(interest_str)
                self.app.df.at[idx, COLUMN_PRINCIPAL_PLUS_INTEREST] = principal_plus_interest

        self.app.load_tree()
        self.destroy()


def main():
    root = ctk.CTk()
    root.geometry("1200x700")
    app = InvestmentApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
