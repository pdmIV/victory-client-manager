import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime, timedelta
import os

from tkcalendar import DateEntry  # pip install tkcalendar
from fpdf import FPDF  # pip install fpdf

# Configure CustomTkinter appearance mode
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Constants
EXCEL_FILE = "investments.xlsx"
OUTPUT_DIR = "output"

COLUMN_FIRST_NAME = "First Name"
COLUMN_LAST_NAME = "Last Name"
COLUMN_PROJECT_NAME = "Project Name"
COLUMN_ORIGIN_DATE = "Note Origin Date"
COLUMN_MONTHS_TO_MATURITY = "Months To Maturity"
COLUMN_MATURITY_DATE = "Note Maturity Date"
COLUMN_PRINCIPAL = "Principal"
COLUMN_INTEREST_RATE = "Interest Rate"
COLUMN_PRINCIPAL_PLUS_INTEREST = "Principal + Interest"

def load_investments():
    """Reads the Excel file into a pandas DataFrame, ensuring all required columns exist."""
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=[
            COLUMN_FIRST_NAME, COLUMN_LAST_NAME, COLUMN_PROJECT_NAME,
            COLUMN_ORIGIN_DATE, COLUMN_MONTHS_TO_MATURITY, COLUMN_MATURITY_DATE,
            COLUMN_PRINCIPAL, COLUMN_INTEREST_RATE, COLUMN_PRINCIPAL_PLUS_INTEREST
        ])
        df.to_excel(EXCEL_FILE, index=False)
        return df
    
    df = pd.read_excel(EXCEL_FILE)
    required_cols = [
        COLUMN_FIRST_NAME, COLUMN_LAST_NAME, COLUMN_PROJECT_NAME,
        COLUMN_ORIGIN_DATE, COLUMN_MONTHS_TO_MATURITY, COLUMN_MATURITY_DATE,
        COLUMN_PRINCIPAL, COLUMN_INTEREST_RATE, COLUMN_PRINCIPAL_PLUS_INTEREST
    ]
    for col in required_cols:
        if col not in df.columns:
            df[col] = None

    return df

def save_investments(df):
    """Writes the DataFrame back to the Excel file."""
    df.to_excel(EXCEL_FILE, index=False)

def calculate_maturity_date(origin_date_str, months):
    """
    Return maturity date string, given origin_date_str and months to maturity.
    Approximates each month as 30 days.
    """
    try:
        origin_date = datetime.strptime(origin_date_str, "%Y-%m-%d")
    except ValueError:
        origin_date = datetime.strptime(origin_date_str, "%m/%d/%Y")

    maturity_date = origin_date + timedelta(days=30 * months)
    return maturity_date.strftime("%Y-%m-%d")

def calculate_principal_plus_interest(principal, interest_rate, months):
    """
    Simple interest approximation: principal * rate * (months / 12).
    """
    try:
        principal = float(principal)
        rate = float(interest_rate)
        interest = principal * rate * (months / 12)
        total = principal + interest
        return round(total, 2)
    except:
        return None

class PDFExporter(FPDF):
    """Simple subclass of FPDF for a custom header, if desired."""
    def header(self):
        self.set_font("Arial", "B", 14)
        self.cell(0, 10, "Client Investment Letter", ln=True, align="C")
        self.ln(5)

def export_rows_to_individual_pdfs(rows, output_dir):
    """
    For each client in 'rows', create a separate letter-style PDF named 'firstname_lastname.pdf'
    placed in the 'output_dir' folder.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    for row_data in rows:
        pdf = PDFExporter()
        pdf.add_page()

        pdf.set_font("Arial", "B", 14)
        first_name = str(row_data.get(COLUMN_FIRST_NAME, "")).strip()
        last_name = str(row_data.get(COLUMN_LAST_NAME, "")).strip()

        # Construct PDF filename: "John_Doe.pdf"
        safe_first = first_name.replace(" ", "_")
        safe_last = last_name.replace(" ", "_")
        pdf_filename = f"{safe_first}_{safe_last}.pdf"
        pdf_path = os.path.join(output_dir, pdf_filename)

        # Greeting
        pdf.cell(0, 10, f"Dear {first_name} {last_name},", ln=True)
        pdf.ln(5)

        # Body
        pdf.set_font("Arial", size=12)
        project_name = str(row_data.get(COLUMN_PROJECT_NAME, ""))
        origin_date = str(row_data.get(COLUMN_ORIGIN_DATE, ""))
        months_to_maturity = str(row_data.get(COLUMN_MONTHS_TO_MATURITY, ""))
        maturity_date = str(row_data.get(COLUMN_MATURITY_DATE, ""))
        principal = str(row_data.get(COLUMN_PRINCIPAL, ""))
        rate = str(row_data.get(COLUMN_INTEREST_RATE, ""))
        total_payoff = str(row_data.get(COLUMN_PRINCIPAL_PLUS_INTEREST, ""))

        letter_body = (
            f"Thank you for your investment in {project_name}.\n\n"
            f"Your note originated on {origin_date} for a term of {months_to_maturity} months. "
            f"The maturity date is {maturity_date}.\n\n"
            f"Principal Amount: ${principal}\n"
            f"Interest Rate: {rate}\n"
            f"Total Expected Payout at Maturity: ${total_payoff}\n\n"
            "We appreciate your continued support. If you have any questions regarding your "
            "investment, please feel free to contact us.\n\n"
            "Sincerely,\n"
            "Your Investment Firm"
        )
        pdf.multi_cell(0, 10, letter_body, align="L")

        pdf.ln(10)
        pdf.cell(0, 10, "---------------------------------", ln=True, align="C")
        pdf.cell(0, 10, "Authorized Signature", ln=True, align="C")

        pdf.output(pdf_path)
        print(f"PDF Exported: {pdf_path}")

class InvestmentApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Client Investments Manager")

        # Load DataFrame
        self.df = load_investments()

        # Main frame
        self.main_frame = ctk.CTkFrame(self.master, corner_radius=10)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        style = ttk.Style(self.main_frame)
        style.configure("Treeview.Heading", font=("Helvetica", 12, "bold"))
        style.configure("Treeview", font=("Helvetica", 11, "bold"))

        self.columns = [
            COLUMN_FIRST_NAME, COLUMN_LAST_NAME, COLUMN_PROJECT_NAME,
            COLUMN_ORIGIN_DATE, COLUMN_MONTHS_TO_MATURITY, COLUMN_MATURITY_DATE,
            COLUMN_PRINCIPAL, COLUMN_INTEREST_RATE, COLUMN_PRINCIPAL_PLUS_INTEREST
        ]

        self.tree = ttk.Treeview(
            self.main_frame, 
            columns=self.columns, 
            show="headings", 
            height=15, 
            style="Treeview"
        )
        self.tree.tag_configure("alert", background="red", foreground="white")

        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, anchor=tk.CENTER)

        self.tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.load_tree()

        # Frame for Buttons, Search, Export
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

        # --- NEW: Export Option and Button ---
        self.export_options = ["Export Selected Client", "Export Matured Clients", "Export All Clients"]
        self.export_var = ctk.StringVar(value=self.export_options[0])
        self.export_menu = ctk.CTkOptionMenu(self.btn_frame, values=self.export_options, variable=self.export_var)
        self.export_menu.pack(side="left", padx=(20,5))

        self.export_btn = ctk.CTkButton(self.btn_frame, text="Export to PDF", command=self.on_export_button)
        self.export_btn.pack(side="left", padx=(0,5))


    def load_tree(self, df=None):
        """
        Clear and repopulate Treeview. 
        Highlights rows in red if maturity <= 7 days away or if it is already past due.
        """
        for row in self.tree.get_children():
            self.tree.delete(row)

        if df is None:
            df = self.df

        for _, row_data in df.iterrows():
            row_tags = ()
            maturity_str = row_data.get(COLUMN_MATURITY_DATE, "")
            try:
                maturity_dt = datetime.strptime(maturity_str, "%Y-%m-%d")
                days_to_maturity = (maturity_dt - datetime.now()).days
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
                    row_data.get(COLUMN_MONTHS_TO_MATURITY, ""),
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
                (self.df[COLUMN_MONTHS_TO_MATURITY].astype(str) == str(item_values[4])) &
                (self.df[COLUMN_MATURITY_DATE].astype(str) == str(item_values[5])) &
                (self.df[COLUMN_PRINCIPAL].astype(str) == str(item_values[6])) &
                (self.df[COLUMN_INTEREST_RATE].astype(str) == str(item_values[7])) &
                (self.df[COLUMN_PRINCIPAL_PLUS_INTEREST].astype(str) == str(item_values[8]))
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

    def on_export_button(self):
        """
        Handles the three export scenarios:
          - Export Selected Client (one row)
          - Export Matured Clients (rollover logic)
          - Export All Clients
        Then calls 'export_rows_to_individual_pdfs(...)' to produce letter PDFs named firstname_lastname.pdf.
        """
        export_choice = self.export_var.get()
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)

        if export_choice == "Export Selected Client":
            selected_item = self.tree.selection()
            if not selected_item:
                messagebox.showwarning("Warning", "Select a client row to export.")
                return
            item_values = self.tree.item(selected_item[0], "values")
            row_dict = self.item_values_to_dict(item_values)
            rows_to_export = [row_dict]
            export_rows_to_individual_pdfs(rows_to_export, OUTPUT_DIR)
            messagebox.showinfo("Exported", f"PDF exported for selected client.")
        
        elif export_choice == "Export Matured Clients":
            matured_rows = []
            matured_condition = []
            now = datetime.now()
            for idx, row_data in self.df.iterrows():
                maturity_str = str(row_data.get(COLUMN_MATURITY_DATE, ""))
                try:
                    maturity_dt = datetime.strptime(maturity_str, "%Y-%m-%d")
                    if maturity_dt < now:
                        matured_rows.append(row_data)
                        matured_condition.append(idx)
                except:
                    pass
            
            if not matured_rows:
                messagebox.showinfo("No Matured Clients", "No clients found with matured notes.")
                return

            # Rollover logic
            for idx in matured_condition:
                old_principal_plus_interest = float(self.df.at[idx, COLUMN_PRINCIPAL_PLUS_INTEREST])
                old_interest_rate = float(self.df.at[idx, COLUMN_INTEREST_RATE])
                old_months = float(self.df.at[idx, COLUMN_MONTHS_TO_MATURITY])
                old_maturity_date_str = self.df.at[idx, COLUMN_MATURITY_DATE]

                new_origin_date_str = old_maturity_date_str
                new_principal = old_principal_plus_interest

                new_maturity_date_str = calculate_maturity_date(new_origin_date_str, old_months)
                new_p_plus_i = calculate_principal_plus_interest(new_principal, old_interest_rate, old_months)

                self.df.at[idx, COLUMN_ORIGIN_DATE] = new_origin_date_str
                self.df.at[idx, COLUMN_PRINCIPAL] = new_principal
                self.df.at[idx, COLUMN_MATURITY_DATE] = new_maturity_date_str
                self.df.at[idx, COLUMN_PRINCIPAL_PLUS_INTEREST] = new_p_plus_i

            save_investments(self.df)
            self.load_tree()

            rows_dict_list = [rd.to_dict() for rd in matured_rows]
            export_rows_to_individual_pdfs(rows_dict_list, OUTPUT_DIR)
            messagebox.showinfo("Exported", f"Matured clients exported individually and rolled over.")
        
        else:  # Export All Clients
            all_rows = [row.to_dict() for _, row in self.df.iterrows()]
            export_rows_to_individual_pdfs(all_rows, OUTPUT_DIR)
            messagebox.showinfo("Exported", f"All clients exported as individual PDFs.")

    def item_values_to_dict(self, item_values):
        """
        Convert the tuple of item values from the TreeView back into a dictionary
        matching the DataFrame columns.
        """
        # item_values = (First, Last, Project, Origin, Months, Maturity, Principal, Rate, P+I)
        row_dict = {
            COLUMN_FIRST_NAME: item_values[0],
            COLUMN_LAST_NAME: item_values[1],
            COLUMN_PROJECT_NAME: item_values[2],
            COLUMN_ORIGIN_DATE: item_values[3],
            COLUMN_MONTHS_TO_MATURITY: item_values[4],
            COLUMN_MATURITY_DATE: item_values[5],
            COLUMN_PRINCIPAL: item_values[6],
            COLUMN_INTEREST_RATE: item_values[7],
            COLUMN_PRINCIPAL_PLUS_INTEREST: item_values[8]
        }
        return row_dict

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
        lbl_month_slider = ctk.CTkLabel(self.frame, text="Months To Maturity:")
        lbl_principal = ctk.CTkLabel(self.frame, text="Principal:")
        lbl_interest = ctk.CTkLabel(self.frame, text="Interest Rate (decimal):")

        self.entry_fn = ctk.CTkEntry(self.frame, width=200)
        self.entry_ln = ctk.CTkEntry(self.frame, width=200)
        self.entry_project = ctk.CTkEntry(self.frame, width=200)
        self.calendar_origin = DateEntry(self.frame, date_pattern="yyyy-mm-dd", selectmode='day', width=18)

        self.slider_months = ctk.CTkSlider(self.frame, from_=1, to=60, number_of_steps=59, width=200)
        self.label_month_value = ctk.CTkLabel(self.frame, text="1")

        self.entry_principal = ctk.CTkEntry(self.frame, width=200)
        self.entry_interest = ctk.CTkEntry(self.frame, width=200)

        lbl_fn.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        lbl_ln.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        lbl_project.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        lbl_origin.grid(row=3, column=0, padx=5, pady=5, sticky="e")
        lbl_month_slider.grid(row=4, column=0, padx=5, pady=5, sticky="e")
        lbl_principal.grid(row=5, column=0, padx=5, pady=5, sticky="e")
        lbl_interest.grid(row=6, column=0, padx=5, pady=5, sticky="e")

        self.entry_fn.grid(row=0, column=1, padx=5, pady=5)
        self.entry_ln.grid(row=1, column=1, padx=5, pady=5)
        self.entry_project.grid(row=2, column=1, padx=5, pady=5)
        self.calendar_origin.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        self.slider_months.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        self.label_month_value.grid(row=4, column=2, padx=5, pady=5, sticky="w")

        self.entry_principal.grid(row=5, column=1, padx=5, pady=5)
        self.entry_interest.grid(row=6, column=1, padx=5, pady=5)

        if self.mode == "edit" and self.item_values:
            # item_values: (FN, LN, Project, OriginDate, Months, MaturityDate, Principal, Rate, P+I)
            self.entry_fn.insert(0, self.item_values[0])
            self.entry_ln.insert(0, self.item_values[1])
            self.entry_project.insert(0, self.item_values[2])
            try:
                origin_dt = datetime.strptime(self.item_values[3], "%Y-%m-%d")
                self.calendar_origin.set_date(origin_dt)
            except:
                pass
            if self.item_values[4]:
                try:
                    months_val = float(self.item_values[4])
                    if months_val < 1: months_val = 1
                    if months_val > 60: months_val = 60
                    self.slider_months.set(months_val)
                    self.label_month_value.configure(text=str(int(months_val)))
                except:
                    pass
            self.entry_principal.insert(0, self.item_values[6])
            self.entry_interest.insert(0, self.item_values[7])
        else:
            # default slider to 9 months
            self.slider_months.set(9)
            self.label_month_value.configure(text="9")

        btn_action = ctk.CTkButton(self.frame, text="Save", command=self.save_entry)
        btn_action.grid(row=7, column=0, columnspan=3, pady=10)

        self.slider_months.bind("<B1-Motion>", self.update_month_label)
        self.slider_months.bind("<ButtonRelease-1>", self.update_month_label)

    def update_month_label(self, event):
        current_val = int(self.slider_months.get())
        self.label_month_value.configure(text=str(current_val))

    def save_entry(self):
        fn = self.entry_fn.get().strip()
        ln = self.entry_ln.get().strip()
        project_name = self.entry_project.get().strip()
        origin_date_obj = self.calendar_origin.get_date()
        origin_date_str = origin_date_obj.strftime("%Y-%m-%d")

        months_val = int(self.slider_months.get())
        principal_str = self.entry_principal.get().strip()
        interest_str = self.entry_interest.get().strip()

        if not (fn and ln and project_name and origin_date_str and principal_str and interest_str):
            messagebox.showerror("Error", "All fields are required.")
            return

        maturity_date_str = calculate_maturity_date(origin_date_str, months_val)
        principal_plus_interest = calculate_principal_plus_interest(principal_str, interest_str, months_val)

        if self.mode == "add":
            new_row = {
                COLUMN_FIRST_NAME: fn,
                COLUMN_LAST_NAME: ln,
                COLUMN_PROJECT_NAME: project_name,
                COLUMN_ORIGIN_DATE: origin_date_str,
                COLUMN_MONTHS_TO_MATURITY: months_val,
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
                (self.app.df[COLUMN_MONTHS_TO_MATURITY].astype(str) == str(item_values[4])) &
                (self.app.df[COLUMN_MATURITY_DATE].astype(str) == str(item_values[5])) &
                (self.app.df[COLUMN_PRINCIPAL].astype(str) == str(item_values[6])) &
                (self.app.df[COLUMN_INTEREST_RATE].astype(str) == str(item_values[7])) &
                (self.app.df[COLUMN_PRINCIPAL_PLUS_INTEREST].astype(str) == str(item_values[8]))
            )
            idx = self.app.df[cond].index
            if not idx.empty:
                self.app.df.at[idx, COLUMN_FIRST_NAME] = fn
                self.app.df.at[idx, COLUMN_LAST_NAME] = ln
                self.app.df.at[idx, COLUMN_PROJECT_NAME] = project_name
                self.app.df.at[idx, COLUMN_ORIGIN_DATE] = origin_date_str
                self.app.df.at[idx, COLUMN_MONTHS_TO_MATURITY] = months_val
                self.app.df.at[idx, COLUMN_MATURITY_DATE] = maturity_date_str
                self.app.df.at[idx, COLUMN_PRINCIPAL] = float(principal_str)
                self.app.df.at[idx, COLUMN_INTEREST_RATE] = float(interest_str)
                self.app.df.at[idx, COLUMN_PRINCIPAL_PLUS_INTEREST] = principal_plus_interest

        self.app.load_tree()
        self.destroy()

def main():
    root = ctk.CTk()
    root.geometry("1500x700")
    app = InvestmentApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()