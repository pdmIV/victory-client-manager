import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from datetime import datetime

# Import column constants and output dir if needed
from models import (
    COLUMN_FIRST_NAME, COLUMN_LAST_NAME, COLUMN_PROJECT_NAME, COLUMN_ORIGIN_DATE,
    COLUMN_MONTHS_TO_MATURITY, COLUMN_MATURITY_DATE, COLUMN_PRINCIPAL,
    COLUMN_INTEREST_RATE, COLUMN_PRINCIPAL_PLUS_INTEREST
)
from controllers import InvestmentController


class InvestmentApp(ctk.CTkFrame):
    """
    The main GUI for the Investment Manager. 
    Relies on an InvestmentController to manage the data logic.
    """
    def __init__(self, master, controller: InvestmentController):
        super().__init__(master)
        self.master = master
        self.controller = controller  # The new controller instance
        self.pack(fill="both", expand=True)
        self.setup_ui()

    def setup_ui(self):
        self.master.title("Victory Client Manager")

        # Main frame
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
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

        # Buttons + Search Frame
        self.btn_frame = ctk.CTkFrame(self, corner_radius=10)
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

        # Export Option + Button
        self.export_options = ["Export Selected Client", "Export Matured Clients", "Export All Clients"]
        self.export_var = tk.StringVar(value=self.export_options[0])
        self.export_menu = ctk.CTkOptionMenu(self.btn_frame, values=self.export_options, variable=self.export_var)
        self.export_menu.pack(side="left", padx=(20,5))

        self.export_btn = ctk.CTkButton(self.btn_frame, text="Export to PDF", command=self.on_export_button)
        self.export_btn.pack(side="left", padx=(0,5))

        # Finally, populate the tree
        self.load_tree()

    def load_tree(self, df=None):
        """
        Clears and repopulates the Treeview.
        If df is None, uses self.controller.model.df.
        Highlights rows in red if maturity <= 7 days or past due.
        """
        for row in self.tree.get_children():
            self.tree.delete(row)

        if df is None:
            df = self.controller.model.df

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
        EntryWindow(self.master, self.controller, self, mode="add")

    def edit_entry(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Select an entry to edit.")
            return
        item_values = self.tree.item(selected_item[0], "values")
        EntryWindow(self.master, self.controller, self, mode="edit", item_values=item_values)

    def del_entry(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Select an entry to delete.")
            return
        
        confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete?")
        if confirm:
            item_values = self.tree.item(selected_item[0], "values")
            self.controller.delete_investment(item_values)
            self.tree.delete(selected_item[0])

    def save_changes(self):
        self.controller.save_to_excel()
        messagebox.showinfo("Saved", "Changes saved to Excel file successfully.")

    def search_by_project_name(self):
        query = self.search_entry.get().strip()
        if not query:
            messagebox.showwarning("Warning", "Enter a project name to search.")
            return
        filtered_df = self.controller.search_by_project_name(query)
        self.load_tree(filtered_df)

    def reset_view(self):
        self.search_entry.delete(0, tk.END)
        self.load_tree()  # reload from the full model

    def on_export_button(self):
        export_choice = self.export_var.get()

        if export_choice == "Export Selected Client":
            selected_item = self.tree.selection()
            if not selected_item:
                messagebox.showwarning("Warning", "Select a client row to export.")
                return
            item_values = self.tree.item(selected_item[0], "values")
            row_dict = self.item_values_to_dict(item_values)
            self.controller.export_selected_client([row_dict])
            messagebox.showinfo("Exported", f"PDF exported for selected client.")

        elif export_choice == "Export Matured Clients":
            matured_rows = self.controller.export_matured_clients()
            if not matured_rows:
                messagebox.showinfo("No Matured Clients", "No clients found with matured notes.")
                return
            # We already rolled them over in the model. 
            # The returned 'matured_rows' is the old data before rollover, which we can export.
            rows_dict_list = [rd.to_dict() for rd in matured_rows]
            messagebox.showinfo("Exported", f"Matured clients exported individually and rolled over.")
            self.load_tree()  # refresh with updated data

        else:  # Export All Clients
            all_rows = self.controller.export_all_clients()
            self.controller.export_selected_client(all_rows)  # re-use the same method
            messagebox.showinfo("Exported", f"All clients exported as individual PDFs.")

    def item_values_to_dict(self, item_values):
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
    """
    The window for adding/editing a single client entry.
    Uses the controller to manipulate the model.
    """
    def __init__(self, parent, controller: InvestmentController, main_app, mode="add", item_values=None):
        super().__init__(parent)
        self.controller = controller
        self.main_app = main_app
        self.mode = mode
        self.item_values = item_values

        if self.mode == "add":
            self.title("Add New Entry")
        else:
            self.title("Edit Entry")

        self.setup_ui()

    def setup_ui(self):
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

        # Grid layout
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
            self.populate_fields()
        else:
            self.slider_months.set(9)
            self.label_month_value.configure(text="9")

        btn_action = ctk.CTkButton(self.frame, text="Save", command=self.save_entry)
        btn_action.grid(row=7, column=0, columnspan=3, pady=10)

        self.slider_months.bind("<B1-Motion>", self.update_month_label)
        self.slider_months.bind("<ButtonRelease-1>", self.update_month_label)

    def populate_fields(self):
        # item_values = (FN, LN, Project, OriginDate, Months, MaturityDate, Principal, Rate, P+I)
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

        # Let the controller handle the calculations or use the model directly
        maturity_date_str = self.controller.model.calculate_maturity_date(origin_date_str, months_val)
        principal_plus_interest = self.controller.model.calculate_principal_plus_interest(principal_str, interest_str, months_val)

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
            self.controller.add_investment(new_row)
        else:
            old_values = self.item_values
            updated_row = {
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
            self.controller.edit_investment(old_values, updated_row)

        # Refresh the main tree
        self.main_app.load_tree()
        self.destroy()
