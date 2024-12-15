from datetime import datetime
import os
import pandas as pd

from models import (
    InvestmentModel, export_rows_to_individual_pdfs, OUTPUT_DIR,
    COLUMN_FIRST_NAME, COLUMN_LAST_NAME, COLUMN_PROJECT_NAME,
    COLUMN_ORIGIN_DATE, COLUMN_MONTHS_TO_MATURITY, COLUMN_MATURITY_DATE,
    COLUMN_PRINCIPAL, COLUMN_INTEREST_RATE, COLUMN_PRINCIPAL_PLUS_INTEREST
)

class InvestmentController:
    """
    A controller class bridging the InvestmentModel and the GUI (views).
    It handles logic for adding/editing/deleting entries, searching, and exporting.
    """

    def __init__(self, model: InvestmentModel):
        self.model = model

    def add_investment(self, new_row: dict):
        """
        Adds a new row (investment) to the DataFrame.
        new_row should have the necessary columns. (Already validated in the view)
        """
        self.model.df = pd.concat([self.model.df, pd.DataFrame([new_row])], ignore_index=True)

    def edit_investment(self, old_values: tuple, updated_row: dict):
        """
        Finds the row that matches 'old_values' in the DataFrame and updates it with 'updated_row'.
        """
        cond = (
            (self.model.df[COLUMN_FIRST_NAME] == old_values[0]) &
            (self.model.df[COLUMN_LAST_NAME] == old_values[1]) &
            (self.model.df[COLUMN_PROJECT_NAME] == old_values[2]) &
            (self.model.df[COLUMN_ORIGIN_DATE].astype(str) == str(old_values[3])) &
            (self.model.df[COLUMN_MONTHS_TO_MATURITY].astype(str) == str(old_values[4])) &
            (self.model.df[COLUMN_MATURITY_DATE].astype(str) == str(old_values[5])) &
            (self.model.df[COLUMN_PRINCIPAL].astype(str) == str(old_values[6])) &
            (self.model.df[COLUMN_INTEREST_RATE].astype(str) == str(old_values[7])) &
            (self.model.df[COLUMN_PRINCIPAL_PLUS_INTEREST].astype(str) == str(old_values[8]))
        )
        idx = self.model.df[cond].index
        if not idx.empty:
            for k, v in updated_row.items():
                self.model.df.at[idx, k] = v

    def delete_investment(self, item_values: tuple):
        cond = (
            (self.model.df[COLUMN_FIRST_NAME] == item_values[0]) &
            (self.model.df[COLUMN_LAST_NAME] == item_values[1]) &
            (self.model.df[COLUMN_PROJECT_NAME] == item_values[2]) &
            (self.model.df[COLUMN_ORIGIN_DATE].astype(str) == str(item_values[3])) &
            (self.model.df[COLUMN_MONTHS_TO_MATURITY].astype(str) == str(item_values[4])) &
            (self.model.df[COLUMN_MATURITY_DATE].astype(str) == str(item_values[5])) &
            (self.model.df[COLUMN_PRINCIPAL].astype(str) == str(item_values[6])) &
            (self.model.df[COLUMN_INTEREST_RATE].astype(str) == str(item_values[7])) &
            (self.model.df[COLUMN_PRINCIPAL_PLUS_INTEREST].astype(str) == str(item_values[8]))
        )
        self.model.df.drop(self.model.df[cond].index, inplace=True)

    def save_to_excel(self):
        self.model.save_investments()

    def search_by_project_name(self, query: str):
        """Returns a filtered DataFrame for the specified project name query."""
        return self.model.df[self.model.df[COLUMN_PROJECT_NAME].str.contains(query, case=False, na=False)]

    def export_selected_client(self, row_dict_list: list):
        """Export a single client or multiple selected rows to PDF."""
        export_rows_to_individual_pdfs(row_dict_list, OUTPUT_DIR, self.model)

    def export_matured_clients(self):
        """
        Find matured clients, roll them over, export them individually.
        Return a list of rolled-over row dictionaries for further use.
        """
        matured_rows = []
        matured_condition = []
        now = datetime.now()

        for idx, row_data in self.model.df.iterrows():
            maturity_str = str(row_data.get(COLUMN_MATURITY_DATE, ""))
            try:
                maturity_dt = datetime.strptime(maturity_str, "%Y-%m-%d")
                if maturity_dt < now:
                    matured_rows.append(row_data)
                    matured_condition.append(idx)
            except:
                pass

        if not matured_rows:
            return []

        # Rollover logic
        for idx in matured_condition:
            old_p_plus_i = float(self.model.df.at[idx, COLUMN_PRINCIPAL_PLUS_INTEREST])
            old_interest_rate = float(self.model.df.at[idx, COLUMN_INTEREST_RATE])
            old_months = float(self.model.df.at[idx, COLUMN_MONTHS_TO_MATURITY])
            old_maturity_date_str = self.model.df.at[idx, COLUMN_MATURITY_DATE]

            new_origin_date_str = old_maturity_date_str
            new_principal = old_p_plus_i
            new_maturity_date_str = self.model.calculate_maturity_date(new_origin_date_str, old_months)
            new_p_plus_i = self.model.calculate_principal_plus_interest(new_principal, old_interest_rate, old_months)

            self.model.df.at[idx, COLUMN_ORIGIN_DATE] = new_origin_date_str
            self.model.df.at[idx, COLUMN_PRINCIPAL] = new_principal
            self.model.df.at[idx, COLUMN_MATURITY_DATE] = new_maturity_date_str
            self.model.df.at[idx, COLUMN_PRINCIPAL_PLUS_INTEREST] = new_p_plus_i

        self.model.save_investments()
        return matured_rows  # list of Series

    def export_all_clients(self):
        """Return all client rows as dictionaries to export individually."""
        all_rows = [row_data.to_dict() for _, row_data in self.model.df.iterrows()]
        return all_rows
