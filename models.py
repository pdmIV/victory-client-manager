import os
import pandas as pd
from datetime import datetime, timedelta
from fpdf import FPDF

# Column constants
COLUMN_FIRST_NAME = "First Name"
COLUMN_LAST_NAME = "Last Name"
COLUMN_PROJECT_NAME = "Project Name"
COLUMN_ORIGIN_DATE = "Note Origin Date"
COLUMN_MONTHS_TO_MATURITY = "Months To Maturity"
COLUMN_MATURITY_DATE = "Note Maturity Date"
COLUMN_PRINCIPAL = "Principal"
COLUMN_INTEREST_RATE = "Interest Rate"
COLUMN_PRINCIPAL_PLUS_INTEREST = "Principal + Interest"
COLUMN_AUTO_ROLLOVER = "Auto Rollover"


EXCEL_FILE = "investments.xlsx"
OUTPUT_DIR = "output"


class InvestmentModel:
    """
    Handles the DataFrame that represents all investments.
    Responsible for loading/saving Excel and performing calculations.
    """
    def __init__(self):
        self.df = self.load_investments()

    def load_investments(self):
        """Reads Excel into a pandas DataFrame, ensuring required columns exist."""
        if not os.path.exists(EXCEL_FILE):
            df = pd.DataFrame(columns=[
                COLUMN_FIRST_NAME, COLUMN_LAST_NAME, COLUMN_PROJECT_NAME,
                COLUMN_ORIGIN_DATE, COLUMN_MONTHS_TO_MATURITY, COLUMN_MATURITY_DATE,
                COLUMN_PRINCIPAL, COLUMN_INTEREST_RATE, COLUMN_PRINCIPAL_PLUS_INTEREST,
                COLUMN_AUTO_ROLLOVER
            ])
            df.to_excel(EXCEL_FILE, index=False)
            return df
        
        df = pd.read_excel(EXCEL_FILE)
        required_cols = [
            COLUMN_FIRST_NAME, COLUMN_LAST_NAME, COLUMN_PROJECT_NAME,
            COLUMN_ORIGIN_DATE, COLUMN_MONTHS_TO_MATURITY, COLUMN_MATURITY_DATE,
            COLUMN_PRINCIPAL, COLUMN_INTEREST_RATE, COLUMN_PRINCIPAL_PLUS_INTEREST,
            COLUMN_AUTO_ROLLOVER
        ]
        for col in required_cols:
            if col not in df.columns:
                df[col] = None

        return df

    def save_investments(self):
        """Writes DataFrame to Excel file."""
        self.df.to_excel(EXCEL_FILE, index=False)

    def calculate_maturity_date(self, origin_date_str, months):
        """Approx each month as 30 days for maturity calculation."""
        try:
            origin_date = datetime.strptime(origin_date_str, "%Y-%m-%d")
        except ValueError:
            origin_date = datetime.strptime(origin_date_str, "%m/%d/%Y")
        maturity_date = origin_date + timedelta(days=30 * months)
        return maturity_date.strftime("%Y-%m-%d")

    def calculate_principal_plus_interest(self, principal, interest_rate, months):
        """Simple interest: principal * rate * (months / 12)."""
        try:
            principal = float(principal)
            rate = float(interest_rate)
            interest = principal * rate * (months / 12)
            total = principal + interest
            return round(total, 2)
        except:
            return None


class PDFExporter(FPDF):
    """Simple FPDF subclass for consistent styling."""
    def header(self):
        self.set_font("Arial", "B", 14)
        self.cell(0, 10, "Client Investment Letter", ln=True, align="C")
        self.ln(5)


def export_rows_to_individual_pdfs(rows, output_dir, model: InvestmentModel):
    """
    For each client row, create a letter-style PDF named 'firstname_lastname.pdf'
    in 'output_dir'. 'model' can be used if we need logic from the model.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    for row_data in rows:
        pdf = PDFExporter()
        pdf.add_page()

        pdf.set_font("Arial", "B", 14)
        first_name = str(row_data.get(COLUMN_FIRST_NAME, "")).strip()
        last_name = str(row_data.get(COLUMN_LAST_NAME, "")).strip()

        safe_first = first_name.replace(" ", "_")
        safe_last = last_name.replace(" ", "_")
        pdf_filename = f"{safe_first}_{safe_last}.pdf"
        pdf_path = os.path.join(output_dir, pdf_filename)

        pdf.cell(0, 10, f"Dear {first_name} {last_name},", ln=True)
        pdf.ln(5)
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
