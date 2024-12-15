import customtkinter as ctk

from models import InvestmentModel
from controllers import InvestmentController
from views import InvestmentApp

def main():
    model = InvestmentModel()
    controller = InvestmentController(model)  # create the Model + Controller
    root = ctk.CTk()
    root.geometry("1665x700")
    app = InvestmentApp(root, controller)  # pass controller to the view
    root.mainloop()

if __name__ == "__main__":
    main()
