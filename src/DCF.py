from openpyxl import Workbook
from openpyxl.utils.cell import get_column_letter
import constants
import inputs
from spreadsheet import spreadsheet
from compute import compute

class DCF:
    def __init__(self, ticker, form_type="10-K", num_past_filings=1, num_years=10):
        self.ticker = ticker
        self.num_years = 10
        self.rates = constants.Default_rates
        # only supports 10-K with 1 past filing right now
        self.inputs = inputs.get_inputs(ticker, "10-K", 1)
    
    def compute_npv_outputs(self):
        npv = compute(self.ticker, self.inputs, self.num_years)
        return npv.compute_npv_outputs()

    def create_excel_sheet(self, filename): 
        sheet = spreadsheet(self.ticker, self.inputs, self.num_years, self.rates)
        sheet.create_excel_sheet(filename)
    
    def set_rates(self, **new_rates):
        for key, value in new_rates.items():  
            if(key in self.rates): 
                self.rates[key] = value
            else:
                print("Could not find rate parameter: %s" % key)
                print("Check constants.py for the correct name")

    def reset_rates(self): 
        self.rates = constants.Default_rates

if __name__ == "__main__":
    dcf = DCF("AAPL")
    npv = dcf.compute_npv_outputs()
    dcf.create_excel_sheet("Apple.xlsx")
    print("%.2f" % npv["Value per Share"])
    print("%.2f" % npv["Current Price"])