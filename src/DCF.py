from openpyxl import Workbook
from openpyxl.utils.cell import get_column_letter
import constants
import inputs
from spreadsheet import spreadsheet
from compute import compute

class DCF:
    def __init__(self, ticker, form_type="10-K", num_past_filings=1, num_years=10):
        self.ticker = ticker
        self.num_years = num_years
        self.rates = constants.Default_rates
        self.inputs = inputs.get_inputs(ticker, form_type, num_past_filings)
    
    def compute_npv_outputs(self):
        npv = compute(self.ticker, self.inputs, self.num_years)
        return npv.compute_npv_outputs()

    def create_excel_sheet(self, filename): 
        sheet = spreadsheet(self.ticker, self.inputs, self.num_years, self.rates)
        sheet.create_excel_sheet(filename)
    
    def set_risk_free_rate(self, risk_free_rate): 
        self.rates['risk_free_rate'] = risk_free_rate

if __name__ == "__main__":
    tickers = ['aapl', 'pfe', 'msft', 'jpm', 't']
    for i in tickers: 
        dcf = DCF(i)
        npv = dcf.compute_npv_outputs()
        dcf.create_excel_sheet(i + ".xlsx")
        print(npv["Value per Share"])
        print(npv["Current Price"])