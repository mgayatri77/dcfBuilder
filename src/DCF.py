import constants
import inputs
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

class DCF:
    def __init__(self, ticker, form_type="10-K", num_past_filings=1, num_years=10):
        self.ticker = ticker
        self.num_years = num_years
        self.rates = constants.Default_rates
        self.inputs = inputs.get_inputs(ticker, form_type, num_past_filings)

    def percent_to_dec(self, number):
        return number/100.0

    def get_revenues(self): 
        revenues = [0]*(self.num_years+2)
        revenue_growth = [1]*(self.num_years+2)
        
        revenue_growth[0] = 0
        revenues[0] = int(self.inputs["Revenue"][0])

        for i in range(1, len(revenues)): 
            revenue_multiple = 1
            if(i == 1): 
                revenue_multiple += self.percent_to_dec(self.rates["growth_rate_1_year"])
            elif(i >= 2 and i <= 5):
                revenue_multiple += self.percent_to_dec(self.rates["growth_rate_2_to_5_year"])
            elif(i >= len(revenues)-1): 
                revenue_multiple = self.percent_to_dec(self.rates['risk_free_rate'])
            else: 
                fraction = self.percent_to_dec(self.rates["growth_rate_2_to_5_year"] - self.rates["risk_free_rate"])
                growth_rate = fraction*(self.num_years-i)/(self.num_years - 5)
                revenue_multiple += growth_rate + self.percent_to_dec(self.rates["risk_free_rate"]/100.0)
            revenues[i] = revenues[i-1]*revenue_multiple
            revenue_growth[i] = revenue_multiple-1
        return revenues, revenue_growth

    def get_operating_income(self, revenues):
        costs_and_expenses = [0]*(self.num_years+2)
        operating_incomes = [0]*(self.num_years+2)
        operating_margins = [0]*(self.num_years+2)

        operating_incomes[0] = int(self.inputs["Operating Income"][0])
        costs_and_expenses[0] = int(self.inputs["COGS"][0])
        operating_margins[0] = 100*operating_incomes[0]/revenues[0]

        c_and_e_theoretical = int(self.inputs["Revenue"][0]) - operating_incomes[0]
        if(c_and_e_theoretical != int(self.inputs['COGS'][0])):
            self.inputs["COGS"][0] = c_and_e_theoretical

        for i in range(1, len(costs_and_expenses)):
            if(i == 1): 
                operating_margins[i] = self.rates["operating_margin_1_year"]
            elif(i == (len(costs_and_expenses)-1)): 
                operating_margins[i] = self.rates["target_pretax_operating_margin"]
            elif(i > 1 and i <= 5): 
                target = self.rates["target_pretax_operating_margin"]
                year_1 = self.rates["operating_margin_1_year"]
                margin =  target - ((target-year_1)/self.num_years)*(self.num_years-i)
                operating_margins[i] = margin
            else: 
                target = self.rates["target_pretax_operating_margin"]
                initial = operating_margins[0]
                margin =  target - ((target-initial)/self.num_years)*(self.num_years-i)
                operating_margins[i] = margin
            operating_incomes[i] = self.percent_to_dec(operating_margins[i])*revenues[i]
            costs_and_expenses[i] = revenues[i] - operating_incomes[i]

        return costs_and_expenses, operating_incomes, operating_margins

    def get_taxes(self, operating_incomes):
        tax_rate = [0]*(self.num_years+2)
        taxes = [1]*(self.num_years+2)

        taxes[0] = self.inputs["Taxes"][0]
        tax_rate[0] = (taxes[0])/operating_incomes[0]
        marginal_tax_rate = self.percent_to_dec(self.rates["marginal_tax_rate"])

        if(tax_rate[0] < 0): 
            tax_rate[0] = 0
        elif(tax_rate[0] > 1): 
            tax_rate[0] = marginal_tax_rate
        
        for i in range(1, len(tax_rate)):
            if(i <= 5):
                tax_rate[i] = tax_rate[i-1] 
            elif(i == len(tax_rate)-1): 
                tax_rate[i] = marginal_tax_rate
            else: 
                difference = abs(tax_rate[0] - marginal_tax_rate)
                tax_rate[i] = tax_rate[i-1] + difference/5
            taxes[i] = tax_rate[i]*operating_incomes[i]
        return taxes, tax_rate

    def get_net_income(self, tax_rate, operating_incomes): 
        net_incomes = [0]*(self.num_years+2)
        for i in range(len(operating_incomes)):
            if(operating_incomes[i] > 0):  
                net_incomes[i] = operating_incomes[i]*(1-(tax_rate[i]))
            else: 
                net_incomes[i] = operating_incomes[i]
        return net_incomes
    
    def get_s_to_c_ratio(self): 
        s_to_c_ratio = [0]*(self.num_years+2)
        for i in range(1, len(s_to_c_ratio)): 
            if(i == 1): 
                s_to_c_ratio[i] = self.rates["sales_to_capital_ratio_1_year"]
            elif(i > 1 and i < 6): 
                s_to_c_ratio[i] = self.rates["sales_to_capital_ratio_2_to_5_year"]
            else: 
                s_to_c_ratio[i] = self.rates["sales_to_capital_ratio_6_to_10_year"]
        return s_to_c_ratio
    
    def get_reinvestment(self, revenues, net_incomes, s_to_c_ratio): 
        reinvestment = [0]*(self.num_years+2)
        
        for i in range(1, len(reinvestment)): 
            if(i == 1): 
                if(revenues[i] > revenues[i-1]):
                    reinvestment[i] = (revenues[i] - revenues[i-1])/s_to_c_ratio[i]
            else: 
                reinvestment[i] = (revenues[i] - revenues[i-1])/s_to_c_ratio[i]
        return reinvestment
    
    def get_free_cash_flow(self, net_incomes, reinvestment): 
        fcf = [0]*(self.num_years+2)
        for i in range(1, len(fcf)): 
            fcf[i] = net_incomes[i] - reinvestment[i]
        return fcf

    def get_cost_of_capital(self):
        cost_of_capital = [self.rates["initial_cost_of_capital"]]*(self.num_years+2)
        cost_of_capital[0] = 0
        return cost_of_capital

    def get_cumulative_discount_factor(self, cost_of_capital): 
        discount_factor = [0]*(self.num_years+2)
        discount_factor[0] = 1
        for i in range(1, len(discount_factor)): 
            discount_factor[i] = discount_factor[i-1]/(1+self.percent_to_dec(cost_of_capital[i]))
        return discount_factor

    def get_present_value(self, fcf, discount_factor):
        pv = [0]*(self.num_years+2)
        for i in range(1, len(pv)-1): 
            pv[i] = fcf[i]*discount_factor[i]
        return pv

    def get_packaged_inputs(self): 
        revenues, revenue_growth = self.get_revenues()
        costs_and_expenses, operating_incomes, operating_margins = self.get_operating_income(revenues)
        taxes, tax_rate = self.get_taxes(operating_incomes)
        net_incomes = self.get_net_income(tax_rate, operating_incomes)
        s_to_c_ratio = self.get_s_to_c_ratio()
        reinvestment = self.get_reinvestment(revenues, net_incomes, s_to_c_ratio)
        packaged_inputs = {"Revenue" : revenues, "Revenue Growth Rate" : revenue_growth, "Costs and Expenses" : costs_and_expenses, 
        "Operating Income" : operating_incomes, "Margin": operating_margins, "Taxes" : taxes, "Tax Rate" : tax_rate,
        "Net Income": net_incomes, "Reinvestment" : reinvestment, "Sales to Capital Ratio" : s_to_c_ratio}
        return packaged_inputs

    def get_cash_flows(self):
        inputs = self.get_packaged_inputs()
        fcf = self.get_free_cash_flow(inputs["Net Income"], inputs["Reinvestment"])
        cost_of_capital = self.get_cost_of_capital()
        cumulative_discount_factor = self.get_cumulative_discount_factor(cost_of_capital)
        pv = self.get_present_value(fcf, cumulative_discount_factor)
        
        cashflows = {"FCFF" : fcf, "Cost of Capital" : cost_of_capital, 
        "Cumulative Discount Factor" : cumulative_discount_factor, "PV (FCFF)": pv}
        return cashflows
    
    def get_npv_outputs(self): 
        cashflows = self.get_cash_flows()
        tcf = cashflows["FCFF"][-1]
        terminal_coc = cashflows["Cost of Capital"][-1]
        terminal_value = tcf/self.percent_to_dec(terminal_coc-self.rates['risk_free_rate'])
        
        pv_terminal = terminal_value*cashflows["Cumulative Discount Factor"][-1]
        pv_cashflows = sum(cashflows["PV (FCFF)"])
        pv_sum = pv_terminal + pv_cashflows
        cash = self.inputs["Cash"][0] + self.inputs["Current Marketable Securities"][0] 
        cash += self.inputs["Noncurrent Marketable Securities"][0]  
        
        value_of_equity = pv_sum + cash - self.inputs["Debt"][0]
        value_per_share = value_of_equity/self.inputs["Shares"][0]

        packaged_outputs = {"Terminal Cash Flow" : tcf, "Terminal Cost of Capital" : terminal_coc, 
        "Terminal Value" : terminal_value, "PV (Terminal Value)": pv_terminal, 
        "PV (Cashflows)" : pv_cashflows, "Sum of PV": pv_sum, "Cash" : cash, 
        "Debt" : self.inputs["Debt"][0], "Value of Equity" : value_of_equity, 
        "Common Shares Outstanding": self.inputs["Shares"][0], "Value per Share": value_per_share, 
        "Current Price" : self.inputs["Price"][0]}
        return packaged_outputs
    
    def get_excel_column_names(self): 
        column_names = [0]*(self.num_years+2)
        for i in range(len(column_names)):
            if(i == 0): 
                column_names[i] = "Base Year"
            elif(i == len(column_names) - 1): 
                column_names[i] = "Terminal Year"
            else: 
                column_names[i] = "Year " + str(i)
        return column_names
    
    def create_excel_sheet(self, filename):
        assert(filename.endswith(".xlsx")), "File must have .xlsx extension"
        packaged_inputs = self.get_packaged_inputs()
        cashflows = self.get_cash_flows()
        npv_outputs = self.get_npv_outputs()
        excel_data = {**packaged_inputs, **cashflows, **npv_outputs}
        excel_data["Quantities"] = self.get_excel_column_names()

        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Valuation Output - " + self.ticker.upper()

        for row, label in constants.excel_list_row_names.items():
            current_data = excel_data[label]
            for col in range(len(current_data)+1):
                cell = get_column_letter(col+1) + str(row+1)
                if(col == 0): 
                    ws1[cell] = label
                    ws1.column_dimensions[get_column_letter(col+1)].width = constants.excel_column_widths["big"]
                else: 
                    ws1[cell] = current_data[col-1]
                    ws1.column_dimensions[get_column_letter(col+1)].width = constants.excel_column_widths["small"]
        
        for row, label in constants.excel_singular_row_names.items(): 
            label_cell = get_column_letter(1) + str(row+1) 
            value_cell = get_column_letter(2) + str(row+1)
            ws1[label_cell] = label
            ws1[value_cell] = excel_data[label]

        wb.save(filename)
        print('DCF Spreadsheet generated succesfully!')
        print('Filename: ', filename)

if __name__ == "__main__":
    dcf = DCF("orcl")
    npv = dcf.get_npv_outputs()
    dcf.create_excel_sheet("orcl.xlsx")
    print(npv["Value per Share"])
    print(npv["Current Price"])