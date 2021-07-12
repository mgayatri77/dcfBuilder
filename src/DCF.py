import constants
import inputs
import openpyxl

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
        revenues[0] = int(self.inputs["Revenue"][0])

        for i in range(1, len(revenues)): 
            revenue_multiple = 1
            if(i == 1): 
                revenue_multiple += self.percent_to_dec(self.rates["growth_rate_1_year"])
            elif(i >= 2 and i <= 5):
                revenue_multiple += self.percent_to_dec(self.rates["growth_rate_2_to_5_year"])
            elif(i == len(revenues)-1): 
                revenue_multiple = self.percent_to_dec(self.rates['risk_free_rate'])
            else: 
                fraction = self.percent_to_dec(self.rates["growth_rate_2_to_5_year"] - self.rates["risk_free_rate"])
                growth_rate = fraction*(self.num_years-i)/(self.num_years - 5)
                revenue_multiple += growth_rate + self.percent_to_dec(self.rates["risk_free_rate"]/100.0)
            revenues[i] = revenues[i-1]*revenue_multiple
        return revenues

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

    def get_tax_rate(self, operating_incomes):
        tax_rate = [0]*(self.num_years+2)
        tax_rate[0] = 100*int(self.inputs["Taxes"][0])/operating_incomes[0]

        if(tax_rate[0] < 0): 
            tax_rate[0] = 0
        elif(tax_rate[0] > 100): 
            tax_rate[0] = self.rates["marginal_tax_rate"]
        
        for i in range(1, len(tax_rate)):
            if(i <= 5):
                tax_rate[i] = tax_rate[i-1] 
            elif(i == len(tax_rate)-1): 
                tax_rate[i] = self.rates["marginal_tax_rate"]
            else: 
                difference = abs(tax_rate[0] - self.rates["marginal_tax_rate"])
                tax_rate[i] = tax_rate[i-1] + difference/5
        return tax_rate

    def get_net_income(self, tax_rate, operating_incomes): 
        net_incomes = [0]*(self.num_years+2)
        for i in range(len(operating_incomes)):
            if(operating_incomes[i] > 0):  
                net_incomes[i] = operating_incomes[i]*(1-self.percent_to_dec(tax_rate[i]))
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
        revenues = self.get_revenues()
        costs_and_expenses, operating_incomes, operating_margins = self.get_operating_income(revenues)
        tax_rate = self.get_tax_rate(operating_incomes)
        net_incomes = self.get_net_income(tax_rate, operating_incomes)
        s_to_c_ratio = self.get_s_to_c_ratio()
        reinvestment = self.get_reinvestment(revenues, net_incomes, s_to_c_ratio)
        packaged_inputs = {"Revenue" : revenues, "COGS" : costs_and_expenses, 
        "Operating income" : operating_incomes, "Margin": operating_margins, "Taxes" : tax_rate,
        "Net income": net_incomes, "Reinvestment" : reinvestment, "S to C" : s_to_c_ratio}
        return packaged_inputs

    def get_cash_flows(self):
        inputs = self.get_packaged_inputs()
        fcf = self.get_free_cash_flow(inputs["Net income"], inputs["Reinvestment"])
        cost_of_capital = self.get_cost_of_capital()
        cumulative_discount_factor = self.get_cumulative_discount_factor(cost_of_capital)
        pv = self.get_present_value(fcf, cumulative_discount_factor)
        
        cashflows = {"FCF" : fcf, "Cost of Capital" : cost_of_capital, 
        "CDF" : cumulative_discount_factor, "PV": pv}
        return cashflows
    
    def get_value_per_share(self): 
        cashflows = self.get_cash_flows()
        tcf = cashflows["FCF"][-1]
        terminal_coc = cashflows["Cost of Capital"][-1]
        terminal_value = tcf/self.percent_to_dec(terminal_coc-self.rates['risk_free_rate'])
        
        pv_terminal = terminal_value*cashflows["CDF"][-1]
        pv_cashflows = sum(cashflows["PV"])
        pv_sum = pv_terminal + pv_cashflows
        cash = self.inputs["Cash"][0] + self.inputs["Current Marketable Securities"][0] 
        cash += self.inputs["Noncurrent Marketable Securities"][0]  
        
        value_of_equity = pv_sum + cash - self.inputs["Debt"][0]
        value_per_share = value_of_equity/self.inputs["Shares"][0]

        packaged_outputs = {"Terminal cash flow" : tcf, "Terminal cost of capital" : terminal_coc, 
        "Terminal value" : terminal_value, "PV (terminal value)": pv_terminal, 
        "PV of Sum(Cashflows)" : pv_cashflows, "Total PV": pv_sum, "Cash" : cash, 
        "Debt" : self.inputs["Debt"][0], "Value of equity" : value_of_equity, 
        "Number of shares ": self.inputs["Shares"][0], "Value per share": value_per_share, 
        "Current price" : self.inputs["Price"][0]}
        return packaged_outputs

    def create_excel_sheet(self, filename):
        inputs = self.get_packaged_inputs()
        cashflow = self.get_cash_flows()
        outputs = self.get_value_per_share()        
        pass
        
if __name__ == "__main__":
    dcf = DCF("ge")
    npv = dcf.get_value_per_share()
    print(npv["Value per share"])
    print(npv["Current price"])