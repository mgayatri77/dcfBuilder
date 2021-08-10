from openpyxl import Workbook
from openpyxl.utils.cell import get_column_letter
import constants
from util import percent_to_dec
from util import get_excel_cell

class compute: 
    "Class to internally compute DCF model outputs"
    def __init__(self, ticker, inputs, num_years=10, rates=constants.Default_rates):
        self.ticker = ticker
        self.num_years = num_years
        self.rates = rates
        self.inputs = inputs

    # computes revenue and revenue growth rate rows internally as list
    # @return revenues, revenue_growth - revenue and revenue growth as list
    def compute_revenues(self): 
        revenues = [0]*(self.num_years+2)
        revenue_growth = [1]*(self.num_years+2)
        
        revenue_growth[0] = 0
        revenues[0] = int(self.inputs["Revenue"][0])

        for i in range(1, len(revenues)): 
            revenue_multiple = 1
            if(i == 1): 
                revenue_multiple += percent_to_dec(self.rates["growth_rate_1_year"])
            elif(i >= 2 and i <= 5):
                revenue_multiple += percent_to_dec(self.rates["growth_rate_2_to_5_year"])
            elif(i >= len(revenues)-1): 
                revenue_multiple += percent_to_dec(self.rates['risk_free_rate'])
            else: 
                fraction = percent_to_dec(self.rates["growth_rate_2_to_5_year"] - self.rates["risk_free_rate"])
                growth_rate = fraction*(self.num_years-i)/(self.num_years - 5)
                revenue_multiple += growth_rate + percent_to_dec(self.rates["risk_free_rate"])
            revenues[i] = revenues[i-1]*revenue_multiple
            revenue_growth[i] = revenue_multiple-1
        return revenues, revenue_growth

    # computes operating income, costs and expenses and operating
    # margin rows internally as list
    # @return costs_and_expenses, operating_incomes, operating_margins - 
    # costs and expenses operating income and operating_margins as lists
    def compute_operating_income(self, revenues):
        costs_and_expenses = [0]*(self.num_years+2)
        operating_incomes = [0]*(self.num_years+2)
        operating_margins = [0]*(self.num_years+2)

        c_and_e_theoretical = int(self.inputs["Revenue"][0]) - operating_incomes[0]
        if(c_and_e_theoretical != int(self.inputs['COGS'][0])):
            self.inputs["COGS"][0] = c_and_e_theoretical
        
        operating_incomes[0] = int(self.inputs["Operating Income"][0])
        costs_and_expenses[0] = int(self.inputs["COGS"][0])
        operating_margins[0] = 100*operating_incomes[0]/revenues[0]

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
            operating_incomes[i] = percent_to_dec(operating_margins[i])*revenues[i]
            costs_and_expenses[i] = revenues[i] - operating_incomes[i]

        return costs_and_expenses, operating_incomes, operating_margins

    # computes tax and tax rate rows internally as list
    # @return taxes, tax_rate - taxes and tax rates as lists
    def compute_taxes(self, operating_incomes):
        tax_rate = [0]*(self.num_years+2)
        taxes = [1]*(self.num_years+2)

        taxes[0] = self.inputs["Taxes"][0]
        tax_rate[0] = (taxes[0])/operating_incomes[0]
        marginal_tax_rate = percent_to_dec(self.rates["marginal_tax_rate"])

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
    
    # computes net income row internally as list
    # @return net_incomes - net income as list
    def compute_net_income(self, tax_rate, operating_incomes): 
        net_incomes = [0]*(self.num_years+2)
        for i in range(len(operating_incomes)):
            if(operating_incomes[i] > 0):  
                net_incomes[i] = operating_incomes[i]*(1-(tax_rate[i]))
            else: 
                net_incomes[i] = operating_incomes[i]
        return net_incomes
    
    # computes sales to capital ratio row internally as list
    # @return s_to_c_ratio - sales to capital ratio as list
    def compute_s_to_c_ratio(self): 
        s_to_c_ratio = [0]*(self.num_years+2)
        for i in range(1, len(s_to_c_ratio)): 
            if(i == 1): 
                s_to_c_ratio[i] = self.rates["sales_to_capital_ratio_1_year"]
            elif(i > 1 and i < 6): 
                s_to_c_ratio[i] = self.rates["sales_to_capital_ratio_2_to_5_year"]
            else: 
                s_to_c_ratio[i] = self.rates["sales_to_capital_ratio_6_to_10_year"]
        return s_to_c_ratio

    # computes sales to capital ratio row internally as list
    # @return reinvestment - reinvestment as list
    def compute_reinvestment(self, revenues, net_incomes, s_to_c_ratio): 
        reinvestment = [0]*(self.num_years+2)
        
        for i in range(1, len(reinvestment)): 
            if(i == 1): 
                if(revenues[i] > revenues[i-1]):
                    reinvestment[i] = (revenues[i] - revenues[i-1])/s_to_c_ratio[i]
            else: 
                reinvestment[i] = (revenues[i] - revenues[i-1])/s_to_c_ratio[i]
        return reinvestment

    # computes free cash flow rows internally as list
    # @return fcf - fcf as list
    def compute_free_cash_flow(self, net_incomes, reinvestment): 
        fcf = [0]*(self.num_years+2)
        for i in range(1, len(fcf)): 
            fcf[i] = net_incomes[i] - reinvestment[i]
        return fcf

    # computes cost of capital row internally as list
    # @return cost_of_capital - cost of capital as list
    def compute_cost_of_capital(self):
        cost_of_capital = [self.rates["initial_cost_of_capital"]]*(self.num_years+2)
        cost_of_capital[0] = 0
        return cost_of_capital
    
    # computes discount rate row internally as list
    # @return discount_factor - cumulative discount factor as list
    def compute_cumulative_discount_factor(self, cost_of_capital): 
        discount_factor = [0]*(self.num_years+2)
        discount_factor[0] = 1
        for i in range(1, len(discount_factor)): 
            discount_factor[i] = discount_factor[i-1]/(1+percent_to_dec(cost_of_capital[i]))
        return discount_factor
    
    # computes present value row internally as list
    # @return pv - present value as list
    def compute_present_value(self, fcf, discount_factor):
        pv = [0]*(self.num_years+2)
        for i in range(1, len(pv)-1): 
            pv[i] = fcf[i]*discount_factor[i]
        return pv

    # gets and all DCF input related values as lists and returns them in dictionary
    # @return dictionary containing lists of revenue, operating income, etc.  
    def compute_packaged_inputs(self): 
        revenues, revenue_growth = self.compute_revenues()
        c_and_e, op_income, op_margin = self.compute_operating_income(revenues)
        taxes, tax_rate = self.compute_taxes(op_income)
        net_incomes = self.compute_net_income(tax_rate, op_income)
        s_to_c_ratio = self.compute_s_to_c_ratio()
        reinvestment = self.compute_reinvestment(revenues, net_incomes, s_to_c_ratio)
        packaged_inputs = {"Revenue" : revenues, "Revenue Growth Rate" : revenue_growth, 
        "Costs and Expenses" : c_and_e, "Operating Income" : op_income, "Margin": op_margin, 
        "Taxes" : taxes, "Tax Rate" : tax_rate,"Net Income": net_incomes, 
        "Reinvestment" : reinvestment, "Sales to Capital Ratio" : s_to_c_ratio}
        return packaged_inputs

    # gets and all DCF cashflow related outputs as lists and returns them in dictionary
    # @return dictionary containing lists of free cash flow, discount factor, etc.  
    def compute_cash_flows(self):
        inputs = self.compute_packaged_inputs()
        fcf = self.compute_free_cash_flow(inputs["Net Income"], inputs["Reinvestment"])
        cost_of_capital = self.compute_cost_of_capital()
        cumulative_discount_factor = self.compute_cumulative_discount_factor(cost_of_capital)
        pv = self.compute_present_value(fcf, cumulative_discount_factor)
        
        cashflows = {"FCFF" : fcf, "Cost of Capital" : cost_of_capital, 
        "Cumulative Discount Factor" : cumulative_discount_factor, "PV (FCFF)": pv}
        return cashflows

    # computes all DCF npv output values and return them in a dictionary
    # @return dictionary containing npv outputs such as Sum of PV, Value of Equity etc. 
    def compute_npv_outputs(self): 
        cashflows = self.compute_cash_flows()
        tcf = cashflows["FCFF"][-1]
        terminal_coc = cashflows["Cost of Capital"][-1]
        terminal_value = tcf/percent_to_dec(terminal_coc-self.rates['risk_free_rate'])
        
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