# define errorcode
ERRORCODE = -7

# links to SEC API
CIK_link = "https://www.sec.gov/files/company_tickers.json"
Company_Facts_link = "https://data.sec.gov/api/xbrl/companyfacts/$$CIK$$.json"
headers = {'User-Agent': 'Mohit Gayatri (mohitgayatri77@gmail.com)'}

# XBRL tags of required quantities as dictionary
XBRL_tags = {
    "Revenue" : ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax"],
    "COGS" : ["CostsAndExpenses"],
    "Operating Income" : ["OperatingIncomeLoss", 
    "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", 
    "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"],
    "Interest Expense" : ["InterestPaid"],
    "Taxes" : ["IncomeTaxExpenseBenefit"],
    "Depreciation" : ["Depreciation"],
    "Amortization" : ["AmortizationOfIntangibleAssets"],
    "Net Income" : ["NetIncomeLoss", "ProfitLoss"],
    "Shares" : ["CommonSharesOutstanding", "CommonStockSharesOutstanding", 
    "WeightedAverageNumberOfDilutedSharesOutstanding", "WeightedAverageNumberOfShareOutstandingBasicAndDiluted"],
    "Cash" : 	["CashAndCashEquivalentsAtCarryingValue"],
    "Current Marketable Securities" : ["MarketableSecuritiesCurrent"],
    "Noncurrent Marketable Securities" : ["MarketableSecuritiesNoncurrent"],
    "Debt" : ["DebtAndCapitalLeaseObligations", "DebtCurrent"], 
    "Current Long Term Debt" : ["LongTermDebtCurrent"],
    "Noncurrent Long Term Debt" : ["LongTermDebtNoncurrent"]
}

# Units of each quantity, only USD supported as currency for now 
Units_map = {
    "Revenue" :  "USD", 
    "COGS" :  "USD", 
    "Operating Income" :  "USD", 
    "Interest Expense" :  "USD", 
    "Taxes" : "USD", 
    "Depreciation" : "USD", 
    "Amortization" : "USD", 
    "Net Income" : "USD", 
    "Shares" : "shares", 
    "Cash" : "USD", 
    "Current Marketable Securities" : "USD", 
    "Noncurrent Marketable Securities" :"USD",   
    "Debt" :"USD", 
    "Current Long Term Debt" :"USD", 
    "Noncurrent Long Term Debt" :"USD"
}

# default rates used in DCF model
Default_rates = {
    "marginal_tax_rate" : 25.00,
    "growth_rate_1_year" : 50.00,
    "growth_rate_2_to_5_year" :	5.00,
    "operating_margin_1_year" : 3.00,
    "target_pretax_operating_margin" : 7.50,
    "year_of_convergence" : 10.00,
    "sales_to_capital_ratio_1_year" : 10.00,
    "sales_to_capital_ratio_2_to_5_year" : 5.00,
    "sales_to_capital_ratio_6_to_10_year" : 1.50,
    "risk_free_rate" : 1.69,
    "initial_cost_of_capital" : 6.74
}

# excel row names and numbers of list quantities 
excel_list_row_numbers = {
    'Quantities': 1, 'Revenue': 2, 'Revenue Growth Rate': 3, 'Costs and Expenses': 4, 
    'Operating Income': 5, 'Operating Margin': 6, 'Taxes': 7, 'Tax Rate': 8, 'Net Income': 9, 
    'Reinvestment': 10, 'FCFF': 11, 'Cost of Capital': 13, 'Cumulative Discount Factor': 14, 
    'PV (FCFF)': 15, 'Sales to Capital Ratio': 16
}

# excel row names and numbers of singular quantities 
excel_singular_row_numbers = {
    'Terminal Cash Flow': 18, 'Terminal Cost of Capital': 19, 'Terminal Value': 20, 
    'PV (Terminal Value)': 21, 'PV (Cashflows)': 22, 'Sum of PV': 23, 'Cash': 24, 'Debt': 25, 
    'Value of Equity': 26, 'Common Shares Outstanding': 27, 'Value per Share': 28, 
    'Current Price': 29
}

# excel column widths
excel_column_widths = {
    "big" : 28, 
    "small" : 16
}