ERRORCODE = -7
CIK_link = "https://www.sec.gov/files/company_tickers.json"
Company_Facts_link = "https://data.sec.gov/api/xbrl/companyfacts/$$CIK$$.json"

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
    "Debt" : ["DebtAndCapitalLeaseObligations", "LongTermDebtAndCapitalLeaseObligations"], 
    "Current Long Term Debt" : ["LongTermDebtCurrent"],
    "Noncurrent Long Term Debt" : ["LongTermDebtNoncurrent"]
}
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