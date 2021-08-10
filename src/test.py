from DCF import DCF

tickers = ["ddog", "twlo"]
for i in tickers: 
    dcf = DCF(i)
    #rates = {
    #    "operating_margin_1_year": 10.0, 
    #    "target_pretax_operating_margin": 60.0, 
    #    "marginal_tax_rate": 14.0,
    #    "growth_rate_1_year": 30.0,
    #    "growth_rate_2_to_5_year": 25.0
    #}
    #sdcf.set_rates(**rates)
    dcf.create_excel_sheet(i)
    print("%.2f" % dcf.compute_npv_outputs()['Value per Share'])