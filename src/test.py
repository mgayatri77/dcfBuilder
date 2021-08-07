from DCF import DCF

apple_dcf = DCF("aapl")
apple_dcf.create_excel_sheet("Apple.xlsx")
print(apple_dcf.compute_npv_outputs()['Value per Share'])