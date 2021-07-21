from DCF import DCF

apple_dcf = DCF("aapl")
apple_dcf.create_excel_sheet("apple_dcf.xlsx")
print(apple_dcf.get_npv_outputs()['Value per Share'])