DCF Builder
An automated discounted cash flow model builder using the SEC's Edgar API: https://www.sec.gov/edgar/sec-api-documentation
DCF is built using "fcffsimpleginzu.xlsx" template, from Ashwath Damodaran's public DCF spreadsheets: http://pages.stern.nyu.edu/~adamodar/New_Home_Page/spreadsh.htm 

Example: 
To generate a DCF for Oracle Corporation, run: 

from dcf import Company
oracle_dcf = DCF("orcl")
oracle_dcf.create_excel_sheet(orcl.xlsx)

