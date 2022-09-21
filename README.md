# DCF Builder
A tool that takes in a stock ticker (listed on NYSE/NASDAQ) and outputs a DCF model into an Excel sheet. The tool uses the SEC's Edgar RESTful API for obtaining historical financial data. The DCF is built using the format/calculation method specified in the "fcffsimpleginzu.xlsx" template, from Ashwath Damodaran's (Professor at NYU Stern) public DCF spreadsheets: http://pages.stern.nyu.edu/~adamodar/New_Home_Page/spreadsh.htm 

## Examples: 
Note, you must change the email ID in the headers variable in "constants.py" to your email address

To generate a DCF for Oracle Corporation, run: 

```python
from DCF import DCF
oracle_dcf = DCF("orcl")
oracle_dcf.create_excel_sheet("orcl.xlsx")
```

To print the estimated value per share, run: 

```python
from DCF import DCF
oracle_dcf = DCF("orcl")
npv_output = oracle_dcf.compute_npv_outputs()
print(npv_output['Value per Share'])
```
