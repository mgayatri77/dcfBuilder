# DCF Builder
A discounted cash flow model builder that uses the SEC's Edgar RESTful API for obtaining historical financial data. DCF is built using "fcffsimpleginzu.xlsx" template, from Ashwath Damodaran's public DCF spreadsheets: http://pages.stern.nyu.edu/~adamodar/New_Home_Page/spreadsh.htm 

## Examples: 
To generate a DCF for Oracle Corporation, run: 

```python
from dcf import Company
oracle_dcf = DCF("orcl")
oracle_dcf.create_excel_sheet("orcl.xlsx")
```

To print the estimated value per share, run: 

```python
from dcf import Company
oracle_dcf = DCF("orcl")
npv_output = oracle_dcf.get_npv_outputs()
print(npv_output['Value per Share'])
```
