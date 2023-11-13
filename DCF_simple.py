import numpy as np
import pandas as pd
from scipy.optimize import fsolve


prices = pd.read_excel("Pricing_Oil_Gas_NGL.xlsx")

production = pd.read_excel("Production_Oil_Gas_NGL.xlsx")

merged_data = pd.merge(prices, production, how='outer', on='Date')
merged_data['Oil_Revenue'] = merged_data['Oil_Price'] * merged_data['Oil_Prod']
merged_data['NGL_Revenue'] = merged_data['NGL_Price'] * merged_data['NGL_Prod']
merged_data['Gas_Revenue'] = merged_data['Gas_Price'] * merged_data['Gas_Prod']


nri = 0.75 #Net Revenue Interest
merged_data['Oil_Rev_NRI'] = merged_data['Oil_Price'] * merged_data['Oil_Prod']
merged_data['NGL_Rev_NRI'] = merged_data['NGL_Price'] * merged_data['NGL_Prod']
merged_data['Gas_Rev_NRI'] = merged_data['Gas_Price'] * merged_data['Gas_Prod']
merged_data['Total_Revenue'] = merged_data['Oil_Rev_NRI'] + merged_data['NGL_Rev_NRI'] + merged_data['Gas_Rev_NRI']
merged_data[['Date','Oil_Rev_NRI','NGL_Rev_NRI','Gas_Rev_NRI','Total_Revenue']]




LOE = 9500
Ad_Valorem_Tax = 0.025
Severance_Tax_Oil = 0.046
Severance_Tax_NGL = 0.075
Severance_Tax_Gas = 0.075

merged_data['LOE'] = [LOE for i in range(293)]
merged_data['Ad_Valorem_Tax'] = [Ad_Valorem_Tax * merged_data['Total_Revenue'][i] for i in range(293)]
merged_data['Sev_Tax_Oil'] = [Severance_Tax_Oil * merged_data['Total_Revenue'][i] for i in range(293)]
merged_data['Sev_Tax_NGL'] = [Severance_Tax_NGL * merged_data['Total_Revenue'][i] for i in range(293)]
merged_data['Sev_Tax_Gas'] = [Severance_Tax_Gas * merged_data['Total_Revenue'][i] for i in range(293)]
merged_data[['Date','LOE','Ad_Valorem_Tax','Sev_Tax_Oil','Sev_Tax_NGL','Sev_Tax_Gas']]




merged_data['D&CCAPEX'] = [0 for i in range(293)]
merged_data.loc[0, 'D&CCAPEX'] = 6159101
merged_data['Cash Flow'] = merged_data['Total_Revenue'] - merged_data['D&CCAPEX'] - merged_data['Ad_Valorem_Tax'] - merged_data['Sev_Tax_Oil'] - merged_data['Sev_Tax_NGL'] - merged_data['Sev_Tax_Gas']
merged_data['Cumulative Cash Flows'] = np.cumsum(merged_data['Cash Flow'])
merged_data[['Date','D&CCAPEX','Cash Flow','Cumulative Cash Flows']]



merged_data['Cash Flow'].sum()

cash_flow = merged_data['Cash Flow'].values
merged_data['Time'] = [i for i in range(293)]
t = merged_data['Time'].values


def npv(irr, cashfs, t):
    x = np.sum(cashfs / (1 + irr) ** t)
    return round(x, 3)


def irr(cashfs, t, x0, **kwargs):
    y = (fsolve(npv, x0=x0, args=(cashfs, t), **kwargs)).item()
    return round(y * 100, 4)


def payback():
    final_full_year = merged_data[merged_data['Cumulative Cash Flows'] < 0].index.values.max()
    fractional_yr = -merged_data['Cumulative Cash Flows'][final_full_year] / merged_data['Cash Flow'][
        final_full_year + 1]
    period = final_full_year + fractional_yr
    return round(period, 1)


def moic():
    z = merged_data['Cash Flow'].sum() / merged_data['D&CCAPEX'].sum()
    return round(z, 1)


print("NPV : $", npv(irr=0.10, cashfs=cash_flow, t=t))
print("IRR : ", irr(cashfs=cash_flow, t=t, x0=0.10), "%")
print("Payback Period : ", payback(), "Months")
print("MOIC : ", moic(), "x")

merged_data.to_csv("Model_Single_Well.csv")