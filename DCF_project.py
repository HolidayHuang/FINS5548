import numpy as np
import pandas as pd
from scipy.optimize import fsolve

# 读取数据
hours = pd.read_excel("Rental_Hours.xlsx")
freq = pd.read_excel("Users_Frequency.xlsx")

# 合并数据
merged_data = pd.merge(hours, freq, how='outer', on='Date')

# 创建包含 150 个 LOE 值的列表
LOE_values = [9500] * 150

# 创建一个包含 NaN 值的列表，长度等于数据框的行数减去 150
nan_values = [np.nan] * (len(merged_data) - 150)

# 将两个列表合并为一个
combined_values = LOE_values + nan_values

# 将合并的值分配给数据框的 'LOE' 列
merged_data['LOE'] = combined_values

# 计算 'Our_Revenue' 列
merged_data['Our_Revenue'] = merged_data['Rental_hours'] * 6 * merged_data['Users_frequency']

# 设置 Net Revenue (NR)
nr = 0.20

# 计算 'Our_Revenue_NR' 列
merged_data['Our_Revenue_NR'] = merged_data['Rental_hours'] * 6 * merged_data['Users_frequency']

# 设置 LOE 和 GST 值
LOE = 550
GST = 0.10

# 填充 'LOE' 和 'GST' 列
merged_data['LOE'] = [LOE] * len(merged_data)
merged_data['GST'] = [GST * merged_data['Our_Revenue_NR'][i] for i in range(len(merged_data))]

# 填充 'D&CCAPEX' 列和计算 'Cash Flow'、'Cumulative Cash Flows' 列
merged_data['D&CCAPEX'] = [0] * len(merged_data)
merged_data.loc[0, 'D&CCAPEX'] = 6159101
merged_data['Cash Flow'] = merged_data['Our_Revenue_NR'] - merged_data['D&CCAPEX'] - merged_data['GST']
merged_data['Cumulative Cash Flows'] = np.cumsum(merged_data['Cash Flow'])

# 计算 'Cash Flow' 列的总和
total_cash_flow = merged_data['Cash Flow'].sum()

# 创建 'Time' 列
merged_data['Time'] = [i for i in range(len(merged_data))]
t = merged_data['Time'].values

def npv(irr, cashfs, t):
    x = np.sum(cashfs / (1 + irr) ** t)
    return round(x, 3)

def irr(cashfs, t, x0, **kwargs):
    y = (fsolve(npv, x0=0.1, args=(cashfs, t), **kwargs)).item()
    return round(y * 100, 4)

irr(cashfs=merged_data['Cash Flow'].values, t=t, x0=0.05)


def payback():
    final_full_year = merged_data[merged_data['Cumulative Cash Flows'] < 0].index.max()
    fractional_yr = -merged_data['Cumulative Cash Flows'][final_full_year] / merged_data['Cash Flow'][final_full_year + 1]
    period = final_full_year + fractional_yr
    return round(period, 1)

def moic():
    z = total_cash_flow / merged_data['D&CCAPEX'].sum()
    return round(z, 1)

# 输出计算结果
print("NPV : $", npv(irr=0.10, cashfs=merged_data['Cash Flow'].values, t=t))
print("IRR : ", irr(cashfs=merged_data['Cash Flow'].values, t=t, x0=0.10), "%")
print("Payback Period : ", payback(), "Months")
print("MOIC : ", moic(), "x")
