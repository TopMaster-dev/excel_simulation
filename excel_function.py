from vba_function import monthly_repayment_fs

def IF(condition, true_value, false_value):
    if condition:
        return true_value
    else:
        return false_value

def SUM(values):
    return sum(values)

def AVERAGE(values):
    return sum(values) / len(values)

def MAX(values):
    return max(values)

def MIN(values):
    return min(values)

def DATE(year, month, day):
    if month > 12:
        year += 1
        month -= 12
    return f"{year}年{month}月"

def YEAR(date):
    year = date.split("年")[0]
    return int(year)

def MONTH(date):
    val = date.split("年")[1]
    month = val.split("月")[0]
    return int(month)

def MonthlyRepayment_FS(e_month, d_method, F1, F2, d_month, d_rate, pre_pay, pre_repay, dbt_remain):
    return monthly_repayment_fs(e_month, d_method, F1, F2, d_month, d_rate, pre_pay, pre_repay, dbt_remain)