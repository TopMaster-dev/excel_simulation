from vba_function import monthly_repayment_fs, set_rate_fs, calc_rent_fs, repair_fund_fs, sell_price_fs

def IF(condition, true_value, false_value):
    if condition:
        return true_value
    else:
        return false_value

def AND(condition1, condition2):
    return condition1 and condition2

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

def MOD(value, divisor):
    return value % divisor

def IFERROR(value, error_value):
    try:
        return value
    except Exception:
        return error_value

def OR_Empty_Value(v1, v2):
    return 1 if (v1 == "" or v2 == "") else 0

def SUMIF(range_list, criteria, sum_range):
    return sum(g for l, g in zip(range_list, sum_range) if l == criteria)

def VLOOKUP(date_key, table, col_index, match_type):
    if(not match_type):
        for row in table:
            if row[0] == date_key:
                return row[col_index-1]
        return None

def AVERAGEIF(range_list, criteria, value_list):
    nums = [v for r, v in zip(range_list, value_list) if r == criteria and v is not None]
    return sum(nums) / len(nums) if nums else 0

def MonthlyRepayment_FS(e_month, d_method, F1, F2, d_month, d_rate, pre_pay, pre_repay, dbt_remain):
    return monthly_repayment_fs(e_month, d_method, F1, F2, d_month, d_rate, pre_pay, pre_repay, dbt_remain)

def calcrent_fs(passed_month, d_date, pre_rent, year_st, month_st, rgA, rgB):
    return calc_rent_fs(passed_month, d_date, pre_rent, year_st, month_st, rgA, rgB)

def SetRate_FS(passed_month, d_date, pre_rate, yearly_interval, add_rate, year_s, year_e):
    return set_rate_fs(passed_month, d_date, pre_rate, yearly_interval, add_rate, year_s, year_e)

def RepairFund_FS(passed_year, d_year, buy_year, buy_month, base_fund, yearly_interval, add_rate, year_s, year_e):
    return repair_fund_fs(passed_year, d_year, buy_year, buy_month, base_fund, yearly_interval, add_rate, year_s, year_e)

def SellPrice_FS(passed_year, d_year, pre_price, yearly_interval, add_rate, year_s, year_e):
    return sell_price_fs(passed_year, d_year, pre_price, yearly_interval, add_rate, year_s, year_e)