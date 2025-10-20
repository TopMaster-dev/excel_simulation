"""
Python port of the provided VBA module.
Contains functions named similarly to the original VBA functions:
 - stop_ctrl_fs, move_ctrl_fs (no-ops for Python/OpenPyXL)
 - reset_data_fs (works with openpyxl worksheet if provided)
 - set_chart_object_fs (no-op / placeholder)
 - set_rate_fs
 - monthly_repayment_fs
 - calc_repayment_fs
 - calc_rent_fs
 - repair_fund_fs
 - get_pre_fund_fs
 - sell_price_fs

Financial helper functions (pmt, _balance_before_period, ipmt, ppmt) are included.
"""

from typing import Optional, Any, Tuple
import math

# # Optional Excel helper (only if you want to actually open/modify XLSX files)
# try:
#     from openpyxl.worksheet.worksheet import Worksheet
#     OPENPYXL_AVAILABLE = True
# except Exception:
#     Worksheet = Any
#     OPENPYXL_AVAILABLE = False


# -------------------------
# UI / Excel environment helpers (no-op for Python)
# -------------------------
# def stop_ctrl_fs() -> None:
#     """
#     In VBA: Application.ScreenUpdating = False, EnableEvents = False, Calculation = xlCalculationManual
#     In Python: No equivalent (openpyxl can't control Excel UI); this is a no-op placeholder.
#     """
#     # no-op in Python
#     return None


# def move_ctrl_fs() -> None:
#     """
#     In VBA: Application.Calculation = xlCalculationAutomatic, EnableEvents = True, ScreenUpdating = True
#     In Python: No equivalent; placeholder.
#     """
#     # no-op in Python
#     return None


# def set_chart_object_fs(*args, **kwargs) -> None:
#     """
#     Placeholder for chart scaling logic in VBA. In Python one would need a charting library.
#     Kept as a no-op to preserve API parity.
#     """
#     return None


# # -------------------------
# # Helpers for Excel-range operations (openpyxl)
# # -------------------------
# def reset_data_fs(ws: Optional[Worksheet]) -> None:
#     """
#     Clears ranges in sheet named equivalent to VBA ResetData_FS.
#     If ws is None or openpyxl not available, does nothing.
#     VBA cleared:
#       Range(Cells(4,5) to Cells(10,5))
#       Range(Cells(13,5) to Cells(47,5))
#       Range(Cells(8,7) to Cells(8,7))
#       Range(Cells(15,7) to Cells(41,7))
#     (VBA used 1-based row/col)
#     """
#     stop_ctrl_fs()
#     if ws is None or not OPENPYXL_AVAILABLE:
#         move_ctrl_fs()
#         return

#     # clear ranges (openpyxl uses 1-based indexing via ws.cell)
#     def clear_range(r1, c1, r2, c2):
#         for r in range(r1, r2 + 1):
#             for c in range(c1, c2 + 1):
#                 ws.cell(row=r, column=c).value = None

#     clear_range(4, 5, 10, 5)
#     clear_range(13, 5, 47, 5)
#     clear_range(8, 7, 8, 7)
#     clear_range(15, 7, 41, 7)
#     move_ctrl_fs()


# -------------------------
# Financial helper functions (PMT, IPMT, PPMT equivalents)
# -------------------------
def pmt(rate_per_period: float, nper: int, pv: float, fv: float = 0.0, when: int = 0) -> float:
    """
    Excel-like PMT. when=0 means payments at period end (Excel default).
    Returns payment (negative for typical finance library; we return positive payment amount).
    """
    if nper <= 0:
        return 0.0
    if abs(rate_per_period) < 1e-12:
        # zero interest
        return pv / nper
    r = rate_per_period
    payment = r * (pv * (1 + r) ** nper + fv) / ((1 + r * when) * ((1 + r) ** nper - 1))
    return payment


def _balance_before_period(pv: float, rate: float, nper: int, period: int, payment: float) -> float:
    """
    Balance immediately before `period` (1-based) payment.
    formula: balance_before = pv*(1+r)^(period-1) - payment * ( ((1+r)^(period-1)-1)/r )
    Handles r == 0 specially.
    """
    if period <= 1:
        return pv
    if abs(rate) < 1e-12:
        return pv - payment * (period - 1)
    return pv * (1 + rate) ** (period - 1) - payment * (((1 + rate) ** (period - 1) - 1) / rate)


def ipmt(rate: float, per: int, nper: int, pv: float, fv: float = 0.0, when: int = 0) -> float:
    """
    Interest portion for period `per` (1-based). Returns positive interest amount.
    """
    if nper <= 0:
        return 0.0
    pay = pmt(rate, nper, pv, fv, when)
    bal_before = _balance_before_period(pv, rate, nper, per, pay)
    # interest = balance_before * rate
    return bal_before * rate


def ppmt(rate: float, per: int, nper: int, pv: float, fv: float = 0.0, when: int = 0) -> float:
    """
    Principal portion for period `per` (1-based). Returns positive principal amount.
    principal = payment - interest
    """
    payment = pmt(rate, nper, pv, fv, when)
    interest = ipmt(rate, per, nper, pv, fv, when)
    principal = payment - interest
    return principal


# -------------------------
# Ported VBA functions
# -------------------------
def set_rate_fs(passed_month: int,
                d_date,  # datetime.date or compatible (we only need date comparisons)
                pre_rate: float,
                yearly_interval: Optional[int],
                add_rate: Optional[float],
                year_s: Optional[int],
                year_e: Optional[int]) -> float:
    """
    Port of SetRate_FS.
    - passed_month: integer months since some origin
    - d_date: date (Python date or datetime.date)
    - pre_rate: base rate (float)
    - yearly_interval: integer interval (years) or None/empty
    - add_rate: value to add to pre_rate (float) or None/empty
    - year_s, year_e: optional start and end years (ints) to bound when the rate can change
    Returns the rate (float).
    """
    # converter 2025-10-01
    if isinstance(d_date, str):
        # Convert "2025年10月" or "2025-10-01" to a Python date
        try:
            if "年" in d_date and "月" in d_date:
                # Format: 2025年10月 or 2025年10月1日
                import re
                m = re.match(r"(\d{4})年(\d{1,2})月", d_date)
                if m:
                    year = int(m.group(1))
                    month = int(m.group(2))
                    d_date = __import__("datetime").date(year, month, 1)
            else:
                # Try parsing as "YYYY-MM-DD"
                d_date = __import__("datetime").datetime.strptime(d_date, "%Y-%m-%d").date()
        except Exception:
            pass
    # Check missing inputs (VBA's IsNull/"" checks)
    if yearly_interval is None or yearly_interval == "" or add_rate is None or add_rate == "":
        return pre_rate

    # Guard conversions
    try:
        yi = int(yearly_interval)
    except Exception:
        return pre_rate

    # Build date bounds if provided
    date_s = None
    date_e = None
    if year_s is not None and year_s != "":
        date_s = __import__("datetime").date(int(year_s), 1, 1)
    if year_e is not None and year_e != "":
        date_e = __import__("datetime").date(int(year_e), 12, 31)
    
    # if d_date is before start or after end, return pre_rate
    if date_s is not None and d_date < date_s:
        return pre_rate
    if date_e is not None and d_date > date_e:
        return pre_rate

    # Use passed_month mod (12 * yearly_interval) == 0
    if (passed_month % (12 * yi)) == 0:
        return pre_rate + float(add_rate)
    else:
        return pre_rate


def calc_repayment_fs(e_month: int, dbt_remain: float, d_month: int, d_rate: float) -> int:
    """
    Port of CalcRepayment_FS
    - e_month: end months (nper)
    - dbt_remain: remaining debt (pv)
    - d_month: starting month offset (so nper = e_month - d_month + 1)
    - d_rate: yearly rate in percent (e.g., 1.5 for 1.5%)
    Returns integer: principal + interest for the first payment of the sub-term (rounded down as VBA Int).
    """
    nper = e_month - d_month + 1
    if nper <= 0:
        return 0
    monthly_rate = d_rate / 100.0 / 12.0
    # period = 1 per VBA call
    per = 1

    # Using Excel approach: PPmt and IPmt results multiplied by -1 then Int
    # We return positive value; VBA did: dPrinc = Int(PPmt(...) * (-1)), same for dIntre.
    # We'll compute principal and interest positive and sum, then floor to mimic Int()
    payment = pmt(monthly_rate, nper, dbt_remain, 0.0, 0)
    interest = ipmt(monthly_rate, per, nper, dbt_remain, 0.0, 0)
    principal = ppmt(monthly_rate, per, nper, dbt_remain, 0.0, 0)

    # In VBA they took Int( value * (-1) ) where PPmt returns negative numbers in Excel convention.
    # Here principal and interest are positive; we use int() (trunc toward zero) similar to VBA Int for positives.
    d_princ = int(math.floor(principal + 1e-9))
    d_intre = int(math.floor(interest + 1e-9))
    return d_princ + d_intre


def monthly_repayment_fs(e_month: int,
                         d_method: int,
                         F1: int,
                         F2: int,
                         d_month: int,
                         d_rate: float,
                         pre_pay: Optional[float],
                         pre_repay: float,
                         dbt_remain: float) -> float:
    """
    Port of MonthlyRepayment_FS returning a numeric result:
    - e_month: allowed months (120 - 540)
    - d_method: method flag, 1 or other
    - F1: flag whether to calculate repayment (1 = calc)
    - F2: presence flag (if 0 -> return 0)
    - d_month: starting month offset
    - d_rate: yearly interest percent (e.g. 1.5)
    - pre_pay: prepayment amount (could be None/0)
    - pre_repay: pre-specified repay amount
    - dbt_remain: remaining debt
    """
    # Basic validation like VBA
    if e_month < 120 or e_month > 540:
        return 0.0
    if F2 == 0:
        return 0.0

    # If final month or preRepay/prePay >= remaining debt
    if d_month == e_month or pre_repay >= dbt_remain or (pre_pay is not None and pre_pay >= dbt_remain):
        if pre_pay and pre_pay > 0:
            result = int(dbt_remain * d_rate / 100.0 / 12.0)
            return float(result)
        else:
            result = dbt_remain + int(dbt_remain * d_rate / 100.0 / 12.0)
            return float(result)

    # Decide base repayment amount
    if F1 == 1:
        d_repay = calc_repayment_fs(e_month, dbt_remain, d_month, d_rate)
    else:
        d_repay = pre_repay

    if not pre_pay or pre_pay == 0:
        return float(d_repay)
    else:
        # Prepayment handling
        if d_method == 1:
            # if debt <= pre_pay then monthly repayment becomes interest only on remaining balance
            if dbt_remain <= pre_pay:
                return float(int(dbt_remain * d_rate / 100.0 / 12.0))
            else:
                return float(d_repay)
        else:
            # recalc repayment with reduced principal
            # call CalcRepayment_FS with dbt_remain - pre_pay
            return float(calc_repayment_fs(e_month, dbt_remain - pre_pay, d_month, d_rate))


def calc_rent_fs(passed_month: int,
                 d_date,  # datetime.date
                 pre_rent: float,
                 year_st: int,
                 month_st: int,
                 rgA: Tuple[Any, Any, Any],
                 rgB: Tuple[Any, Any, Any]) -> float:
    """
    Port of CalcRent_FS
    rgA and rgB are expected to be tuples/lists or objects with .value:
      rgA[0] -> yearlyInterval
      rgA[1] -> start year (year part)
      rgA[2] -> afterYear (years after purchase when rent changes)
      rgB[0] -> addRate (note VBA multiplies by -1)
      rgB[1] -> end year
      rgB[2] -> newRent
    The function returns the rent (rounded similarly to VBA).
    """
    if isinstance(d_date, str):
        # Convert "2025年10月" or "2025-10-01" to a Python date
        try:
            if "年" in d_date and "月" in d_date:
                # Format: 2025年10月 or 2025年10月1日
                import re
                m = re.match(r"(\d{4})年(\d{1,2})月", d_date)
                if m:
                    year = int(m.group(1))
                    month = int(m.group(2))
                    d_date = __import__("datetime").date(year, month, 1)
            else:
                # Try parsing as "YYYY-MM-DD"
                d_date = __import__("datetime").datetime.strptime(d_date, "%Y-%m-%d").date()
        except Exception:
            pass
    # Extract values (accept values directly or .value objects)
    def val(x):
        return getattr(x, "value", x)

    yearly_interval = val(rgA[0])
    add_rate = val(rgB[0])
    date_s = None
    date_e = None

    if val(rgA[1]) not in (None, ""):
        date_s = __import__("datetime").date(int(val(rgA[1])), 1, 1)
    if val(rgB[1]) not in (None, ""):
        date_e = __import__("datetime").date(int(val(rgB[1])), 12, 31)

    buy_date = __import__("datetime").date(int(year_st), int(month_st), 1)
    after_year = val(rgA[2]) if val(rgA[2]) not in (None, "") else None
    new_rent = val(rgB[2]) if val(rgB[2]) not in (None, "") else None

    # In VBA they did addRate = rgB(1).Value * (-1)
    if add_rate not in (None, ""):
        try:
            add_rate = float(add_rate) * (-1)
        except Exception:
            add_rate = None

    # If no yearly interval or no add_rate -> return pre_rent
    if yearly_interval in (None, "") or add_rate in (None, ""):
        result = pre_rent
    else:
        yi = int(yearly_interval)
        # date bounds
        if date_s is not None and d_date < date_s:
            result = pre_rent
        elif date_e is not None and d_date > date_e:
            result = pre_rent
        else:
            # rent increase every (yearly_interval * 12) months
            if (passed_month % (yi * 12)) == 0:
                # Round(preRent * (1 + addRate/100), 0)
                result = round(pre_rent * (1.0 + add_rate / 100.0))
            else:
                result = pre_rent

    # ChangeRent section: afterYear/newRent override
    if after_year not in (None, "") and new_rent not in (None, ""):
        try:
            if passed_month == int(after_year) * 12:
                result = float(new_rent)
        except Exception:
            pass

    return float(result)


def get_pre_fund_fs(base_fund: float, add_rate: float, add_ct: int) -> int:
    """
    Port of GetPreFund_FS - computes preFund after add_ct increments (compound integer truncation per iteration).
    VBA did: preFund = Int(preFund * (1 + addRate / 100)) per iteration.
    We mimic by applying floor on each iteration.
    """
    pre_fund = base_fund
    for _ in range(add_ct):
        pre_fund = math.floor(pre_fund * (1.0 + add_rate / 100.0))
    return int(pre_fund)


def repair_fund_fs(passed_year: int,
                   d_year: int,
                   buy_year: int,
                   buy_month: int,
                   base_fund: float,
                   yearly_interval: Optional[int],
                   add_rate: Optional[float],
                   year_s: Optional[int],
                   year_e: Optional[int]) -> float:
    """
    Port of RepairFund_FS
    - passed_year: integer (count of years since origin)
    - d_year: target calendar year (used to build a date to compare to yearS/yearE)
    - buy_year, buy_month: purchase date parts (not heavily used in port)
    - base_fund: base monthly fund (per month?) The VBA multiplies by 12 at the end.
    - yearly_interval, add_rate, year_s, year_e: same semantics as other functions.
    Returns fund * 12 (per VBA).
    """

    # If interval/addRate missing return baseFund * 12
    if yearly_interval in (None, "") or add_rate in (None, ""):
        return float(base_fund * 12.0)

    monthly_rate = add_rate
    yi = int(yearly_interval)
    d_date = __import__("datetime").date(int(d_year), 1, 1)

    date_s = None
    date_e = None
    if year_s not in (None, ""):
        date_s = __import__("datetime").date(int(year_s), 1, 1)
    if year_e not in (None, ""):
        date_e = __import__("datetime").date(int(year_e), 12, 31)

    # Calculate addCt based on VBA logic
    add_ct = (passed_year - 1) // yi
    pre_fund = get_pre_fund_fs(base_fund, add_rate, add_ct)

    # If yearS present, VBA recalculates addCt with (Year(dDate) + yi -1 - yearS) \ yi
    if year_s not in (None, ""):
        add_ct = (d_date.year + yi - 1 - int(year_s)) // yi
        if add_ct < 0:
            add_ct = 0
        pre_fund = get_pre_fund_fs(base_fund, add_rate, add_ct)
        if d_date < date_s:
            return float(pre_fund * 12)

    if year_e not in (None, ""):
        if d_date > date_e:
            add_ct = (int(year_e) + yi - 1 - int(year_s)) // yi
            if add_ct < 0:
                add_ct = 0
            pre_fund = get_pre_fund_fs(base_fund, add_rate, add_ct)
            return float(pre_fund * 12)

    # If passed_year mod yearly_interval == 0 -> special formula
    if (passed_year % yi) == 0:
        # RepairFund_FS = preFund * 12 * (1 + addRate / 100 * (12 - buyMonth + 1) / 12)
        return float(pre_fund * 12 * (1.0 + add_rate / 100.0 * (12 - int(buy_month) + 1) / 12.0))
    else:
        return float(pre_fund * 12)


def sell_price_fs(passed_year: int,
                  d_year: int,
                  pre_price: float,
                  yearly_interval: Optional[int],
                  add_rate: Optional[float],
                  year_s: Optional[int],
                  year_e: Optional[int]) -> float:
    """
    Port of SellPrice_FS:
    - If passed_year mod yearly_interval == 0 => price * (1 - add_rate/100)
    - Observes yearS/yearE bounds
    """
    # Basic checks
    if yearly_interval in (None, "") or add_rate in (None, ""):
        return float(pre_price)

    yi = int(yearly_interval)
    d_date = __import__("datetime").date(int(d_year), 1, 1)
    date_s = None
    date_e = None
    if year_s not in (None, ""):
        date_s = __import__("datetime").date(int(year_s), 1, 1)
    if year_e not in (None, ""):
        date_e = __import__("datetime").date(int(year_e), 12, 31)

    if date_s is not None and d_date < date_s:
        return float(pre_price)
    if date_e is not None and d_date > date_e:
        return float(pre_price)

    if (passed_year % yi) == 0:
        return float(pre_price * (1.0 - add_rate / 100.0))
    else:
        return float(pre_price)
