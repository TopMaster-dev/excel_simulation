import excel_function as ef
from excel_log import log1_Sheet_All

data = {
    "入力!E4": "サンプル",
    "入力!E5": 30,
    "入力!E7": "ミラージュ",
    "入力!E8": 2025,
    "入力!G8": 9,
    "入力!E9": 43800000,
    "入力!E10": "",
    "入力!E12": 43800000,
    "入力!E13": 35,
    "入力!E14": 1.675,
    "入力!E15": 5,
    "入力!G15": 0.1,
    "入力!E16": 2026,
    "入力!G16": 2040,
    "入力!E17": 2,
    "入力!G17": 350000,
    "入力!E18": 1,
    "入力!E20": 210000,
    "入力!E21": 5,
    "入力!G21": 0.5,
    "入力!E22": 2026,
    "入力!G22": 2040,
    "入力!E23": 4,
    "入力!G23": 140000,
    "入力!E24": "",
    "入力!G24": "",
    "入力!E25": 8920,
    "入力!E26": 12090,
    "入力!E27": 2,
    "入力!G27": 0.5,
    "入力!E28": 2026,
    "入力!G28": 2036,
    "入力!E29": 0,
    "入力!E30": 10,
    "入力!G30": 150000,
    "入力!E31": 4,
    "入力!G31": 2,
    "入力!E32": 4,
    "入力!G32": 150000,
    "入力!E33": 126758716,
    "入力!E34": 21126452,
    "入力!E35": 42242905,
    "入力!E36": 5376,
    "入力!G36": 321776,
    "入力!E37": 5501100,
    "入力!E39": 43800000,
    "入力!G39": 100,
    "入力!E40": 3,
    "入力!G40": 1,
    "入力!E41": 2026,
    "入力!G41": 2040,
    "入力!E43": 3,
    "入力!E44": 1.4,
    "入力!E45": 0.3,
    "入力!E47": 15,
    "Log1!D6": 0
}

output = {}

def output_init_cell():
    global output
    key = "出力1!G2"
    value = data["入力!E7"]
    output[key] = value
    key = "出力1!O2"
    value = data["入力!E9"]
    output[key] = value
    key = "出力1!T2"
    value = data["入力!E10"]
    output[key] = value
    key = "出力1!T3"
    value = data["入力!E12"]
    output[key] = value
    key = "出力1!V3"
    value = data["入力!E13"]
    output[key] = value
    key = "出力1!W3"
    value = data["入力!E14"] / 100
    output[key] = value
    key = "入力!I11"
    value = 12 - data["入力!G8"] + 1
    output[key] = value
    key = "入力!I32"
    value = ef.IFERROR(int(data["入力!E33"] / 2 *  data["入力!E36"] / data["入力!G36"] * data["入力!E43"] / 100), 0)
    output[key] = value
    key = "入力!I33"
    value = int((data["入力!E37"] *  data["入力!E43"]) / 100)
    output[key] = value
    key = "入力!I34"
    value = ef.IFERROR(int(data["入力!E34"] * data["入力!E36"] / data["入力!G36"] * data["入力!E44"] / 100), 0)
    output[key] = value
    key = "入力!I35"
    value = int((data["入力!E37"] *  data["入力!E44"]) / 100)
    output[key] = value
    key = "入力!I36"
    value = ef.IFERROR(int(data["入力!E35"] * data["入力!E36"] / data["入力!G36"] * data["入力!E45"] / 100), 0)
    output[key] = value
    key = "入力!I37"
    value = int((data["入力!E37"] *  data["入力!E45"]) / 100)
    output[key] = value

def output_C_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!C{i}"
    if(i == 7):
        value = data["入力!E8"]
    else:
        value = output[f"{sheet_name}!C{i - 1}"] + 1
    output[key] = value

def output_D_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!D{i}"
    if(i == 7):
        value = 0
    else:
        value = output[f"{sheet_name}!D{i - 1}"] + 1
    output[key] = value

def output_E_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!E{i}"
    if(i == 7):
        value = data["入力!E5"]
    else:
        value = output[f"{sheet_name}!E{i - 1}"] + 1
    output[key] = value

def output_F_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!F{i}"
    range_list = []
    sum_range = []
    if(sheet_name == "出力2"):
        for l in range(6, 547):
            range_list.append(output[f"Log2!L{l}"])
            sum_range.append(output[f"Log2!G{l}"])
    else:
        for l in range(6, 547):
            range_list.append(output[f"Log1!L{l}"])
            sum_range.append(output[f"Log1!G{l}"])
    value = ef.SUMIF(range_list, output[f"{sheet_name}!C{i}"], sum_range)
    output[key] = value

def output_I_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!I{i}"
    if(i == 7):
        value = ""
    else:
        if(sheet_name == "出力2"):
            value = ef.IFERROR(ef.IF(ef.MOD(output[f"{sheet_name}!D{i}"], data["入力!E15"]) == 0, 1, 0), 0)
        else:
            value = 0
    output[key] = value

def output_H_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!H{i}"
    input_G15 = data["入力!G15"]
    if(i == 7):
        value = ""
    else:
        if(sheet_name == "出力2"):
            value = ef.IF(output[f"{sheet_name}!F{i}"] == 0, "", input_G15 * output[f"{sheet_name}!I{i}"])
        else:
            value = 0
    output[key] = value

def output_G_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!G{i}"
    input_E14 = data["入力!E14"]
    if(i == 7):
        value = input_E14
    else:
        if(sheet_name == "出力2"):
            range_list = []
            sum_range = []
            for l in range(6, 547):
                range_list.append(output[f"Log2!L{l}"])
                sum_range.append(output[f"Log2!F{l}"])
            value = ef.IF(output[f"{sheet_name}!F{i}"] > 0, ef.AVERAGEIF(range_list, output[f"{sheet_name}!C{i}"], sum_range), "")
        else:
            val = output[f"{sheet_name}!G{i - 1}"]
            if(val == ""):
                val = 0
            value = ef.IF(output[f"{sheet_name}!F{i}"] == 0, "", val + output[f"{sheet_name}!H{i}"])
    output[key] = value

def output_L_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!L{i}"
    input_E25 = data["入力!E25"]
    if(i == 7):
        value = input_E25 * output["入力!I11"] * output["入力!I11"] / 12
    elif(i == 8):
        value = input_E25 * 12
    else:
        value = output[f"{sheet_name}!L{i - 1}"]
    output[key] = value

def output_M_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!M{i}"
    input_E26 = data["入力!E26"]
    if(i == 7):
        value = input_E26 * output["入力!I11"]
    else:
        if(sheet_name == "出力2"):
            value = ef.RepairFund_FS(output[f"{sheet_name}!D{i}"], output[f"{sheet_name}!C{i}"], data["入力!E8"], data["入力!G8"], input_E26, data["入力!E27"], data["入力!G27"], data["入力!E28"], data["入力!G28"])
        else:
            value = input_E26 * 12
    output[key] = value

def output_O_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!O{i}"
    input_E29 = data["入力!E29"]
    if(i == 7):
        value = input_E29 * output["入力!I11"] * output["入力!I11"] / 12
    elif(i == 8):
        value = input_E29 * 12
    else:
        value = output[f"{sheet_name}!O{i - 1}"]
    output[key] = value

def output_Q_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!Q{i}"
    input_E30 = data["入力!E30"]
    if(i == 7):
        value = ""
    else:
        value = ef.IFERROR(ef.IF(ef.MOD(output[f"{sheet_name}!D{i}"], input_E30)==0, 1, 0), 0)
    output[key] = value

def output_P_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!P{i}"
    input_G30 = data["入力!G30"]
    if(i == 7):
        value = ""
    else:
        value = input_G30 * output[f"{sheet_name}!Q{i}"]
    output[key] = value

def output_S_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!S{i}"
    input_E32 = data["入力!E32"]
    if(i == 7):
        value = ""
    else:
        value = ef.IFERROR(ef.IF(ef.MOD(output[f"{sheet_name}!D{i}"], input_E32)==0, 1, 0), 0)
    output[key] = value

def output_R_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!R{i}"
    input_G32 = data["入力!G32"]
    if(i == 7):
        value = ""
    else:
        value = input_G32 * output[f"{sheet_name}!S{i}"]
    output[key] = value

def output_T_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!T{i}"
    if(i == 7):
        value = output["入力!I32"] + output["入力!I33"]
    else:
        value = ""
    output[key] = value

def output_J_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!J{i}"
    input_E10 = data["入力!E10"]
    if(input_E10 == ""):
        input_E10 = 0
    if(sheet_name == "出力2"):
        range_list = []
        sum_range = []
        for l in range(6, 547):
            range_list.append(output[f"Log2!L{l}"])
            sum_range.append(output[f"Log2!J{l}"])
        value = ef.SUMIF(range_list, output[f"{sheet_name}!C{i}"], sum_range)
    else:
        value = input_E10
    output[key] = value

def output_U_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!U{i}"
    input_E21 = data["入力!E21"]
    if(i == 7):
        value = ""
    else:
        value = output["入力!I34"] + output["入力!I35"] + output["入力!I36"] + output["入力!I37"]
    output[key] = value

def output_V_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!V{i}"
    t_value = output[f"{sheet_name}!T{i}"]
    u_value = output[f"{sheet_name}!U{i}"]
    p_value = output[f"{sheet_name}!P{i}"]
    r_value = output[f"{sheet_name}!R{i}"]
    if(t_value == ""):
        t_value = 0
    if(u_value == ""):
        u_value = 0
    if(p_value == ""):
        p_value = 0
    if(r_value == ""):
        r_value = 0
    value = float(output[f"{sheet_name}!F{i}"]) + float(output[f"{sheet_name}!J{i}"]) + float(output[f"{sheet_name}!L{i}"]) + float(output[f"{sheet_name}!M{i}"]) + float(output[f"{sheet_name}!O{i}"]) + float(p_value) + float(r_value) + float(t_value) + float(u_value)
    output[key] = value

def output_W_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!W{i}"
    range_list = []
    sum_range = []
    # for l in range(6, output["Log1!D3"] + 7):
    if(sheet_name == "出力2"):
        for l in range(6, 727):
            range_list.append(output[f"Log2!Z{l}"])
            sum_range.append(output[f"Log2!W{l}"])
    else:
        for l in range(6, 727):
            range_list.append(output[f"Log1!Z{l}"])
            sum_range.append(output[f"Log1!W{l}"])
    value = ef.SUMIF(range_list, output[f"{sheet_name}!C{i}"], sum_range)
    output[key] = value

def output_X_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!X{i}"
    input_I11 = output["入力!I11"]
    if(sheet_name == "出力2"):
        if(i == 7):
            value = output[f"{sheet_name}!W{i}"] / input_I11
        else:
            value = output[f"{sheet_name}!W{i}"] / 12
    else:
        value = data["入力!E20"]
    output[key] = value

def output_AA_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!AA{i}"
    input_E31 = data["入力!E31"]
    if(i == 7):
        value = ""
    else:
        value = ef.IFERROR(ef.IF(ef.MOD(output[f"{sheet_name}!D{i}"], input_E31)==0, 1, 0), 0)
    output[key] = value

def output_Z_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!Z{i}"
    input_G31 = data["入力!G31"]
    aa_value = output[f"{sheet_name}!AA{i}"]
    if(aa_value == ""):
        aa_value = 0
    value = output[f"{sheet_name}!X{i}"] * input_G31 * aa_value
    output[key] = value

def output_Y_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!Y{i}"
    input_E21 = data["入力!E21"]
    if(i == 7):
        value = ""
    else:
        value = ef.IFERROR(ef.IF(ef.MOD(output[f"{sheet_name}!D{i}"], input_E21)==0, 1, 0), 0)
    output[key] = value

def output_AB_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!AB{i}"
    range_list = []
    sum_range = []
    if(sheet_name == "出力2"):
        for l in range(6, 727):
            range_list.append(output[f"Log2!Z{l}"])
            sum_range.append(output[f"Log2!Y{l}"])
    else:
        for l in range(6, 727):
            range_list.append(output[f"Log1!Z{l}"])
            sum_range.append(output[f"Log1!Y{l}"])
    value = ef.SUMIF(range_list, output[f"{sheet_name}!C{i}"], sum_range)
    output[key] = value

def output_AC_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!AC{i}"
    input_E47 = data["入力!E47"]
    if(i == 67):
        value = ef.IF(0 > 65, 0, input_E47 * 10000)
    else:
        value = ef.IF(output[f"{sheet_name}!E{i}"] > 65, 0, input_E47 * 10000)
    output[key] = value

def output_AD_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!AD{i}"
    value = output[f"{sheet_name}!W{i}"] + output[f"{sheet_name}!AB{i}"] + output[f"{sheet_name}!AC{i}"]
    output[key] = value

def output_AE_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!AE{i}"
    value = output[f"{sheet_name}!AD{i}"] - output[f"{sheet_name}!V{i}"]
    output[key] = value

def output_AF_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!AF{i}"
    input_I11 = output["入力!I11"]
    if(i == 7):
        value = output[f"{sheet_name}!AE{i}"] / input_I11
    else:
        value = output[f"{sheet_name}!AE{i}"] / 12
    output[key] = value

def output_AG_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!AG{i}"
    if(i == 7):
        value = output[f"{sheet_name}!AE{i}"]
    else:
        value = output[f"{sheet_name}!AE{i}"] + output[f"{sheet_name}!AG{i - 1}"]
    output[key] = value

def output_AI_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!AI{i}"
    table = []
    if(sheet_name == "出力2"):
        for l in range(6, 547):
            col_table = []
            for m in range(0, 7):
                if(m == 0):
                    col_table.append(output[f"Log2!E{l}"])
                elif(m == 1):
                    col_table.append(output[f"Log2!F{l}"])
                elif(m == 2):
                    col_table.append(output[f"Log2!G{l}"])
                elif(m == 3):
                    col_table.append(output[f"Log2!H{l}"])
                elif(m == 4):
                    col_table.append(output[f"Log2!I{l}"])
                elif(m == 5):
                    col_table.append(output[f"Log2!J{l}"])
                elif(m == 6):
                    col_table.append(output[f"Log2!K{l}"])
            table.append(col_table)  
    else:
        for l in range(6, 547):
            col_table = []
            for m in range(0, 7):
                if(m == 0):
                    col_table.append(output[f"Log1!E{l}"])
                elif(m == 1):
                    col_table.append(output[f"Log1!F{l}"])
                elif(m == 2):
                    col_table.append(output[f"Log1!G{l}"])
                elif(m == 3):
                    col_table.append(output[f"Log1!H{l}"])
                elif(m == 4):
                    col_table.append(output[f"Log1!I{l}"])
                elif(m == 5):
                    col_table.append(output[f"Log1!J{l}"])
                elif(m == 6):
                    col_table.append(output[f"Log1!K{l}"])
            table.append(col_table)  
    value = ef.IFERROR(ef.VLOOKUP(ef.DATE(output[f"{sheet_name}!C{i}"], 12, 1), table, 7, False), 0)
    if(value == None):
        value = 0
    output[key] = value

def output_AJ_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!AJ{i}"
    input_E39 = data["入力!E39"]
    if(i == 7):
        value = input_E39
    else:
        if(sheet_name == "出力2"):
            value = ef.SellPrice_FS(output[f"{sheet_name}!D{i}"], output[f"{sheet_name}!C{i}"], output[f"{sheet_name}!AJ{i - 1}"], data["入力!E40"], data["入力!G40"], data["入力!E41"], data["入力!G41"])
        else:
            value = output[f"{sheet_name}!AJ{i - 1}"]
    output[key] = value

def output_AK_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!AK{i}"
    input_E9 = data["入力!E9"]
    value = output[f"{sheet_name}!AJ{i}"] / input_E9 * 100
    output[key] = value

def output_AM_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!AM{i}"
    input_E40 = data["入力!E40"]
    value = output[f"{sheet_name}!AG{i}"] + output[f"{sheet_name}!AJ{i}"] - output[f"{sheet_name}!AI{i}"]
    output[key] = value

def output_AE_total(sheet_name):
    global output
    key = f"{sheet_name}!AE67"
    sum_value = 0
    for i in range(7, 67):
        value = output[f"{sheet_name}!AE{i}"]
        sum_value += value
    output[key] = sum_value

def output_AL_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!AL{i}"
    if(i == 7):
        value = ""
    else:
        value = ef.IFERROR(ef.IF(ef.MOD(output[f"{sheet_name}!D{i}"], data["入力!E40"])==0, 1, 0), 0)
    output[key] = value

def outputAll_one_cell(sheet_name):
    global output
    for i in range(7, 67):
        output_C_cell(sheet_name, i)
        output_D_cell(sheet_name, i)
        output_E_cell(sheet_name, i)
        output_F_cell(sheet_name, i)
        output_I_cell(sheet_name, i)
        output_H_cell(sheet_name, i)
        output_G_cell(sheet_name, i)
        output_L_cell(sheet_name, i)
        output_M_cell(sheet_name, i)
        output_O_cell(sheet_name, i)
        output_Q_cell(sheet_name, i)
        output_P_cell(sheet_name, i)
        output_S_cell(sheet_name, i)
        output_R_cell(sheet_name, i)
        output_T_cell(sheet_name, i)
        output_U_cell(sheet_name, i)
        output_J_cell(sheet_name, i)
        output_V_cell(sheet_name, i)
        output_W_cell(sheet_name, i)
        output_X_cell(sheet_name, i)
        output_AA_cell(sheet_name, i)
        output_Z_cell(sheet_name, i)
        output_Y_cell(sheet_name, i)
        output_AB_cell(sheet_name, i)
        output_AC_cell(sheet_name, i)
        output_AD_cell(sheet_name, i)
        output_AE_cell(sheet_name, i)
        output_AF_cell(sheet_name, i)
        output_AG_cell(sheet_name, i)
        output_AI_cell(sheet_name, i)
        output_AJ_cell(sheet_name, i)
        output_AK_cell(sheet_name, i)
        output_AM_cell(sheet_name, i)
    output_AC_cell(sheet_name, 67)
    output_AE_total(sheet_name)
    return output

def output_ALL_cell():
    global output
    output = log1_Sheet_All("Log1")
    output = log1_Sheet_All("Log2")
    output_init_cell()
    outputAll_one_cell("出力1")
    outputAll_one_cell("出力2")
    return output