import excel_function as ef
import json

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

def log_init_cell(sheet_name):
    global output
    key = f"{sheet_name}!D3"
    value = data["入力!E13"] * 12
    output[key] = value
    key = f"{sheet_name}!F3"
    value = data["入力!E18"]
    output[key] = value

def log_D_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!D{i}"
    value = i - 6
    output[key] = value

def log_C_cell(sheet_name, i):
    global output
    input_E13 = data["入力!E13"]
    Log_D = output[f"{sheet_name}!D{i}"]
    key = f"{sheet_name}!C{i}"
    value = ef.IF(input_E13 * 12 >= Log_D, 1, 0)
    output[key] = value

def log_E_cell(sheet_name, i):
    global output
    input_E8 = data["入力!E8"]
    input_G8 = data["入力!G8"]
    key = f"{sheet_name}!E{i}"
    if(i == 6):
        value = ef.DATE(input_E8, input_G8, 1)
    else:
        year = ef.YEAR(output[f"{sheet_name}!E{i-1}"])
        month = ef.MONTH(output[f"{sheet_name}!E{i-1}"])
        value = ef.DATE(year, month + 1, 1)
    output[key] = value

def log_F_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!F{i}"
    if(sheet_name == "Log2" and i > 6):
        value = ef.SetRate_FS(output[f"{sheet_name}!D{i}"], output[f"{sheet_name}!E{i}"], output[f"{sheet_name}!F{i - 1}"], data["入力!E15"], data["入力!G15"], data["入力!E16"], data["入力!G16"])
    else:
        value = data["入力!E14"]
    output[key] = value

def log_B_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!B{i}"
    if(i == 6):
        value = ""
    elif(i == 7):
        value = 1
    else:
        value = ef.IF(output[f"{sheet_name}!F{i}"] > output[f"{sheet_name}!F{i - 1}"], 1, 0)
    output[key] = value

def log_J_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!J{i}"
    value = 0
    if(sheet_name == "Log2" and i >= 7):
        value = ef.IFERROR(ef.IF(ef.MOD(output[f"{sheet_name}!D{i}"], data["入力!E17"] * 12) == 0, ef.IF(output[f"{sheet_name}!K{i - 1}"] < data["入力!G17"], output[f"{sheet_name}!K{i - 1}"], data["入力!G17"]), 0), 0)
    output[key] = value

def log_G_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!G{i}"
    input_log_D3 = int(output[f"{sheet_name}!D3"])
    input_log_F3 = int(output[f"{sheet_name}!F3"])
    if(i == 6):
        value = 0
    else:
        value = ef.MonthlyRepayment_FS(input_log_D3, input_log_F3, output[f"{sheet_name}!B{i}"], output[f"{sheet_name}!C{i}"], output[f"{sheet_name}!D{i}"], output[f"{sheet_name}!F{i}"], output[f"{sheet_name}!J{i}"], output[f"{sheet_name}!G{i - 1}"], output[f"{sheet_name}!K{i - 1}"])
    output[key] = value

def log_I_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!I{i}"
    if(i == 6):
        value = 0
    else:
        val = output[f"{sheet_name}!K{i - 1}"] * output[f"{sheet_name}!F{i}"]/100/12
        value = int(val)
    output[key] = value

def log_H_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!H{i}"
    if(i == 6):
        value = 0
    else:
        value = output[f"{sheet_name}!G{i}"] - output[f"{sheet_name}!I{i}"]
    output[key] = value

def log_K_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!K{i}"
    if(i == 6):
        value = data["入力!E12"]
    else:
        value = output[f"{sheet_name}!K{i - 1}"] - output[f"{sheet_name}!H{i}"] - output[f"{sheet_name}!J{i}"]
    output[key] = value

def log_L_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!L{i}"
    value = ef.YEAR(output[f"{sheet_name}!E{i}"])
    output[key] = value

def log_N_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!N{i}"
    if(i == 6):
        value = output[f"{sheet_name}!E{i}"]
    else:
        value = ef.DATE(ef.YEAR(output[f"{sheet_name}!N{i - 1}"]), ef.MONTH(output[f"{sheet_name}!N{i - 1}"]) + 1, 1)
    output[key] = value

def log_O_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!O{i}"
    value = ef.YEAR(output[f"{sheet_name}!N{i}"])
    if(i == 6):
        value = 0
    else:
        value = output[f"{sheet_name}!O{i - 1}"] + 1
    output[key] = value

def log_P_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!P{i}"
    value = ef.IFERROR(ef.MOD(output[f"{sheet_name}!O{i}"], data["入力!E31"] * 12) + 1, 0)
    output[key] = value

def log_Q_cell(sheet_name, i):
    global output
    input_E31 = data["入力!E31"]
    input_G31 = data["入力!G31"]
    key = f"{sheet_name}!Q{i}"
    value = ef.IF(ef.OR_Empty_Value(input_E31, input_G31), 1, ef.IF(output[f"{sheet_name}!O{i}"] < input_E31 * 12, 1, ef.IF(output[f"{sheet_name}!P{i}"] <= input_G31, 0, 1)))
    output[key] = value

def log_R_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!R{i}"
    if(i == 6):
        value = 1
    else:
        value = ef.IF(output[f"{sheet_name}!Q{i}"] == 0, 0, output[f"{sheet_name}!R{i - 1}"] + 1)
    output[key] = value

def log_S_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!S{i}"
    input_E24 = data["入力!E24"] if data["入力!E24"] != "" else 0
    input_G24 = data["入力!G24"] if data["入力!G24"] != "" else 0
    value = ef.IF(output[f"{sheet_name}!R{i}"] == (input_E24 * 12 + 1), input_G24,0)
    output[key] = value

def log_U_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!U{i}"
    input_E21 = data["入力!E21"]
    value = ef.IFERROR(ef.IF(output[f"{sheet_name}!O{i}"] < input_E21 * 12, 0, ef.IF(ef.MOD(output[f"{sheet_name}!O{i}"], input_E21 * 12) == 0, 1, 0)), 0)
    output[key] = value

def log_V_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!V{i}"
    input_E20 = data["入力!E20"]
    if(i == 6):
        value = input_E20
    else:
        if(sheet_name == "Log2"):
            array_value = [data["入力!E21"], data["入力!E22"], data["入力!E23"]]
            array_value_1 = [data["入力!G21"], data["入力!G22"], data["入力!G23"]]
            value = ef.calcrent_fs(output[f"{sheet_name}!O{i}"], output[f"{sheet_name}!N{i}"], output[f"{sheet_name}!V{i - 1}"], data["入力!E8"], data["入力!G8"], array_value, array_value_1)
        else:
            value = output[f"{sheet_name}!V{i - 1}"]
    output[key] = value

def log_W_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!W{i}"
    value = output[f"{sheet_name}!V{i}"] * output[f"{sheet_name}!Q{i}"]
    output[key] = value

def log_X_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!X{i}"
    input_E29 = data["入力!E29"]
    value = input_E29 * output[f"{sheet_name}!Q{i}"]
    output[key] = value

def log_Y_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!Y{i}"
    value = int(output[f"{sheet_name}!V{i}"] * output[f"{sheet_name}!S{i}"])
    output[key] = value

def log_Z_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!Z{i}"
    value = ef.YEAR(output[f"{sheet_name}!N{i}"])
    output[key] = value

def log1_Sheet_All(sheet_name):
    global output
    log_init_cell(sheet_name)
    # for i in range(6, output[f"{sheet_name}!D3"] + 7):
    for i in range(6, 727):
        log_D_cell(sheet_name, i)
        log_C_cell(sheet_name, i)
        log_E_cell(sheet_name, i)
        log_F_cell(sheet_name, i)
        log_B_cell(sheet_name, i)
        log_J_cell(sheet_name, i)
        log_G_cell(sheet_name, i)
        log_I_cell(sheet_name, i)
        log_H_cell(sheet_name, i)
        log_K_cell(sheet_name, i)
        log_L_cell(sheet_name, i)
        log_N_cell(sheet_name, i)
        log_O_cell(sheet_name, i)
        log_P_cell(sheet_name, i)
        log_Q_cell(sheet_name, i)
        log_R_cell(sheet_name, i)
        log_S_cell(sheet_name, i)
        log_U_cell(sheet_name, i)
        log_V_cell(sheet_name, i)
        log_W_cell(sheet_name, i)
        log_X_cell(sheet_name, i)
        log_Y_cell(sheet_name, i)
        log_Z_cell(sheet_name, i)
    return output
