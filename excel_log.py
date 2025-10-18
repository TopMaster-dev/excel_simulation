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
    print(f"{key} ------ {value}")
    output[key] = value
    key = f"{sheet_name}!F3"
    value = data["入力!E18"]
    output[key] = value
    print(f"{key} ------ {value}")
    return output

def log_D_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!D{i}"
    value = i - 6
    output[key] = value
    print(f"{key} ------ {value}")

def log_C_cell(sheet_name, i):     #   log1 sheet c6 から c100 までを計算
    global output
    input_E13 = data["入力!E13"]
    Log_D = output[f"{sheet_name}!D{i}"]
    key = f"{sheet_name}!C{i}"
    value = ef.IF(input_E13 * 12 >= Log_D, 1, 0)
    output[key] = value
    print(f"{key} ------ {value}")

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
    print(f"{key} ------ {value}")

def log_F_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!F{i}"
    value = data["入力!E14"]
    output[key] = value
    print(f"{key} ------ {value}")

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
    print(f"{key} ------ {value}")

def log_J_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!J{i}"
    value = 0
    output[key] = value
    print(f"{key} ------ {value}")

def log_G_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!G{i}"
    input_log_D3 = int(output["Log1!D3"])
    input_log_F3 = int(output["Log1!F3"])
    if(i == 6):
        value = 0
    else:
        value = ef.MonthlyRepayment_FS(input_log_D3, input_log_F3, output[f"{sheet_name}!B{i}"], output[f"{sheet_name}!C{i}"], output[f"{sheet_name}!D{i}"], output[f"{sheet_name}!F{i}"], output[f"{sheet_name}!J{i}"], output[f"{sheet_name}!G{i - 1}"], output[f"{sheet_name}!K{i - 1}"])
    output[key] = value
    print(f"{key} ------ {value}")

def log_I_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!I{i}"
    if(i == 6):
        value = 0
    else:
        val = output[f"{sheet_name}!K{i - 1}"] * output[f"{sheet_name}!F{i}"]/100/12
        value = int(val)
    output[key] = value
    print(f"{key} ------ {value}")

def log_H_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!H{i}"
    if(i == 6):
        value = 0
    else:
        value = output[f"{sheet_name}!G{i}"] - output[f"{sheet_name}!I{i}"]
    output[key] = value
    print(f"{key} ------ {value}")

def log_K_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!K{i}"
    if(i == 6):
        value = data["入力!E12"]
    else:
        value = output[f"{sheet_name}!K{i - 1}"] - output[f"{sheet_name}!H{i}"] - output[f"{sheet_name}!J{i}"]
    output[key] = value
    print(f"{key} ------ {value}")

log_init_cell("Log1")
for i in range(6, 100):
    log_D_cell("Log1", i)
    log_C_cell("Log1", i)
    log_E_cell("Log1", i)
    log_F_cell("Log1", i)
    log_B_cell("Log1", i)
    log_J_cell("Log1", i)
    log_G_cell("Log1", i)
    log_I_cell("Log1", i)
    log_H_cell("Log1", i)
    log_K_cell("Log1", i)