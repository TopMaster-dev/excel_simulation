import excel_function as ef
from excel_output import output_ALL_cell
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

def graph_init_cell():
    global output
    key = "グラフ1!B2"
    value = str(data["入力!E4"]) + "様"
    output[key] = value
    key = "グラフ1!D2"
    value = data["入力!E7"]
    output[key] = value
    key = "グラフ1!F2"
    value = data["入力!E9"]
    output[key] = value
    key = "グラフ1!I2"
    value = data["入力!E10"]
    output[key] = value
    key = "グラフ1!I3"
    value = data["入力!E12"]
    output[key] = value
    key = "グラフ1!J3"
    value = str(data["入力!E13"]) + "年"
    output[key] = value
    key = "グラフ1!K3"
    value = data["入力!E14"] / 100
    output[key] = value

def graph_N_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!N{i}"
    if(sheet_name == "グラフ1"):
        value = str(output[f"出力1!C{i + 1}"]) + "(" + str(output[f"出力1!E{i + 1}"]) + "歳)"
    else:
        value = str(output[f"出力2!C{i + 1}"]) + "(" + str(output[f"出力2!E{i + 1}"]) + "歳)"
    output[key] = value

def graph_O_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!O{i}"
    if(sheet_name == "グラフ1"):
        value = output[f"出力1!AI{i + 1}"] / 10000
    else:
        value = output[f"出力2!AI{i + 1}"] / 10000
    output[key] = value

def graph_P_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!P{i}"
    if(sheet_name == "グラフ1"):
        value = output[f"出力1!AJ{i + 1}"] / 10000
    else:
        value = output[f"出力2!AJ{i + 1}"] / 10000
    output[key] = value

def graph_Q_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!Q{i}"
    if(sheet_name == "グラフ1"):
        value = output[f"出力1!AG{i + 1}"] / 10000
    else:
        value = output[f"出力2!AG{i + 1}"] / 10000
    output[key] = value

def graph_R_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!R{i}"
    value = output[f"{sheet_name}!P{i}"] + output[f"{sheet_name}!Q{i}"]
    output[key] = value

def graph_S_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!S{i}"
    value = (output[f"{sheet_name}!P{i}"] + output[f"{sheet_name}!Q{i}"]) - output[f"{sheet_name}!O{i}"]
    output[key] = value

def graph_T_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!T{i}"
    if(i == 6):
        value = 0
    else:
        value = ef.IF(ef.AND(output[f"{sheet_name}!O{i - 1}"] > 0, output[f"{sheet_name}!O{i}"] == 0), 1000, 0)
    output[key] = value

def graph_U_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!U{i}"
    value = output[f"出力1!AI{i + 1}"] / 10000
    output[key] = value

def graph_V_cell(sheet_name, i):
    global output
    key = f"{sheet_name}!V{i}"
    if(i == 6):
        value = ""
    else:
        value = ef.IF(ef.AND(output[f"{sheet_name}!U{i - 1}"] > 0, output[f"{sheet_name}!U{i}"] == 0), 1000, 0)
    output[key] = value

def graph_All_one_cell(sheet_name):
    for i in range(6, 66):
        graph_N_cell(sheet_name, i)
        graph_O_cell(sheet_name, i)
        graph_P_cell(sheet_name, i)
        graph_Q_cell(sheet_name, i)
        graph_R_cell(sheet_name, i)
        graph_S_cell(sheet_name, i)
        graph_T_cell(sheet_name, i)
        if(sheet_name == "グラフ2"):
            graph_U_cell(sheet_name, i)
            graph_V_cell(sheet_name, i)
    return output

def graph_All_cell():
    global output
    output = output_ALL_cell()
    graph_init_cell()
    graph_All_one_cell("グラフ1")
    graph_All_one_cell("グラフ2")
    return output

print(graph_All_cell())