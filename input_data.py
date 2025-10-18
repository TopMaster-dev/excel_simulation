import json

input_data = {}
def get_input_data():
    with open("input.json", "r") as f:
        input_data = json.load(f)
    return input_data

input_data = get_input_data()