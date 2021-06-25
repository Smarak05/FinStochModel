import pandas as pd
from pickle import load

def load_data():
    with open('parsed_data', "rb") as f:
        data = load(f)

    print(data)
    return data

def convert_data(data):
    converted_data = dict()

    for name, company in data.items():
        for group, group_data in company.items():
            if group not in converted_data:
                converted_data[group] = dict()

            converted_data[group][name] = group_data
    
    converted_data['Profitability'] = converted_data['Profitabiliy']
    del converted_data['Profitabiliy']

    return converted_data

def saveToExcel(data):
    for group, group_data in data.items():
        with pd.ExcelWriter(f'./excelData/{group}.xlsx') as writer:
            for name, company in group_data.items():
                name = name.replace(":", "-")
                name = name.replace("/", " ")
                company.to_excel(writer, sheet_name=name, index=False)

if __name__ == "__main__":
    data = load_data()
    # data = convert_data(data)
    # saveToExcel(data)
